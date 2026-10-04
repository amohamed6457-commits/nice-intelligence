"""
Committee-section grounding for app3.

Drop this file next to app3.py. It adds NICE committee discussion sections to
the chat context, so answers can quote the committee's own words and cite
[TA993 section 3.13] instead of paraphrasing an Excel column.

Design notes:
  * app3's chat is a single grounded call, not a tool-use loop, and promises
    answers use only the retrieved results. So sections are selected from the
    appraisals ALREADY in the result set - this deepens the existing scope
    rather than widening it.
  * Everything degrades quietly. No index, no sections for an appraisal, a
    corrupt file: the context is returned unchanged and the app behaves exactly
    as it does today.
  * A character budget is enforced, because 60 appraisals x ~20 sections would
    otherwise be well over a million characters of prompt.

Build the index first:
    python parse_nice_chapters.py
    python build_chapter_index.py --keyword
"""

import json
import os
import re

_HERE = os.path.dirname(os.path.abspath(__file__))

# Compressed first. The index has to be committed alongside the app for a
# deployed instance to see it, and gzip takes ~9,700 sections from roughly
# 31 MB to 4 MB. Plain .json is still read so an older build keeps working.
INDEX_PATHS = [
    os.path.join("hta_index", "keyword_index.json.gz"),
    os.path.join(_HERE, "hta_index", "keyword_index.json.gz"),
    os.path.join("hta_index", "keyword_index.json"),
    os.path.join(_HERE, "hta_index", "keyword_index.json"),
]

# Roughly 15k tokens of committee text. Generous but bounded.
DEFAULT_BUDGET = 60_000
MAX_SECTIONS_PER_APPRAISAL = 6

_CACHE = {"loaded": False, "by_ta": {}, "error": None}

_STOP = {"the", "and", "for", "was", "were", "with", "that", "this", "did",
         "does", "what", "why", "how", "any", "are", "its", "has", "have",
         "from", "about", "which", "them", "they", "not", "but"}


def _norm_ta(value):
    """'TA993', 'ta993', 993, '993' -> '993'. None if not a TA id."""
    s = str(value).strip().upper()
    m = re.search(r"(\d+)", s.replace("TA", ""))
    return m.group(1).lstrip("0") or "0" if m else None


def load_index():
    """Load once. Returns {ta: [section, ...]}; empty dict if unavailable."""
    if _CACHE["loaded"]:
        return _CACHE["by_ta"]
    _CACHE["loaded"] = True

    path = next((p for p in INDEX_PATHS if os.path.exists(p)), None)
    if not path:
        _CACHE["error"] = "no index file"
        return {}

    try:
        if path.endswith(".gz"):
            import gzip
            with gzip.open(path, "rt", encoding="utf-8") as f:
                records = json.load(f)
        else:
            with open(path, encoding="utf-8") as f:
                records = json.load(f)
    except (json.JSONDecodeError, OSError, EOFError) as e:
        _CACHE["error"] = f"could not read index: {e}"
        return {}

    by_ta = {}
    for rec in records:
        md = rec.get("metadata", {})
        ta = _norm_ta(md.get("ta", ""))
        if not ta:
            continue
        by_ta.setdefault(ta, []).append({
            "section": md.get("section", ""),
            "headline": md.get("headline", ""),
            "theme": md.get("theme", ""),
            "tier": md.get("source_tier", "chapter"),
            "url": md.get("url", ""),
            "text": rec.get("text", ""),
        })
    for ta in by_ta:
        by_ta[ta].sort(key=_section_sort_key)

    _CACHE["by_ta"] = by_ta
    return by_ta


def _section_sort_key(s):
    try:
        a, b = s["section"].split(".")
        return (int(a), int(b))
    except (ValueError, AttributeError):
        return (99, 99)


def available():
    """True if any sections are indexed."""
    return bool(load_index())


def status():
    """One line for the UI: what grounding is actually available."""
    idx = load_index()
    if not idx:
        return f"Committee sections: not indexed ({_CACHE['error'] or 'unavailable'})"
    n = sum(len(v) for v in idx.values())
    return f"Committee sections: {n:,} across {len(idx)} appraisals"


def coverage(appraisal_ids):
    """(n_with_sections, n_total) for a result set - honest UI labelling."""
    idx = load_index()
    if not idx:
        return 0, len(list(appraisal_ids))
    tas = [_norm_ta(a) for a in appraisal_ids]
    tas = [t for t in tas if t]
    return sum(1 for t in tas if t in idx), len(tas)


def _score(section, terms):
    """Relevance of one section to the question. Headline counts double:
    NICE writes the committee's conclusion there."""
    if not terms:
        return 0
    head = (section["headline"] + " " + section["theme"]).lower()
    body = section["text"].lower()
    return sum(2 * head.count(t) + body.count(t) for t in terms)


def sections_for(appraisal_id, question="", limit=MAX_SECTIONS_PER_APPRAISAL):
    """Most relevant indexed sections for one appraisal, best first.

    With no question, returns the first `limit` sections in document order,
    which is a reasonable default since NICE orders them by argument.
    """
    ta = _norm_ta(appraisal_id)
    if not ta:
        return []
    found = load_index().get(ta, [])
    if not found:
        return []
    terms = [t for t in re.findall(r"[a-z0-9]+", (question or "").lower())
             if len(t) > 2 and t not in _STOP]
    if not terms:
        return found[:limit]
    ranked = sorted(found, key=lambda s: -_score(s, terms))
    ranked = [s for s in ranked if _score(s, terms) > 0][:limit]
    return sorted(ranked, key=_section_sort_key)


def format_for_context(appraisal_id, question="", char_budget=12_000,
                       limit=MAX_SECTIONS_PER_APPRAISAL):
    """Committee sections for one appraisal, as lines to append to its block.

    Returns "" when nothing is indexed, so the caller can append unconditionally.
    """
    secs = sections_for(appraisal_id, question, limit)
    if not secs:
        return ""

    ta = _norm_ta(appraisal_id)
    preamble = ("  Committee discussion, quotable and citable as "
                f"[TA{ta} section N.N]:\n")
    # +MAX_SECTIONS reserves the newline joining each section line.
    lines, used = [], len(preamble) + MAX_SECTIONS_PER_APPRAISAL
    for s in secs:
        head = f"  [TA{ta} section {s['section']}]"
        if s["headline"]:
            head += f" {s['headline']}"
        # Budget covers the header too, or the block overshoots by ~100 chars
        # per section, which compounds across a 60-appraisal result set.
        room = char_budget - used - len(head) - 6
        if room < 300:
            break
        text = s["text"]
        if len(text) > room:
            text = text[:room].rsplit(" ", 1)[0] + " [truncated]"
        lines.append(f"{head}\n    {text}")
        used += len(head) + len(text) + 6
    if not lines:
        return ""
    return preamble + "\n".join(lines)


def augment_context(context, appraisal_ids, question="", budget=DEFAULT_BUDGET):
    """Append a committee-sections block to an existing app3 chat context.

    Sections are spread evenly across the appraisals in the result set so one
    verbose appraisal cannot consume the whole budget.
    """
    idx = load_index()
    if not idx:
        return context

    tas = []
    for a in appraisal_ids:
        t = _norm_ta(a)
        if t and t in idx and t not in tas:
            tas.append(t)
    if not tas:
        return context

    per = max(1200, budget // max(1, len(tas)))
    blocks = []
    for ta in tas:
        block = format_for_context(ta, question, char_budget=per)
        if block:
            blocks.append(f"[TA{ta}]\n{block}")

    if not blocks:
        return context

    n_with, n_total = coverage(appraisal_ids)
    header = (
        "\n\nCommittee discussion sections, for the appraisals above that have "
        f"them ({n_with} of {n_total}). These are the committee's own published "
        "words. Quote them directly and cite as [TA<number> section <N.N>]. "
        "For appraisals without sections here, cite the appraisal alone and do "
        "not invent a section number."
    )
    return context + header + "\n\n" + "\n\n".join(blocks)


# Appended to app3's existing chat system prompts.
CITATION_RULES = """

When committee discussion sections are supplied:
- Quote the committee's own words for any claim about why something was
  accepted or rejected, and cite as [TA707 section 3.14].
- Only cite a section number that appears in the supplied context. If an
  appraisal has no sections here, cite the appraisal alone.
- Keep the company's argument, the EAG/ERG's critique and the committee's
  conclusion distinct. They are routinely confused and the difference matters.
"""
