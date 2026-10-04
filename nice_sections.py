"""
Committee-section grounding for app3.

Named nice_sections, NOT nice_chapters: fetch_nice_chapters.py writes its HTML
into a directory called nice_chapters/, and on some Python versions that
directory shadows a module of the same name as an empty namespace package.

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

# Roughly 20k tokens of committee text. Generous but bounded.
DEFAULT_BUDGET = 80_000
MAX_SECTIONS_PER_APPRAISAL = 6

# Spreading the budget across every retrieved appraisal truncates nearly every
# section mid-sentence, and a quote that stops before the committee's
# conclusion is worse than no quote. Sections go to the most relevant
# appraisals only; the rest are still present as their Excel rows.
MAX_APPRAISALS_WITH_SECTIONS = 14

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
            # The element id doubles as the page anchor, so a citation can
            # deep-link to the exact paragraph on nice.org.uk.
            "anchor": md.get("anchor", ""),
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
            # Prefer dropping a section to cutting it mid-sentence: a quote
            # that stops before the committee's conclusion misleads. Only the
            # genuinely huge sections get trimmed, and then at a sentence end.
            if room < 1500:
                continue
            cut = text[:room]
            stop = max(cut.rfind(". "), cut.rfind(".\n"))
            text = (cut[:stop + 1] if stop > room * 0.5
                    else cut.rsplit(" ", 1)[0]) + " [section continues]"
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

    n_indexed = len(tas)

    # Rank by how well an appraisal's best section answers the question, then
    # keep only the strongest. Complete quotes from a dozen appraisals beat
    # truncated ones from sixty.
    terms = [t for t in re.findall(r"[a-z0-9]+", (question or "").lower())
             if len(t) > 2 and t not in _STOP]
    if terms and len(tas) > MAX_APPRAISALS_WITH_SECTIONS:
        def best(ta):
            return max((_score(s, terms) for s in idx[ta]), default=0)
        tas = sorted(tas, key=best, reverse=True)[:MAX_APPRAISALS_WITH_SECTIONS]
    else:
        tas = tas[:MAX_APPRAISALS_WITH_SECTIONS]

    per = max(3000, budget // max(1, len(tas)))
    blocks = []
    for ta in tas:
        block = format_for_context(ta, question, char_budget=per)
        if block:
            blocks.append(f"[TA{ta}]\n{block}")

    if not blocks:
        return context

    n_with, n_total = coverage(appraisal_ids)
    shown = len(blocks)
    scope = (f"{shown} of the {n_total} appraisals above"
             if shown >= n_indexed else
             f"the {shown} most relevant of the {n_with} indexed appraisals above "
             f"(of {n_total} retrieved)")
    header = (
        "\n\nCommittee discussion sections, for "
        f"{scope}. These are the committee's own published "
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


# ---------------------------------------------------------------- source panel
# Everything below supports showing a cited section beside the answer. Kept
# free of Streamlit so it stays testable; the app does the rendering.

CITATION_RE = re.compile(
    r"\[TA\s*(\d+)[,\s]*(?:§|sections?|s\.)\s*(\d+\.\d+)\]", re.IGNORECASE)


def parse_citations(text):
    """Citations in an answer, in order of appearance, deduplicated.

    Accepts the forms the model actually produces: [TA993 section 3.13],
    [TA993 §3.13], [TA993, section 3.13].
    """
    seen, out = set(), []
    for ta, section in CITATION_RE.findall(text or ""):
        ta = ta.lstrip("0") or "0"
        if (ta, section) not in seen:
            seen.add((ta, section))
            out.append((ta, section))
    return out


def get_section(appraisal_id, section):
    """One specific section, or None."""
    ta = _norm_ta(appraisal_id)
    if not ta:
        return None
    for s in load_index().get(ta, []):
        if s["section"] == section:
            return s
    return None


def neighbours(appraisal_id, section, span=1):
    """Adjacent sections, so a quote can be read in context."""
    ta = _norm_ta(appraisal_id)
    if not ta or "." not in section:
        return []
    chapter, number = section.split(".", 1)
    if not number.isdigit():
        return []
    out = []
    for delta in range(-span, span + 1):
        if delta == 0:
            continue
        n = int(number) + delta
        if n < 1:
            continue
        found = get_section(ta, f"{chapter}.{n}")
        if found:
            out.append(found)
    return out


def quoted_spans(answer):
    """Every quoted passage in an answer, straight or curly quotes."""
    return re.findall(r"[“\"]([^”\"]{20,600})[”\"]", answer or "")


def mark_quotes(text, answer, marker="**"):
    """Wrap any part of `text` the answer quoted.

    Quotes are routinely elided ("too much uncertainty ... preferred ICER"),
    so each fragment between ellipses is located separately - extending from
    the start by the quote's length would end mid-word.
    """
    fragments = []
    for quote in quoted_spans(answer):
        for frag in re.split(r"…|\.\.\.|\[[^\]]*\]", quote):
            frag = frag.strip(" .,;:…")
            if len(frag) >= 25:
                fragments.append(frag)
    if not fragments:
        return text

    spans, low = [], text.lower()
    for frag in fragments:
        i = low.find(frag.lower())
        if i >= 0:
            spans.append((i, i + len(frag)))
            continue
        loose = re.sub(r"\\\s+", r"\\s+", re.escape(frag.lower()))
        m = re.search(loose, low)
        if m:
            spans.append(m.span())
    if not spans:
        return text

    spans.sort()
    merged = [spans[0]]
    for a, b in spans[1:]:
        if a <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], b))
        else:
            merged.append((a, b))

    out = text
    for a, b in reversed(merged):
        out = out[:a] + marker + out[a:b] + marker + out[b:]
    return out


def section_url(appraisal_id, section, anchor=None):
    """Deep link to the paragraph on nice.org.uk."""
    ta = _norm_ta(appraisal_id)
    chapter = section.split(".")[0] if "." in section else "3"
    frag = f"#{anchor}" if anchor else ""
    return (f"https://www.nice.org.uk/guidance/ta{ta}/chapter/"
            f"{chapter}-Committee-discussion{frag}")
