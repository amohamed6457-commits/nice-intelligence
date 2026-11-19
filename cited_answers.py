"""
Answers that cite the passage they came from, with its page.

The chat in app3 is already grounded — it answers only from the records the
user selected — but it is grounded on the SUMMARY fields of each record and
cites by appraisal id. So an answer can say "the G-BA found no added benefit
[D-1313]" and the reader still has to go and find where in a 20-page
resolution that was decided.

This module grounds the answer on the page-cited passages instead, and makes
the model mark every claim with the passage it used. The app then shows each
cited passage underneath the answer with a link that opens the official PDF
at that page.

The part that matters
---------------------
A language model asked to cite will sometimes cite something that is not
there. So nothing the model says about its sources is taken on trust:

  * Every citation marker is checked against the passages that were actually
    supplied. A marker naming a passage that was not in the context is
    reported as unverified, not rendered as a source.
  * Every verbatim quotation the model gives is checked, character by
    character after whitespace normalisation, against the passage it is
    attributed to. A quote that is not in its passage is flagged in the
    answer.

That check is the difference between a tool a reviewer can rely on and a
wrapper that sounds confident. An unverifiable claim is shown as
unverifiable rather than quietly dropped, because silently removing a
citation would leave a confident sentence with nothing behind it.

Retrieval is deliberately plain — IDF-weighted term overlap with a phrase
bonus, no embeddings. It is explainable, needs no extra dependency, and the
candidate pool is already narrowed to the records on screen.
"""

import math
import re

import pandas as pd
import streamlit as st

_PASSAGES = None          # normalised DataFrame, or None before attach()
_IDF = None               # token -> inverse document frequency

MAX_PASSAGES = 12         # sent to the model per question
MAX_PASSAGE_CHARS = 1800  # per passage, so twelve of them fit comfortably

_TOKEN = re.compile(r"[A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß\-]{2,}|\d[\d.,]*")
_CITE = re.compile(r"\[(S\d{1,2})\]")
_QUOTED = re.compile(r"[\"“]([^\"”]{12,400})[\"”]")

# Words that carry no retrieval signal in this corpus. Short, because an
# over-eager stoplist loses real terms: 'benefit' and 'Nutzen' matter here.
_STOP = {
    "the", "and", "for", "with", "that", "this", "from", "was", "were", "has",
    "have", "had", "not", "are", "but", "its", "their", "which", "what", "how",
    "why", "did", "does", "any", "all", "can", "will", "would", "there",
    "der", "die", "das", "und", "von", "mit", "ist", "sind", "den", "dem",
    "des", "ein", "eine", "nicht", "auch", "wird", "werden", "bei", "für",
}

CITED_ANSWER_RULES = """

--- Citing the source passages ---

The numbered passages below are the ONLY evidence you may use for claims about
what an agency said, decided, measured or recommended.

7. After every sentence that makes such a claim, add the marker of the passage
   it came from, like [S3]. Several passages for one sentence: [S3][S7].
8. Use ONLY markers that appear in the list below. Never invent a marker, and
   never cite a passage you were not given.
9. When a claim rests on specific wording, quote that wording verbatim in
   double quotes and put the marker straight after it. Copy it exactly from the
   passage, including the original German or French — do not translate inside
   the quotation marks. A translation or gloss goes outside them.
10. If the passages do not answer the question, say exactly that. An answer
    with no marker is read as having no source, so never attach a marker to a
    claim the passage does not make in order to appear supported.
"""


def _text(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    s = str(value).strip()
    return s or None


# An English question scores almost nothing against a German resolution:
# only the drug name matches, so every passage ties and retrieval is no
# better than picking at random. The HTA vocabulary is small and formulaic
# in all three languages, so each English concept is expanded to the words
# the documents actually use. This is retrieval only — nothing here is shown
# to the reader or sent as a translation.
_EXPAND = {
    "benefit": ["zusatznutzen", "nutzen", "bénéfice"],
    "added": ["zusatznutzen", "zusätzlichen"],
    "comparator": ["vergleichstherapie", "zweckmäßigen", "comparateur"],
    "therapy": ["therapie", "vergleichstherapie", "traitement"],
    "treatment": ["behandlung", "therapie", "traitement"],
    "cost": ["therapiekosten", "kosten", "coût", "cout"],
    "costs": ["therapiekosten", "kosten", "coûts"],
    "price": ["preis", "prix"],
    "survival": ["überlebenszeit", "gesamtüberleben", "survie"],
    "mortality": ["mortalität", "gesamtüberleben", "mortalité"],
    "morbidity": ["morbidität", "morbidité"],
    "progression": ["progressionsfreies", "progression"],
    "quality": ["lebensqualität", "qualité"],
    "life": ["lebensqualität", "vie"],
    "adverse": ["nebenwirkungen", "unerwünschte", "indésirables"],
    "events": ["ereignisse", "nebenwirkungen", "événements"],
    "safety": ["sicherheit", "sécurité"],
    "patients": ["patienten", "patientinnen", "patients"],
    "population": ["patientenzahl", "population"],
    "proven": ["belegt", "nachgewiesen"],
    "evidence": ["beleg", "hinweis", "anhaltspunkt", "preuve"],
    "extent": ["ausmaß"],
    "probability": ["wahrscheinlichkeit", "probabilité"],
    "subgroup": ["subgruppe", "teilpopulation", "sous-groupe"],
    "indication": ["anwendungsgebiet", "indikation"],
    "effectiveness": ["wirksamkeit", "efficacité"],
    "efficiency": ["effizienz", "efficience"],
    "reimbursement": ["erstattung", "remboursement"],
    "recommendation": ["empfehlung", "recommandation"],
    "uncertainty": ["unsicherheit", "incertitude"],
}


def _tokens(text):
    return [t.lower() for t in _TOKEN.findall(str(text or ""))
            if t.lower() not in _STOP]


def _expand(terms):
    """Query terms plus the German and French words meaning the same thing."""
    out = list(terms)
    for t in terms:
        out.extend(_EXPAND.get(t, ()))
    return out


def _frame(df, agency, text_col, section_col=None, extra=None):
    """Normalise one sheet into the common passage shape."""
    if df is None or not len(df) or text_col not in df.columns:
        return pd.DataFrame()
    out = pd.DataFrame({
        "appraisal_id": df["appraisal_id"].astype(str).str.strip(),
        "agency": agency,
        "section": (df[section_col].astype(str) if section_col in df.columns
                    else pd.Series([""] * len(df), index=df.index)),
        "page": pd.to_numeric(df.get("document_page"), errors="coerce"),
        "url": df.get("source_url"),
        "text": df[text_col].astype(str),
    })
    if extra is not None and extra in df.columns:
        out["value"] = df[extra]
    else:
        out["value"] = None
    out = out[out["text"].str.strip().astype(bool)]
    return out


def attach(gba_passages=None, has_pages=None, economic_findings=None):
    """
    Build the searchable passage pool once per run.

    Safe to call repeatedly and safe with missing sheets — on the short
    build of the workbook the pool is simply empty and the feature reports
    itself as unavailable rather than failing.
    """
    global _PASSAGES, _IDF
    parts = [
        _frame(gba_passages, "G-BA", "source_text", "section"),
        _frame(has_pages, "HAS", "source_text", "section"),
        _frame(economic_findings, "HAS", "source_context", "finding_type", "value"),
    ]
    parts = [p for p in parts if len(p)]
    if not parts:
        _PASSAGES, _IDF = pd.DataFrame(), {}
        return
    pool = pd.concat(parts, ignore_index=True)
    pool["tokens"] = pool["text"].map(lambda t: set(_tokens(t)))

    n = len(pool)
    counts = {}
    for toks in pool["tokens"]:
        for t in toks:
            counts[t] = counts.get(t, 0) + 1
    _IDF = {t: math.log(1 + n / (1 + c)) for t, c in counts.items()}
    _PASSAGES = pool


def is_attached():
    return _PASSAGES is not None


def coverage():
    if _PASSAGES is None or not len(_PASSAGES):
        return {"passages": 0, "appraisals": 0}
    return {"passages": len(_PASSAGES),
            "appraisals": _PASSAGES["appraisal_id"].nunique()}


def select_passages(question, appraisal_ids=None, limit=MAX_PASSAGES):
    """
    The passages most likely to answer the question, highest score first.

    Restricted to appraisal_ids when given — normally the records the user
    has on screen, so the answer cannot quietly reach outside the selection
    the rest of the page is about.
    """
    if _PASSAGES is None or not len(_PASSAGES):
        return []
    pool = _PASSAGES
    if appraisal_ids:
        wanted = {str(a).strip() for a in appraisal_ids}
        pool = pool[pool["appraisal_id"].isin(wanted)]
        if not len(pool):
            return []

    asked = _tokens(question)
    if not asked:
        return []
    terms = _expand(asked)
    weights = {}
    for t in terms:
        # A word the user actually typed counts for more than one this
        # module added on their behalf.
        boost = 1.0 if t in asked else 0.6
        weights[t] = weights.get(t, 0) + _IDF.get(t, 1.0) * boost

    phrase = " ".join(terms[:6]).lower()
    scored = []
    for pos, row in enumerate(pool.itertuples(index=False)):
        toks = row.tokens
        score = sum(w for t, w in weights.items() if t in toks)
        if score <= 0:
            continue
        # A passage containing the question's wording outright beats one that
        # merely shares vocabulary with it.
        if len(phrase) > 12 and phrase in row.text.lower():
            score *= 1.5
        scored.append((score, pos, row))

    scored.sort(key=lambda s: (-s[0], s[1]))
    out = []
    for rank, (score, _pos, row) in enumerate(scored[:limit], 1):
        text = row.text.strip()
        out.append({
            "sid": f"S{rank}",
            "appraisal_id": row.appraisal_id,
            "agency": row.agency,
            "section": _text(row.section) or "",
            "page": int(row.page) if pd.notna(row.page) else None,
            "url": _text(row.url),
            "value": _text(row.value),
            "text": text[:MAX_PASSAGE_CHARS],
            "truncated": len(text) > MAX_PASSAGE_CHARS,
            "score": round(float(score), 2),
        })
    return out


def build_context(passages):
    """The passage list exactly as the model will see it."""
    if not passages:
        return "No source passages matched this question."
    blocks = []
    for p in passages:
        head = f"[{p['sid']}] {p['agency']} {p['appraisal_id']}"
        if p["section"]:
            head += f" — {p['section']}"
        if p["page"]:
            head += f" — page {p['page']}"
        if p["value"]:
            head += f" — reported value: {p['value']}"
        body = p["text"] + ("…" if p["truncated"] else "")
        blocks.append(f"{head}\n{body}")
    return ("--- Source passages (cite these by marker) ---\n\n"
            + "\n\n".join(blocks))


def _normalise(text):
    return re.sub(r"\s+", " ", str(text or "")).strip().lower()


def audit(answer, passages):
    """
    Check the model's citations against what it was actually given.

    Returns a dict with:
      used      — markers that name a supplied passage
      unknown   — markers naming a passage that was never supplied
      bad_quote — [(quote, sid)] where the quoted wording is not in that
                  passage
      uncited   — True when the answer makes claims but cites nothing
    """
    known = {p["sid"]: p for p in passages}
    found = _CITE.findall(answer or "")
    used = [s for s in dict.fromkeys(found) if s in known]
    unknown = [s for s in dict.fromkeys(found) if s not in known]

    bad_quote = []
    for match in _QUOTED.finditer(answer or ""):
        quote = match.group(1)
        # The marker that follows the quotation owns it; failing that, the
        # one just before.
        after = _CITE.search(answer, match.end(), match.end() + 24)
        before = None
        if not after:
            prefix = answer[max(0, match.start() - 24):match.start()]
            marks = _CITE.findall(prefix)
            before = marks[-1] if marks else None
        sid = after.group(1) if after else before
        if sid and sid in known:
            if _normalise(quote) not in _normalise(known[sid]["text"]):
                bad_quote.append((quote, sid))

    substantive = len(_normalise(answer)) > 120
    return {"used": used, "unknown": unknown, "bad_quote": bad_quote,
            "uncited": substantive and not used}


def page_url(url, page):
    """PDF open parameter, preserving any fragment already on the URL."""
    u = _text(url)
    if not u or not page:
        return u
    try:
        n = int(page)
    except (TypeError, ValueError):
        return u
    if n < 1:
        return u
    return f"{u}&page={n}" if "#" in u else f"{u}#page={n}"


def _md_safe(text):
    return str(text).replace("$", r"\$")


def render_sources(answer, passages, show_unused=False):
    """
    The cited passages under the answer, plus any integrity warnings.

    Only passages the answer actually cited are shown by default: a list of
    twelve candidates under a two-line answer tells the reader nothing about
    which one the claim rests on.
    """
    report = audit(answer, passages)
    known = {p["sid"]: p for p in passages}

    if report["unknown"]:
        st.error(
            "The answer cited " + ", ".join(report["unknown"]) +
            ", which was not among the passages it was given. Treat the claims "
            "carrying those markers as unsupported.")
    for quote, sid in report["bad_quote"]:
        st.warning(f"The quotation attributed to {sid} does not appear in that "
                   f"passage: “{quote[:120]}…”" if len(quote) > 120 else
                   f"The quotation attributed to {sid} does not appear in that "
                   f"passage: “{quote}”")
    if report["uncited"]:
        st.warning("The answer cited no passage. Anything in it that sounds "
                   "like a finding is not backed by the indexed sources.")

    sids = report["used"] or ([p["sid"] for p in passages] if show_unused else [])
    if not sids:
        return report

    st.markdown(f"**Sources ({len(sids)})**")
    for sid in sids:
        p = known[sid]
        head = f"{sid} · {p['agency']} {p['appraisal_id']}"
        if p["section"]:
            head += f" · {p['section']}"
        if p["page"]:
            head += f" · page {p['page']}"
        with st.expander(head):
            if p["value"]:
                st.markdown(f"**Reported value: {_md_safe(p['value'])}**")
            body = p["text"]
            st.markdown("> " + _md_safe(body).replace("\n", "  \n> "))
            link = page_url(p["url"], p["page"])
            if link:
                st.caption(f"[Open the source at page {p['page']} ↗]({link})")
    return report
