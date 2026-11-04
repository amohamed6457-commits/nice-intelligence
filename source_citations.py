"""
Page-level source citation for G-BA and HAS records.

The workbook carries three sheets the app has never read:

    GBA_Decision_Passages   3,911 rows / 982 resolutions — the text the G-BA
                            published, with the page it sits on
    HAS_Economic_Pages      2,586 rows / 229 opinions — page-by-page text of
                            the avis d'efficience
    Economic_Findings       5,803 rows / 225 opinions — a value, a unit, a
                            page and the passage it was read from

Every row carries a source_url and an integer document_page, so each passage
can be shown inline with a link that opens the official PDF scrolled to the
page it came from. That is the whole point of this module: a reader should
never have to take the app's word for a finding.

Three things this module is careful about.

**Parts are not pages.** A section's text can run across pages, recorded as
part 1, 2, 3. In every one of the 101 multi-part groups the parts sit on
DIFFERENT pages, so each part gets its own link. Joining them under one page
number would cite the wrong page for most of the text.

**Not every "finding" is an ICER.** Of 5,803 economic findings only 3,054 are
a reported ICER. 1,720 are "Monetary ratio; verify interpretation" and 406 are
a willingness-to-pay threshold that is explicitly NOT an ICER. Presenting
those as cost-effectiveness results would be wrong, so the finding type is
shown as a label on every value and the non-ICER types are visually separated.

**The workbook may not have the sheets.** Two files have shipped under the
name NICE_v14_updated_2026-09-28.xlsx — one with 6 sheets, one with 15. On the
6-sheet file every frame here arrives empty and the panel renders nothing at
all, rather than raising.

Usage, matching the nice_attributes convention:

    source_citations.attach(gba_passages, has_pages, economic_findings)
    source_citations.render_panel("G-BA", appraisal_id)
"""

import pandas as pd
import streamlit as st

# ── state ────────────────────────────────────────────────────────────────────

_GBA = None        # appraisal_id -> DataFrame of passages
_HAS_PAGES = None
_FINDINGS = None

# Section order for G-BA. The sheet's four section labels, most decision-
# relevant first; anything unrecognised sorts last but still renders.
_GBA_SECTION_ORDER = [
    "Decision / subgroup findings",
    "Patient population count",
    "Treatment costs",
    "Quality assured use",
]

# Findings that are a cost-effectiveness result, as opposed to a threshold or
# an unverified ratio. Drives which block a value lands in.
_ICER_TYPES = {"Reported incremental cost-effectiveness ratio"}

# German and French terms worth glossing for an English-speaking reader. Only
# the ones that change the meaning of a decision — not a general dictionary.
GLOSSARY = {
    "zusatznutzen": "additional benefit — the single thing a G-BA assessment decides",
    "nicht belegt": "not proven — the company's evidence did not establish it",
    "zweckmäßige vergleichstherapie": "the comparator therapy the G-BA mandates",
    "anhaltspunkt": "weakest evidence certainty: an indication of benefit",
    "hinweis": "middle evidence certainty: a hint of benefit",
    "beleg": "strongest evidence certainty: proof of benefit",
    "beschluss": "the formal resolution document",
    "service médical rendu": "actual clinical benefit — sets the reimbursement rate",
    "amélioration du service médical rendu": "improvement in clinical benefit — drives price",
    "avis d'efficience": "efficiency opinion — HAS's economic assessment",
    "commission de la transparence": "the HAS committee that issues reimbursement opinions",
}

_PASSAGE_CHARS = 700      # before the "show the rest" expander

# Lowercase words that legitimately begin a passage, so clean_start() leaves
# them alone. Everything else starting lowercase is treated as a fragment.
_CLEAN_OPENERS = {
    "le", "la", "les", "de", "des", "du", "un", "une", "et", "en", "au", "aux",
    "pour", "par", "sur", "dans", "avec", "est", "ce", "ces", "cette", "il",
    "elle", "qui", "que", "ratio", "coût", "cout", "selon", "ainsi", "son",
    "ses", "lors", "dont", "après", "apres", "soit", "der", "die", "das",
    "ein", "eine", "und", "von", "mit", "für", "fur", "im", "zur", "zum",
    "bei", "nach", "ist", "sind", "wird", "werden", "nicht", "auch",
}


# ── helpers ──────────────────────────────────────────────────────────────────

def _text(value):
    """Stripped string, or None for blank / NaN."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    s = str(value).strip()
    return s or None


def _md_safe(text):
    """Stop '$' in source text rendering as LaTeX. Same guard app3 uses."""
    return str(text).replace("$", r"\$")


def clean_start(text):
    """
    Drop a leading mid-word fragment.

    The HAS passages are fixed-width character windows, not sentence-aligned
    extracts, so 42% of them (2,436 of 5,803) open mid-word — 'emporel vie
    entière', 'gies conduisent à', 'obale, le coût'. The text is correct; it
    just starts a few characters early. Dropping the partial first token and
    marking the passage as a continuation costs at most one real word and
    stops every other quotation looking corrupted.

    A passage starting with a capital, a digit or a common French function
    word is left alone — those begin cleanly.
    """
    s = str(text).strip()
    if not s or not s[0].islower():
        return s, False
    head = s.split(" ", 1)[0].strip(".,;:()«»")
    if head.lower() in _CLEAN_OPENERS or len(head) > 20:
        return s, False
    rest = s.split(" ", 1)
    if len(rest) < 2:
        return s, False
    return rest[1].lstrip(), True


def trim(text, limit=_PASSAGE_CHARS):
    """
    Cut at a sentence end if there is one in the last third, otherwise at a
    word boundary. Never mid-word — a passage ending 'were not acce' reads as
    a bug even when the full text is one click away.
    """
    s = str(text).strip()
    if len(s) <= limit:
        return s, False
    window = s[:limit]
    for stop in (". ", ".\n", "; ", " – ", " — "):
        cut = window.rfind(stop)
        if cut > limit * 0.6:
            return window[:cut + 1].strip(), True
    return window.rsplit(" ", 1)[0].rstrip(",;:") + "…", True


def page_url(url, page):
    """
    Append the PDF open parameter so the viewer lands on the right page.

    #page=N is honoured by Chrome, Edge, Firefox and Acrobat. #search= is NOT
    (Acrobat and Firefox only), which is why this module highlights nothing
    and instead shows the passage inline next to the link — the reader
    compares the two themselves.

    A URL that already carries a fragment keeps it; the page parameter is
    appended with '&' so neither is lost.
    """
    u = _text(url)
    if not u:
        return None
    try:
        n = int(page)
    except (TypeError, ValueError):
        return u
    if n < 1:
        return u
    return f"{u}&page={n}" if "#" in u else f"{u}#page={n}"


def _glossary_hits(*texts):
    """Glossary entries whose term appears in any of the given passages."""
    blob = " ".join(str(t) for t in texts if t).lower()
    return [(term, gloss) for term, gloss in GLOSSARY.items() if term in blob]


def _by_appraisal(frame, id_col="appraisal_id"):
    """dict of appraisal_id -> rows. Empty dict for a missing/empty sheet."""
    if frame is None or not len(frame) or id_col not in frame.columns:
        return {}
    f = frame.copy()
    f[id_col] = f[id_col].astype(str).str.strip()
    return {key: group for key, group in f.groupby(id_col, sort=False)}


# ── wiring ───────────────────────────────────────────────────────────────────

def attach(gba_passages=None, has_pages=None, economic_findings=None):
    """Index the three sheets once per run. Safe to call repeatedly."""
    global _GBA, _HAS_PAGES, _FINDINGS
    _GBA = _by_appraisal(gba_passages)
    _HAS_PAGES = _by_appraisal(has_pages)
    _FINDINGS = _by_appraisal(economic_findings)


def is_attached():
    return _GBA is not None


def coverage():
    """Counts for the methodology expander. Zeroes before attach()."""
    return {
        "gba_passages": sum(len(v) for v in (_GBA or {}).values()),
        "gba_resolutions": len(_GBA or {}),
        "has_pages": sum(len(v) for v in (_HAS_PAGES or {}).values()),
        "has_opinions": len(_HAS_PAGES or {}),
        "findings": sum(len(v) for v in (_FINDINGS or {}).values()),
        "finding_opinions": len(_FINDINGS or {}),
    }


def has_sources(agency, appraisal_id):
    """True if anything would render — lets a caller skip an empty header."""
    key = _text(appraisal_id)
    if not key:
        return False
    if agency == "G-BA":
        return key in (_GBA or {})
    if agency == "HAS":
        return key in (_FINDINGS or {}) or key in (_HAS_PAGES or {})
    return False


# ── rendering ────────────────────────────────────────────────────────────────

def value_in_passage(value, text):
    """
    Does the extracted number actually appear in the passage quoted beside it?

    96.2% of findings pass this (98.5% of reported ICERs). The point of the
    check is the other 3.8%: a citation a reader can check is only worth
    having if the app admits the cases where the check fails, rather than
    letting them find it themselves.

    French documents separate thousands with a space, a non-breaking space or
    a dot, so each form is tried.
    """
    try:
        n = float(value)
    except (TypeError, ValueError):
        return None
    if n != n or n in (float("inf"), float("-inf")):   # NaN / inf
        return None
    body = str(text or "")
    if not body:
        return None
    i = int(round(n))
    forms = {str(i)} | {f"{i:,}".replace(",", sep)
                        for sep in (" ", " ", " ", ".", ",")}
    return any(f in body for f in forms)


def _render_passage(text, url, page, caption=None, verified=None):
    """One cited fragment: the text, then a link to its page."""
    body = _text(text)
    if not body:
        return
    body, continued = clean_start(body)
    shown, clipped = trim(body)
    quote = ("… " if continued else "") + _md_safe(shown)
    st.markdown("> " + quote.replace("\n", "\n> "))
    if clipped:
        with st.expander("Show the full passage"):
            st.markdown(_md_safe(body))
    link = page_url(url, page)
    bits = []
    if link:
        bits.append(f"[Open the source at page {int(page)} ↗]({link})")
    if caption:
        bits.append(caption)
    if bits:
        st.caption(" · ".join(bits))
    if verified is False:
        st.caption("⚠️ The figure above was not located in this extract — "
                   "the opinion states it elsewhere on the page. Open the "
                   "page before quoting it.")


def _render_glossary(hits):
    if not hits:
        return
    lines = "  \n".join(f"**{term}** — {gloss}" for term, gloss in hits)
    st.info(lines)


def _render_gba(appraisal_id, max_sections):
    rows = (_GBA or {}).get(appraisal_id)
    if rows is None or not len(rows):
        return False

    order = {name: i for i, name in enumerate(_GBA_SECTION_ORDER)}
    sections = sorted(
        rows.groupby("section", sort=False),
        key=lambda kv: order.get(str(kv[0]), len(order)),
    )

    st.markdown("**Cited source passages — G-BA resolution**")
    st.caption(f"{len(rows)} passage(s) from the published Beschluss. "
               "Original German retained; every link opens the PDF at the page "
               "the text came from.")

    for name, group in sections[:max_sections]:
        group = group.sort_values("part") if "part" in group.columns else group
        pages = sorted({int(p) for p in group["document_page"].dropna()})
        label = f"{_text(name) or 'Passage'} · page {', '.join(map(str, pages))}"
        with st.expander(label, expanded=(str(name) == _GBA_SECTION_ORDER[0])):
            for _, r in group.iterrows():
                part = r.get("part")
                caption = None
                if pd.notna(part) and len(group) > 1:
                    caption = f"part {int(part)} of {len(group)}"
                _render_passage(r.get("source_text"), r.get("source_url"),
                                r.get("document_page"), caption=caption)

    if len(sections) > max_sections:
        st.caption(f"{len(sections) - max_sections} further section(s) not shown.")
    # Once per record, not once per section: glossing "zweckmäßige
    # Vergleichstherapie" four times down one page is noise, and the terms
    # recur across sections by their nature.
    _render_glossary(_glossary_hits(*rows["source_text"].tolist()))
    return True


def _format_value(value, unit):
    v, u = _text(value), _text(unit)
    if not v:
        return None
    return f"{v} {u}" if u and u.lower() != "see passage" else v


def _render_has(appraisal_id, max_findings):
    findings = (_FINDINGS or {}).get(appraisal_id)
    pages = (_HAS_PAGES or {}).get(appraisal_id)
    if (findings is None or not len(findings)) and (pages is None or not len(pages)):
        return False

    if findings is not None and len(findings):
        is_icer = findings["finding_type"].astype(str).isin(_ICER_TYPES)
        st.markdown("**Economic findings — HAS efficiency opinion**")
        st.caption(
            f"{int(is_icer.sum())} reported cost-effectiveness ratio(s) and "
            f"{int((~is_icer).sum())} other economic value(s) extracted from the "
            "avis d'efficience. Scenario and comparator are NOT recorded, so a "
            "value is not comparable with another appraisal's until you have "
            "checked the passage.")

        for block_label, subset, note in (
            ("Reported cost-effectiveness ratios", findings[is_icer], None),
            ("Other economic values", findings[~is_icer],
             "Thresholds and unverified ratios — these are not ICERs."),
        ):
            if not len(subset):
                continue
            st.markdown(f"*{block_label}*")
            if note:
                st.caption(note)
            for _, r in subset.head(max_findings).iterrows():
                value = _format_value(r.get("value"), r.get("unit"))
                ftype = _text(r.get("finding_type"))
                head = " · ".join(x for x in (value, ftype) if x)
                if head:
                    st.markdown(f"**{_md_safe(head)}**")
                _render_passage(r.get("source_context"), r.get("source_url"),
                                r.get("document_page"),
                                caption=_text(r.get("verification")),
                                verified=value_in_passage(r.get("value"),
                                                          r.get("source_context")))
            if len(subset) > max_findings:
                st.caption(f"{len(subset) - max_findings} further value(s) not shown.")

    if pages is not None and len(pages):
        with st.expander(f"Full economic opinion text ({len(pages)} page(s))"):
            for _, r in pages.sort_values("document_page").iterrows():
                st.markdown(f"**Page {int(r['document_page'])}**")
                _render_passage(r.get("source_text"), r.get("source_url"),
                                r.get("document_page"),
                                caption=_text(r.get("extraction_note")))

    gloss_source = []
    if findings is not None and len(findings):
        gloss_source += findings["source_context"].tolist()
    if pages is not None and len(pages):
        gloss_source += pages["source_text"].head(3).tolist()
    _render_glossary(_glossary_hits(*gloss_source))
    return True


def render_panel(agency, appraisal_id, max_sections=4, max_findings=6):
    """
    Cited source passages for one G-BA or HAS record.

    Silent — renders nothing at all — for NICE, for a record with no indexed
    passages, and for a workbook without the sheets. The caller does not need
    to guard the call.
    """
    if not is_attached():
        return False
    key = _text(appraisal_id)
    if not key:
        return False
    if agency == "G-BA":
        return _render_gba(key, max_sections)
    if agency == "HAS":
        return _render_has(key, max_findings)
    return False
