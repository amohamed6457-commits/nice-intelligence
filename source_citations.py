import hashlib
import html as _html
import json
import os
import re

import pandas as pd
import streamlit as st

# ── state ────────────────────────────────────────────────────────────────────

_GBA = None        # appraisal_id -> DataFrame of passages
_HAS_PAGES = None
_FINDINGS = None
_COSTS = None      # appraisal_id -> DataFrame of GBA_Treatment_Costs rows
_TABLES = None     # appraisal_id -> [{page, header, rows, notes}]
_TRANSLATIONS = None       # sha1(text)[:16] -> English rendering
TRANSLATIONS_FILE = "translations.json"

TABLES_FILE = "gba_tables.json"

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

# Phrases a G-BA resolution uses to start a new idea. Breaking before each
# one turns a 10,000-character run of text into readable blocks. English and
# German, because the EN files are courtesy translations and the originals
# are German.
_BREAK_BEFORE = [
    "Appropriate comparator therapy:", "Extent and probability of the additional benefit",
    "Study results according to endpoints:", "Summary of results for relevant clinical endpoints",
    "Explanations:", "Endpoint category", "Probability and extent",
    "Zweckmäßige Vergleichstherapie:", "Ausmaß und Wahrscheinlichkeit",
    "Studienergebnisse nach Endpunkten:", "Erläuterungen:",
    "Wahrscheinlichkeit und Ausmaß",
]
# Subgroup markers: 'a1)', 'b2)', 'c1)' and numbered headings '1.' .. '9.'
_MARKER = re.compile(r"(?<=[\s.;:])(?=(?:[a-h]\d\)|\d\.)\s+[A-ZÄÖÜ])")

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

def attach(gba_passages=None, has_pages=None, economic_findings=None,
           gba_costs=None):
    """
    Index the sheets once per run. Safe to call repeatedly.

    gba_costs is optional and last so an older call site with three
    arguments keeps working; without it the cost table simply isn't drawn
    and the Treatment costs passage renders on its own.
    """
    global _GBA, _HAS_PAGES, _FINDINGS, _COSTS, _TABLES
    _GBA = _by_appraisal(gba_passages)
    _HAS_PAGES = _by_appraisal(has_pages)
    _FINDINGS = _by_appraisal(economic_findings)
    _COSTS = _by_appraisal(gba_costs)
    _TABLES = _load_tables()


@st.cache_data(show_spinner=False)
def _load_tables(path=TABLES_FILE):
    """
    Tables recovered from the resolution PDFs by extract_gba_tables.py.

    Absent until that script has been run, in which case every G-BA record
    simply shows its passage text as before. Never fatal: a missing or
    malformed file degrades to no tables rather than taking the page down.
    """
    import json
    import os
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
    except (ValueError, OSError):
        return {}
    return data if isinstance(data, dict) else {}


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

def structure(text):
    """
    Break one long run of resolution text into readable blocks.

    A G-BA passage can be 10,000 characters that the PDF laid out as
    headings, subgroup markers and paragraphs. The extraction kept the
    words and the line order but nothing that says 'new idea starts here',
    so it arrives as a wall. These documents are formulaic, which is what
    makes this safe: every subgroup is 'a1)', 'b2)', 'c1)', every section
    opens with a number, and the same dozen headings recur.

    Only whitespace is inserted. No word is changed, moved or dropped.
    """
    s = str(text).strip()
    if not s:
        return ""
    for phrase in _BREAK_BEFORE:
        s = s.replace(phrase, "\n\n" + phrase)
    s = _MARKER.sub("\n\n", s)
    # Collapse runs of blank lines the two passes may have doubled up.
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def highlight_html(text, query):
    """
    The passage as HTML, with every match for `query` marked.

    Returns (html, match_count). The text is HTML-escaped BEFORE any markup
    is inserted, and the query is regex-escaped, so neither the workbook nor
    anything typed into the box can inject markup into the page.

    Several words separated by spaces are treated as alternatives, so
    'mortality morbidity' highlights both.
    """
    body = _html.escape(str(text))
    terms = [t for t in str(query or "").split() if len(t) > 1]
    if not terms:
        return body.replace("\n", "<br>"), 0
    pattern = re.compile("|".join(re.escape(t) for t in terms), re.IGNORECASE)
    count = len(pattern.findall(body))
    marked = pattern.sub(
        lambda m: f'<mark style="background:#F6C945;color:#111;'
                  f'border-radius:2px;padding:0 2px">{m.group(0)}</mark>',
        body)
    return marked.replace("\n", "<br>"), count


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


def _load_translations(path=TRANSLATIONS_FILE):
    """English renderings produced by translate_passages.py, keyed by a hash
    of the ORIGINAL text. Missing file is not an error: the panel simply
    stays monolingual."""
    global _TRANSLATIONS
    if _TRANSLATIONS is not None:
        return _TRANSLATIONS
    _TRANSLATIONS = {}
    try:
        if os.path.exists(path):
            with open(path, encoding="utf-8") as fh:
                loaded = json.load(fh)
            if isinstance(loaded, dict):
                _TRANSLATIONS = {k: v for k, v in loaded.items()
                                 if isinstance(v, str)
                                 and not v.startswith("__ERROR__")}
    except (ValueError, OSError):
        _TRANSLATIONS = {}
    return _TRANSLATIONS


def translations_loaded():
    return bool(_load_translations())


def translation_for(text):
    """English for one passage, or None. Keyed exactly as the translator
    wrote it: sha1 of the stripped original, first 16 hex characters."""
    body = _text(text)
    if not body:
        return None
    key = hashlib.sha1(body.strip().encode("utf-8")).hexdigest()[:16]
    return _load_translations().get(key)


def collapse_repeats(text, min_shared=30, min_words=4):
    """Collapse near-identical paragraphs that differ only by subgroup.

    G-BA files its reasoning per subgroup, and app3 renders the joined
    field, so a three-subgroup appraisal prints the same 900 characters
    three times and diverges only at "On patient group a2) / b1) / b2)".
    Reading it, you cannot see where the repetition ends and the difference
    begins -- which is the part that actually matters, because the
    comparator usually differs between subgroups.

    So the shared opening is printed once and each paragraph contributes
    only its own tail. Nothing is discarded: every word still appears.
    """
    body = _text(text)
    if not body:
        return text
    parts = [p.strip() for p in re.split(r"\n\s*\n|\r\n\s*\r\n", body)
             if p.strip()]
    if len(parts) < 2:
        parts = [p.strip() for p in body.split("\n") if p.strip()]
    # Same paragraph filed twice: keep one, preserve order.
    unique = list(dict.fromkeys(parts))
    fallback = text if len(unique) == len(parts) else "\n\n".join(unique)
    parts = unique
    if len(parts) < 2:
        return fallback

    shared = parts[0]
    for part in parts[1:]:
        limit = min(len(shared), len(part))
        i = 0
        while i < limit and shared[i] == part[i]:
            i += 1
        shared = shared[:i]

    # Trim to a word boundary only if the match stopped MID-word, so a
    # complete last word ("mellitus") is kept.
    if any(len(p) > len(shared) and p[len(shared)].isalnum() for p in parts):
        shared = shared[:shared.rfind(" ")] if " " in shared else ""
    shared = shared.rstrip()
    if len(shared) < min_shared or len(shared.split()) < min_words:
        return fallback

    tails = [p[len(shared):].lstrip(" ,;:-").strip() for p in parts]
    if not all(tails):
        return fallback
    return shared + ":\n\n" + "\n".join("- " + t for t in tails)

def display_text(value):
    """What to SHOW for any field on the record.

    The English rendering when the reader has asked for English and one
    exists; otherwise the value exactly as it came from the workbook. Never
    raises and never returns None for a non-None input, so a caller can wrap
    any field with it safely.
    """
    body = _text(value)
    if not body:
        return value
    if _show_english():
        body = translation_for(body) or body
    # Collapse the per-subgroup repetition AFTER translation, so the shared
    # opening is found in whichever language is on screen.
    return collapse_repeats(body)


def _show_english():
    # The "w_" prefix matters: app3 re-saves every w_-keyed value at the top
    # of each run, so the choice survives switching HTA body. Without it,
    # Streamlit drops the value on any run where the widget is not drawn.
    try:
        return st.session_state.get("w_src_lang_mode", "Original") == "English"
    except Exception:                                      # noqa: BLE001
        return False


def _render_passage(text, url, page, caption=None, verified=None, query=None):
    """One cited fragment: the text, then a link to its page.

    In English mode the translation is shown in place of the original and
    labelled as machine translation, but the ORIGINAL is always one click
    away and the page link always opens the official PDF in its own
    language. A translated passage is a reading aid, never the record.
    """
    body = _text(text)
    if not body:
        return
    original = body
    english = translation_for(body) if _show_english() else None
    if english:
        body = english
    body, continued = clean_start(body)
    shown, clipped = trim(body)
    quote = ("… " if continued else "") + _md_safe(shown)
    # Two trailing spaces force a hard line break. Without them markdown
    # collapses every single newline into a space, which turns a laid-out
    # table in the resolution -- 'Pembrolizumab € 82,806.40 - € 87,981.80'
    # on its own line -- into one unreadable run of words. The source text
    # keeps the document's line structure; only the rendering was losing it.
    st.markdown("> " + quote.replace("\n", "  \n> "))
    if clipped:
        structured = structure(body)
        marked, hits = highlight_html(structured, query)
        label = ("Show the full passage"
                 if not query else
                 f"Show the full passage — {hits} match"
                 f"{'' if hits == 1 else 'es'} for '{query}'")
        with st.expander(label, expanded=bool(query and hits)):
            # The full text goes through structure() and is rendered as HTML
            # so that the inserted paragraph breaks survive and matches can
            # be marked. Everything is escaped in highlight_html before any
            # markup is added.
            st.markdown(
                f'<div style="line-height:1.65;max-height:60vh;overflow-y:auto;'
                f'padding-right:8px">{marked}</div>',
                unsafe_allow_html=True)
    if english:
        with st.expander("Show the original wording", expanded=False):
            st.markdown("> " + _md_safe(original).replace("\n", "  \n> "))
    link = page_url(url, page)
    bits = []
    if link:
        bits.append(f"[Open the source at page {int(page)} ↗]({link})")
    if english:
        bits.append("machine translation — the original is the record")
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


def _eur(low, high):
    """'€ 82,806.40 - € 87,981.80', or a single figure where they match."""
    def one(v):
        try:
            return f"€ {float(v):,.2f}"
        except (TypeError, ValueError):
            return None
    lo, hi = one(low), one(high)
    if lo and hi and lo != hi:
        return f"{lo} – {hi}"
    return lo or hi or "—"


def _render_cost_table(appraisal_id):
    """
    The treatment-cost table as a table, from GBA_Treatment_Costs.

    One honest limitation, stated on screen rather than buried: the sheet
    holds NUMERIC rows only. A comparator the G-BA records as "Not
    calculable" — which is common, and is itself a finding — has no row
    here, so a reader looking only at this table would see the assessed
    medicine's cost with no comparison. 627 of 1,041 resolutions carry both
    an assessed-medicine and a comparator row; 274 carry the medicine alone.
    The passage and the page link below it remain the authority.
    """
    rows = (_COSTS or {}).get(appraisal_id)
    if rows is None or not len(rows):
        return False

    table = pd.DataFrame({
        "Role": rows.get("therapy_role", pd.Series(dtype=object)).map(
            lambda v: _text(v) or "Not assigned"),
        "Therapy": rows.get("therapy_label", pd.Series(dtype=object)).map(
            lambda v: _text(v) or "—"),
        "Annual cost per patient": [
            _eur(lo, hi) for lo, hi in zip(rows.get("annual_cost_lower_EUR", []),
                                           rows.get("annual_cost_upper_EUR", []))
        ],
        "Page": rows.get("document_page", pd.Series(dtype=object)).map(
            lambda v: int(v) if pd.notna(v) else None),
    })
    st.markdown("**Extracted cost rows**")
    st.dataframe(table, width="stretch", hide_index=True)

    roles = set(table["Role"])
    if "Assessed medicine" in roles and not any("omparator" in r for r in roles):
        st.caption("⚠️ No comparator cost row was extracted for this "
                   "resolution. The G-BA often records the comparator as "
                   "'not calculable', which carries no figure and so has no "
                   "row here — read the passage below before concluding "
                   "anything about relative cost.")
    else:
        st.caption("Numeric rows only. Entries the resolution states without "
                   "a figure (for example 'not calculable') do not appear "
                   "here — the passage below is the full text.")
    return True


def _section_page_range(rows, group):
    """
    The pages a section actually covers, not the one page it is filed under.

    GBA_Decision_Passages records the page a section STARTS on. D-1058's
    'Decision / subgroup findings' is filed at page 2, but its content — the
    mortality, morbidity, quality-of-life and adverse-event tables — runs to
    page 7. Matching tables against the filed page alone finds nothing, which
    is why no table appeared even once the extraction had run.

    A section therefore owns every page from where it starts up to where the
    next section starts.
    """
    starts = sorted({int(p) for p in rows["document_page"].dropna()})
    first = min(int(p) for p in group["document_page"].dropna())
    later = [p for p in starts if p > first]
    last = (min(later) - 1) if later else 10_000
    return set(range(first, last + 1))


def _render_tables_for_pages(appraisal_id, pages):
    """
    Any recovered tables sitting on these pages, drawn as tables.

    This is the whole point of the extraction: a passage that was a table in
    the resolution is close to unreadable as text, because the line order
    survives but the columns do not. Where a table was recovered AND passed
    every alignment gate, it is shown here and the text below it becomes the
    backup rather than the only option.
    """
    tables = (_TABLES or {}).get(str(appraisal_id)) or []
    on_page = [t for t in tables if t.get("page") in pages]
    if not on_page:
        return False

    for t in on_page:
        rows = t.get("rows") or []
        if not rows:
            continue
        header = t.get("header") or []
        width = max([len(r) for r in rows] + [len(header)])
        # A table that continues from the previous page, or opens with a
        # section heading, has no header of its own. Inventing 'Column 1,
        # Column 2' for it is worse than nothing — the real document shows
        # no heading there either. Non-breaking spaces give Streamlit the
        # unique, non-empty names it requires while rendering as blank.
        if header:
            header = (list(header) + [""] * width)[:width]
        else:
            header = [" " * (i + 1) for i in range(width)]
        seen, cols = {}, []
        for i, h in enumerate(header):
            name = (h or " " * (i + 1)).strip("\n\t ") or " " * (i + 1)
            seen[name] = seen.get(name, 0) + 1
            cols.append(name if seen[name] == 1 else f"{name} ({seen[name]})")
        body = [(list(r) + [""] * width)[:width] for r in rows]
        st.dataframe(pd.DataFrame(body, columns=cols),
                     width="stretch", hide_index=True)
        for note in t.get("notes") or []:
            st.caption(_md_safe(trim(note, 400)[0]))
    st.caption("Recovered from the resolution PDF. Tables whose columns could "
               "not be matched to their headings are not shown here — the "
               "passage below is always the authority.")
    return True


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
    # Say plainly when the recovered tables are simply not present, rather
    # than letting a missing file look like a broken feature.
    if not _TABLES:
        st.caption(f"ℹ️ No recovered tables loaded — `{TABLES_FILE}` is not "
                   "beside app3.py. Run extract_gba_tables.py to render the "
                   "resolution's tables as tables instead of flattened text.")
    elif str(appraisal_id) not in _TABLES:
        st.caption("ℹ️ This resolution has not been through table extraction "
                   "yet.")

    # One box for the whole record. Typing in it marks every match across
    # all the passages below and opens the ones that contain any, so a
    # 10,000-character resolution can be searched without reading it.
    query = st.text_input(
        "Find in these passages",
        # app3 re-saves every 'w_'-prefixed widget value at the top of each
        # run, which is what stops Streamlit discarding it when the widget
        # is not drawn — switching HTA body and back would otherwise clear
        # the box. The id in the key keeps each record's search its own.
        key=f"w_cite_find_{appraisal_id}",
        placeholder="e.g. Zusatznutzen, comparator, mortality, Überlebenszeit",
        help="Several words are treated as alternatives — 'mortality morbidity' "
             "finds both.")

    for name, group in sections[:max_sections]:
        group = group.sort_values("part") if "part" in group.columns else group
        pages = sorted({int(p) for p in group["document_page"].dropna()})
        label = f"{_text(name) or 'Passage'} · page {', '.join(map(str, pages))}"
        with st.expander(label, expanded=(str(name) == _GBA_SECTION_ORDER[0])):
            # The cost table reads far better as a table than as the
            # document's own text layout flattened into a quotation.
            if str(name) == "Treatment costs":
                _render_cost_table(appraisal_id)
            # Any table recovered from the pages this section covers goes
            # above the text it came from, so the readable form is what the
            # eye lands on. The range matters: see _section_page_range.
            _render_tables_for_pages(appraisal_id,
                                     _section_page_range(rows, group))
            for _, r in group.iterrows():
                part = r.get("part")
                caption = None
                if pd.notna(part) and len(group) > 1:
                    caption = f"part {int(part)} of {len(group)}"
                _render_passage(r.get("source_text"), r.get("source_url"),
                                r.get("document_page"), caption=caption,
                                query=query)

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
    # The language control is drawn by app3's sidebar, beside the HTA body
    # switch, so it is visible before a record is opened. This module only
    # READS the choice, through _show_english().
    if agency == "G-BA":
        return _render_gba(key, max_sections)
    if agency == "HAS":
        return _render_has(key, max_findings)
    return False
