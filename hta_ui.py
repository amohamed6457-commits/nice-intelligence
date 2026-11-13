"""
The agency screens: explorer, record view, charts, results, export and cross-agency matching.

Moved out of app3.py by split_app3.py. These functions reach only
constants, never the dataframe or session state, which is what made
the move mechanical.
"""

import inspect
import io
import re
from collections import Counter
from datetime import date
import pandas as pd
import plotly.express as px
import streamlit as st
from reportlab.lib import colors
from hta_data import *  # noqa: F401,F403
from hta_core import *  # noqa: F401,F403
from hta_chat import *  # noqa: F401,F403

try:
    import source_citations
except ImportError:  # optional companion module
    source_citations = None
try:
    import endpoint_panel
except ImportError:  # optional companion module
    endpoint_panel = None
try:
    import evidence_filter
except ImportError:  # optional companion module
    evidence_filter = None
try:
    import nice_attributes
except ImportError:  # optional companion module
    nice_attributes = None

__all__ = [
    "AGENCY_BASE_COLUMNS",
    "AGENCY_FLAGS",
    "AGENCY_LABELS",
    "AGENCY_SORT",
    "AREA_ORDER",
    "ASMR_RAMP",
    "CHART_DEFAULTS",
    "CHART_OPTIONS",
    "CROSS_AGENCY_CAVEAT",
    "DEFERRED_DOWNLOADS",
    "EAGER_EXPORT_MAX_ROWS",
    "EXPLORER_FIELDS",
    "FAMILY_COLORS",
    "MATCH_COLUMNS",
    "MAX_AREA_BARS",
    "NEUTRAL_GREY",
    "RESULTS_SECTIONS",
    "SINGLE_SERIES",
    "SMALL_AGENCY_SET",
    "SMR_ORDER",
    "SYNTHESIS_MAX_ITEMS",
    "TOTAL_ROWS",
    "XLSX_MIME",
    "_agency_export_bytes",
    "_agency_footer",
    "_block_subset",
    "_clear_agency_search",
    "_counted",
    "_coverage_line",
    "_cross_agency_lines",
    "_explorer_key",
    "_integer_axis",
    "_is_narrowed",
    "_query_key",
    "_records_matched",
    "_results_signature",
    "_search_ranking",
    "_section_nav",
    "_style_bars",
    "_supports_deferred_download",
    "agency_match_table",
    "agency_sidebar_guide",
    "cached_agency_search",
    "cross_agency_index",
    "describe_query",
    "df",
    "find_same_molecule",
    "init_range_state",
    "outcome_breakdown_table",
    "outcome_order",
    "prepare_agency_frame",
    "render_agency_charts",
    "render_agency_explorer",
    "render_agency_record",
    "render_agency_switcher",
    "render_agency_view",
    "render_download",
    "render_economic_block",
    "render_nice_cross_agency",
    "render_outcome_table",
    "render_results_decisions",
    "render_results_evidence",
    "render_results_header",
    "render_results_other_bodies",
    "render_synthesis_block",
    "run_agency_search",
]

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
df = load_data(data_version())
TOTAL_ROWS = len(df)
AGENCY_LABELS = {
    "NICE": "🇬🇧 NICE — England",
    "G-BA": "🇩🇪 G-BA — Germany",
    "HAS": "🇫🇷 HAS — France",
}
AGENCY_FLAGS = {"NICE": "🇬🇧", "G-BA": "🇩🇪", "HAS": "🇫🇷"}
AGENCY_SORT = {"NICE": 0, "G-BA": 1, "HAS": 2}
SMALL_AGENCY_SET = 20       # below this, a view carries an "illustrative only" note
SYNTHESIS_MAX_ITEMS = 15    # per evidence-synthesis block before "narrow the selection"
SMR_ORDER = ["Important", "Moderate", "Low", "Mixed", "Insufficient"]
FAMILY_COLORS = {
    "light": {
        "Added benefit": "#008300", "Mixed by subgroup": "#2a78d6", "No added benefit": "#e34948",
        "Improvement (ASMR I–IV)": "#008300", "No improvement (ASMR V)": "#e87ba4",
        "SMR only, no ASMR": "#4a3aa7", "SMR insufficient": "#e34948", "Early access": "#2a78d6",
        "No rating (yet)": "#a8a7a1", "No rating recorded": "#a8a7a1",
    },
    "dark": {
        "Added benefit": "#008300", "Mixed by subgroup": "#3987e5", "No added benefit": "#e66767",
        "Improvement (ASMR I–IV)": "#008300", "No improvement (ASMR V)": "#d55181",
        "SMR only, no ASMR": "#9085e9", "SMR insufficient": "#e66767", "Early access": "#3987e5",
        "No rating (yet)": "#62615c", "No rating recorded": "#62615c",
    },
}
ASMR_RAMP = {"light": ["#0f5a1f", "#1d7a2e", "#2f9a42", "#5cbf6a"],
             "dark": ["#a3e0ab", "#74cc80", "#4cb45c", "#2f9a42"]}
SINGLE_SERIES = {"light": "#2a78d6", "dark": "#3987e5"}
NEUTRAL_GREY = {"light": "#a8a7a1", "dark": "#62615c"}
AREA_ORDER = [
    "Oncology", "Haematology", "Immunology / Inflammatory", "Infectious disease",
    "Endocrine / Metabolic", "Neurology", "Mental health", "Cardiovascular", "Respiratory",
    "Gastroenterology / Hepatology", "Renal / Urology", "Musculoskeletal", "Dermatology",
    "Ophthalmology", "Reproductive & sexual health", "Pain & inflammation",
    "Rare / Genetic diseases", "Other", "Not classified",
]
CROSS_AGENCY_CAVEAT = (
    "Matched on active substance (including combination regimens), not on indication — "
    "check the indication before comparing outcomes. The bodies answer different questions: "
    "NICE weighs cost-effectiveness for the NHS, the G-BA rates added clinical benefit "
    "against a comparator it sets, and HAS rates clinical benefit (SMR/ASMR) and decides "
    "early access. A different outcome is expected, not a contradiction."
)
AGENCY_BASE_COLUMNS = [
    "appraisal_id", "drug_name", "generic_name", "brand_name", "indication",
    "therapeutic_area", "decision_raw", "decision_simple", "decision_date",
    "publication_date", "assessment_type", "patient_population", "decision_reasoning",
    "ICER_status", "official_source_url", "methodology_note",
]
@st.cache_data
def prepare_agency_frame(agency, version):
    """One agency's sheet with the columns every view relies on added:
    inn, brand, comparator, EUR ICER bounds, decision_dt / year / date_label,
    outcome, outcome_family, favourable, therapy_area_en (+ _areas, every
    area), procedure_type, condition, orphan, _molecules, and accent-folded
    search text (_primary_text: names, condition, indication, category;
    _secondary_text: decision wording and the start of the reasoning), plus
    any agency-specific derived fields. The row index is kept as the record
    key, so two rows sharing an ID (e.g. an initial decision and its renewal)
    stay two records everywhere."""
    cfg = AGENCY_CONFIG[agency]
    raw = load_agency_sheet(cfg["sheet"], version)
    if raw.empty:
        return raw
    frame = raw.copy()
    frame.columns = [str(c).strip() for c in frame.columns]
    for col in AGENCY_BASE_COLUMNS + cfg["extra_columns"]:
        if col not in frame.columns:
            frame[col] = None

    # Blank IDs get the Excel row number so every record can still be traced.
    frame["appraisal_id"] = [_text(v) or f"{agency} row {n + 2}"
                             for n, v in enumerate(frame["appraisal_id"])]
    frame["assessment_type"] = frame["assessment_type"].map(_text)
    frame["indication"] = frame["indication"].map(lambda v: _text(v) or "Indication not recorded")
    frame["inn"] = [
        _clean_inn(_text(g)) or (_text(d).title() if _text(d) else "Unknown")
        for g, d in zip(frame["generic_name"], frame["drug_name"])
    ]
    frame["brand"] = [_text(b) or "" for b in frame["brand_name"]]
    frame["comparator"] = frame[cfg["comparator_col"]]
    frame["icer_eur_lower"] = pd.to_numeric(frame[cfg["icer_lower_col"]], errors="coerce")
    frame["icer_eur_upper"] = (pd.to_numeric(frame[cfg["icer_upper_col"]], errors="coerce")
                               if cfg["icer_upper_col"] else float("nan"))

    decided = _to_datetime(frame["decision_date"])
    frame["decision_dt"] = decided.fillna(_to_datetime(frame["publication_date"]))
    frame["year"] = frame["decision_dt"].dt.year
    frame["date_label"] = frame["decision_dt"].dt.strftime("%d %b %Y").fillna("Date not recorded")

    # The per-record classifiers take one plain dict per row, built once: a
    # row-wise DataFrame.apply builds a 45-column Series for every row on every
    # pass, which made this step ~14s on the 12k-row HAS sheet.
    records = frame.to_dict("records")

    def per_record(name, fn):
        values = [fn(r) for r in records]
        for r, value in zip(records, values):
            r[name] = value
        frame[name] = values

    per_record("outcome", cfg["outcome_fn"])
    for name, fn in cfg["derived"].items():
        per_record(name, fn)
    per_record("favourable", cfg["favourable_fn"])
    frame["outcome_family"] = [outcome_family(agency, o) for o in frame["outcome"]]

    atc = frame["ATC_codes"] if "ATC_codes" in frame.columns else [None] * len(frame)
    areas = [therapy_areas(agency, a, c) for a, c in zip(frame["therapeutic_area"], atc)]
    frame["therapy_area_en"] = [primary for primary, _ in areas]
    frame["_areas"] = [every for _, every in areas]
    frame["procedure_type"] = ([gba_procedure_group(r) for r in records] if agency == "G-BA"
                               else frame["assessment_type"].map(has_procedure_group))
    frame["condition"] = [condition_label(agency, r) for r in records]
    frame["orphan"] = (frame["orphan_status"].map(orphan_flag)
                       if "orphan_status" in frame.columns else None)

    def folded(cols):
        joined = None
        for col in (c for c in cols if c in frame.columns):
            part = frame[col].map(lambda v: _text(v) or "").astype(str)
            joined = part if joined is None else joined + " | " + part
        return _fold_series(joined) if joined is not None else pd.Series("", index=frame.index)

    name_cols = ["inn", "brand", "drug_name", "generic_name", "source_title",
                 "title_generic_name"]
    frame["_drug_text"] = folded(name_cols)
    # The therapy-area tag is searched too: G-BA's names the disease
    # ('Mammakarzinom (onkologische Erkrankungen)'), HAS's often does ('Sein/
    # Cancer du sein'). HAS's anatomical ones ('Appareil digestif/Œsophage,
    # estomac, intestin') are why CONDITION_TERMS uses disease phrases
    # ('cancer de l'oesophage'), never a bare organ word, for its cancers.
    frame["_primary_text"] = folded(
        ["appraisal_id"] + name_cols + ["condition", "indication", "therapeutic_area"])
    secondary_cols = [c for c in ("decision_simple", "decision_raw", "added_benefit_assessment",
                                  "patient_population", "comparator") if c in frame.columns]
    frame["_secondary_text"] = (folded(secondary_cols) + " | "
                                + _fold_series(frame["decision_reasoning"], limit=700))
    frame["_molecules"] = [
        molecule_set(g) if _text(g) else molecule_set(d)
        for g, d in zip(frame["generic_name"], frame["drug_name"])
    ]
    return frame
@st.cache_data
def cross_agency_index(version):
    """Every NICE, G-BA and HAS record in one light frame, carrying the
    active-substance keys used to line the same molecule up across agencies."""
    nice = load_data(data_version())
    frames = [pd.DataFrame({
        "body": "NICE",
        "appraisal_id": nice["appraisal_id"].astype(str),
        "drug": nice["drug_name"].astype(str),
        "indication": nice["indication"],
        "decision": nice["decision_simple"],
        "when": nice["year_label"].astype(str),
        "url": nice["url"],
        "_molecules": [molecule_set(d, g) for d, g in zip(nice["drug_name"], nice["generic_name"])],
    })]
    for agency in AGENCY_CONFIG:
        f = prepare_agency_frame(agency, version)
        if f.empty:
            continue
        frames.append(pd.DataFrame({
            "body": agency,
            "appraisal_id": f["appraisal_id"].astype(str),
            "drug": [f"{i} ({b})" if b else i for i, b in zip(f["inn"], f["brand"])],
            "indication": f["indication"],
            "decision": f["outcome"],
            "when": f["date_label"],
            "url": f["official_source_url"],
            "_molecules": f["_molecules"],
        }))
    return pd.concat(frames, ignore_index=True)
MATCH_COLUMNS = ["_src", "body", "appraisal_id", "drug", "indication", "decision", "when",
                 "url", "_body_sort", "_overlap"]
@st.cache_data
def agency_match_table(agency, version):
    """
    Every same-molecule match for one agency's rows, as one long table sorted
    by record, then agency, then closest indication. _src is the record's row
    index in prepare_agency_frame. Built once through an inverted molecule
    index and cached — scanning every record against every other on each
    rerun took seconds once the sheets reach ~1,000 rows.
    """
    frame = prepare_agency_frame(agency, version)
    if frame.empty:
        return pd.DataFrame(columns=MATCH_COLUMNS)
    index = cross_agency_index(version)
    others = index[index["body"] != agency]
    by_molecule, details = {}, {}
    for pos, body, aid, drug, ind, decision, when, url, molecules in zip(
            others.index, others["body"], others["appraisal_id"], others["drug"],
            others["indication"], others["decision"], others["when"], others["url"],
            others["_molecules"]):
        details[pos] = (body, aid, drug, ind, decision, when, url)
        for key in molecules:
            by_molecule.setdefault(key, []).append(pos)
    records = []
    for src, molecules, indication in zip(frame.index, frame["_molecules"], frame["indication"]):
        for pos in sorted({p for key in molecules for p in by_molecule.get(key, ())}):
            body, aid, drug, ind, decision, when, url = details[pos]
            records.append((src, body, aid, drug, ind, decision, when, url,
                            AGENCY_SORT.get(body, 9), indication_overlap(indication, ind)))
    table = pd.DataFrame.from_records(records, columns=MATCH_COLUMNS)
    return table.sort_values(["_src", "_body_sort", "_overlap"], ascending=[True, True, False],
                             kind="stable").reset_index(drop=True)
def find_same_molecule(index, molecules, exclude_body, ref_indication=None):
    """Records at other agencies sharing an active substance, closest indication first."""
    if not molecules:
        return index.iloc[0:0]
    hits = index[(index["body"] != exclude_body)
                 & index["_molecules"].map(lambda m: bool(m & molecules))]
    if hits.empty:
        return hits
    hits = hits.assign(
        _body_sort=hits["body"].map(AGENCY_SORT),
        _overlap=hits["indication"].map(lambda t: indication_overlap(ref_indication, t)),
    )
    return hits.sort_values(["_body_sort", "_overlap"], ascending=[True, False])
def _cross_agency_lines(hits, limit=8):
    lines = []
    for _, h in hits.head(limit).iterrows():
        link = f" · {_md_link('source', h['url'])}" if _text(h["url"]) else ""
        lines.append(
            f"- {AGENCY_FLAGS.get(h['body'], '')} **{h['body']} {h['appraisal_id']}** · "
            f"{_md_safe(h['drug'])} — {h['decision']} ({h['when']}) — "
            f"_{_md_safe(_shorten(h['indication'], 110))}_{link}")
    if len(hits) > limit:
        lines.append(f"- …and {len(hits) - limit} more")
    return "\n".join(lines)
def render_nice_cross_agency(nice_rows):
    """
    'Also assessed by' lines for the drug open in the NICE Drug Detail panel,
    so a NICE-first user sees the G-BA / HAS verdicts on the same molecule
    without switching views. Guarded: a problem in the G-BA / HAS sheets must
    never take the NICE view down with it.
    """
    if nice_rows.empty:
        return
    try:
        molecules = frozenset().union(*(
            molecule_set(d, g) for d, g in zip(nice_rows["drug_name"], nice_rows["generic_name"])))
        hits = find_same_molecule(cross_agency_index(data_version()), molecules, "NICE",
                                  " ".join(nice_rows["indication"].astype(str)))
    except Exception:
        st.caption("Couldn't check the G-BA / HAS sheets for this drug — open those views "
                   "from the sidebar to see what's wrong.")
        return
    if hits.empty:
        return
    st.markdown("**Also assessed by other HTA bodies**")
    st.markdown(_cross_agency_lines(hits))
    st.caption("Matched on active substance, not indication. Switch HTA body in the sidebar "
               "for the full record.")
def outcome_order(agency, frame):
    """Outcome categories present in the data, in the agency's scale order."""
    order = AGENCY_CONFIG[agency]["outcome_order"]
    present = set(frame["outcome"].dropna())
    return [o for o in order if o in present] + sorted(present - set(order), key=str.lower)
def init_range_state(key, lo, hi):
    """Default a range slider to its full span, and pull a remembered range back
    inside the data's bounds (the workbook can change under a live session)."""
    current = st.session_state.get(key)
    if not (isinstance(current, (tuple, list)) and len(current) == 2):
        st.session_state[key] = (lo, hi)
        return
    a = max(lo, min(int(current[0]), hi))
    b = max(lo, min(int(current[1]), hi))
    if (min(a, b), max(a, b)) != tuple(current):
        st.session_state[key] = (min(a, b), max(a, b))
def render_agency_switcher():
    counts = {"NICE": TOTAL_ROWS}
    for agency, cfg in AGENCY_CONFIG.items():
        counts[agency] = len(load_agency_sheet(cfg["sheet"], data_version()))
    st.sidebar.title("🌍 HTA body")
    choice = st.sidebar.radio(
        "HTA body",
        list(AGENCY_LABELS),
        format_func=lambda k: f"{AGENCY_LABELS[k]} · {counts.get(k, 0):,}",
        key="hta_body",
        label_visibility="collapsed",
    )
    st.sidebar.caption(
        "NICE weighs cost-effectiveness for the NHS · the G-BA rates added clinical benefit "
        "against a comparator it sets · HAS rates clinical benefit (SMR/ASMR) and decides "
        "early access.")
    # Passage language. Drawn beside the HTA body switch, not inside the
    # citation panel, so it is visible before a record is opened. Only shown
    # when a translations.json actually exists, and never for NICE, which is
    # already English. source_citations reads this through _show_english();
    # the "w_" prefix keeps the choice when switching HTA body.
    if choice in ("G-BA", "HAS") and getattr(
            source_citations, "translations_loaded", lambda: False)():
        st.sidebar.markdown("**\U0001F310 Passage language**")
        st.sidebar.radio(
            "Passage language",
            ["Original", "English"],
            horizontal=True,
            key="w_src_lang_mode",
            label_visibility="collapsed",
        )
        st.sidebar.caption(
            ("German" if choice == "G-BA" else "French") + " is the official "
            "record. English is a machine translation shown for reading only "
            "\u2014 quotes and page links always use the original.")
    st.sidebar.divider()
    return choice
def agency_sidebar_guide(agency):
    """The G-BA / HAS sidebar is a short how-to. All searching and filtering
    lives in the Explorer on the page, so there's one place to look."""
    cfg = AGENCY_CONFIG[agency]
    st.sidebar.markdown("**How to use this page**")
    st.sidebar.markdown(
        f"1. **Search** in the Explorer — a drug, a condition (English or "
        f"{cfg['language']}) and/or tick therapy areas and outcomes.\n"
        f"2. **Pick a section** to read the results: the decisions, analysis, evidence, "
        f"other HTA bodies, or ask AI about them.\n"
        f"3. **Clear** the search to go back to all {cfg['noun']}.")
    st.sidebar.caption("Everything on the page — tables, charts, evidence and the AI chat — "
                       "covers exactly the decisions your search found.")
EXPLORER_FIELDS = ("drug", "condition", "areas", "outcomes", "procedures", "years", "orphan", "evidence")
RESULTS_SECTIONS = ["📋 Decisions", "📊 Analysis", "🧾 Evidence", "🌍 Other HTA bodies",
                    "💬 Ask AI"]
def _explorer_key(slug, field):
    return f"w_{slug}_q_{field}"
def _clear_agency_search(slug):
    """Button callback: back to the full catalogue, every Explorer field blank."""
    for field in EXPLORER_FIELDS:
        st.session_state.pop(_explorer_key(slug, field), None)
    st.session_state.pop(f"{slug}_query", None)
def _counted(counts):
    return lambda option: f"{option} · {counts.get(option, 0):,}"
def render_agency_explorer(agency, frame):
    """
    The search form. Nothing is filtered until Search is pressed, so ticking
    several boxes doesn't rerun the page each time; the submitted search is
    kept in session state and drives every results section below.
    """
    cfg = AGENCY_CONFIG[agency]
    slug = cfg["slug"]
    st.subheader("🔎 HTA Evidence Explorer")
    st.caption(
        f"Search the {len(frame):,} {agency} {cfg['noun']} by drug, condition or therapy area. "
        f"English, {cfg['language']} and abbreviations all work — e.g. "
        f"{cfg['search_examples']}. Leave everything blank to see every {cfg['noun'][:-1]}, "
        f"newest first.")

    area_counts = Counter(a for areas in frame["_areas"] for a in areas)
    area_options = [a for a in AREA_ORDER if area_counts.get(a)]
    outcome_counts = frame["outcome"].value_counts().to_dict()
    outcome_options = outcome_order(agency, frame)
    procedure_counts = frame["procedure_type"].value_counts().to_dict()
    procedure_options = [p for p in cfg["procedure_order"] if procedure_counts.get(p)]
    years = frame["year"].dropna()
    year_bounds = (int(years.min()), int(years.max())) if len(years) else None
    if year_bounds and year_bounds[0] < year_bounds[1]:
        init_range_state(_explorer_key(slug, "years"), *year_bounds)

    with st.form(f"{slug}_explorer_form"):
        c1, c2 = st.columns(2)
        c1.text_input(
            "Drug or active substance", key=_explorer_key(slug, "drug"),
            placeholder="e.g. pembrolizumab or Keytruda",
            help="Optional. INN or brand name; combination regimens match on any of their "
                 "active substances.")
        c2.text_input(
            "Condition or keyword", key=_explorer_key(slug, "condition"),
            placeholder=cfg["condition_placeholder"],
            help=f"Optional. English terms are translated to the {cfg['language']} used in "
                 f"the {agency} records, so 'breast cancer' also finds "
                 + ("'Mammakarzinom'." if agency == "G-BA" else "'cancer du sein'."))
        st.pills(
            "Therapy areas — tick to narrow (none ticked = all areas)", area_options,
            selection_mode="multi", format_func=_counted(area_counts),
            key=_explorer_key(slug, "areas"),
            help="One English vocabulary shared with the NICE and "
                 + ("HAS" if agency == "G-BA" else "G-BA") + " pages. A decision filed under "
                 "several areas counts in each.")
        st.pills(
            "Outcomes — tick to narrow (none ticked = all outcomes)", outcome_options,
            selection_mode="multi", format_func=_counted(outcome_counts),
            key=_explorer_key(slug, "outcomes"))
        st.pills(
            "Procedure type — tick to narrow (none ticked = all)", procedure_options,
            selection_mode="multi", format_func=_counted(procedure_counts),
            key=_explorer_key(slug, "procedures"),
            help=("New indication and re-assessment are read from the procedure title; "
                  "everything else is a standard benefit assessment." if agency == "G-BA"
                  else "HAS records ~700 free-text procedure names; these are the kinds "
                       "of procedure behind them. Tick First listing and Extension of "
                       "indication to leave out routine renewals."))
        _evidence_options = evidence_filter.options_for(agency)
        if _evidence_options:
            st.pills(
                "Evidence available — tick to narrow (none ticked = all)",
                _evidence_options, selection_mode="multi",
                format_func=_counted(evidence_filter.counts_for(agency)),
                key=_explorer_key(slug, "evidence"),
                help=evidence_filter.help_for(agency))
        c3, c4 = st.columns([3, 2])
        with c3:
            if year_bounds and year_bounds[0] < year_bounds[1]:
                st.slider("Decision years", min_value=year_bounds[0],
                          max_value=year_bounds[1], step=1, key=_explorer_key(slug, "years"))
        with c4:
            if agency == "G-BA" and frame["orphan"].notna().any():
                st.session_state.setdefault(_explorer_key(slug, "orphan"), "All")
                st.radio("Orphan drugs", ["All", "Orphan only", "Non-orphan only"],
                         horizontal=True, key=_explorer_key(slug, "orphan"))
        b1, b2, _ = st.columns([1, 1, 4])
        submitted = b1.form_submit_button("🔎 Search", type="primary", width="stretch")
        b2.form_submit_button("Clear", on_click=_clear_agency_search, args=(slug,),
                              width="stretch")

    if submitted:
        st.session_state[f"{slug}_query"] = {
            field: st.session_state.get(_explorer_key(slug, field)) for field in EXPLORER_FIELDS}
        # A new search opens on its decisions, not on the section last viewed.
        st.session_state[f"w_{slug}_section"] = RESULTS_SECTIONS[0]
    return st.session_state.get(f"{slug}_query") or {}
def run_agency_search(agency, frame, query):
    """
    The records a search finds, most relevant first, with _score and _why.

    Filters (areas, outcomes, procedure types, years, orphan status) narrow;
    drug and condition terms must match; qualifiers and recency only rank. A
    drug match counts most, then a condition named in the record's own title,
    indication or category, then one found only in the decision wording or
    reasoning. With no drug or condition the newest decisions come first.
    """
    result = frame
    areas = set(query.get("areas") or [])
    if areas:
        result = result[result["_areas"].map(lambda values: bool(areas & set(values)))]
    if query.get("outcomes"):
        result = result[result["outcome"].isin(query["outcomes"])]
    if query.get("procedures"):
        result = result[result["procedure_type"].isin(query["procedures"])]
    _evidence_ids = evidence_filter.ids_for(agency, query.get("evidence"))
    if _evidence_ids is not None:
        result = result[result["appraisal_id"].astype(str).str.strip()
                        .isin(_evidence_ids)]
    years = query.get("years")
    if years and len(years) == 2 and result["year"].notna().any():
        lo, hi = int(years[0]), int(years[1])
        bounds = (int(frame["year"].min()), int(frame["year"].max()))
        if (lo, hi) != bounds:
            result = result[result["year"].between(lo, hi)]
    orphan = query.get("orphan")
    if orphan == "Orphan only":
        result = result[result["orphan"].eq("Yes")]
    elif orphan == "Non-orphan only":
        result = result[result["orphan"].eq("No")]

    score = pd.Series(0.0, index=result.index)
    why = pd.Series("", index=result.index)
    drug = _fold(query.get("drug"))
    if drug:
        # Ranked by where the drug appears: as the record's own active
        # substance or brand first — for a combination query, more of the named
        # substances ranks higher — then merely mentioned (as a partner or prior
        # therapy) in the indication or decision text.
        keys = molecule_set(query.get("drug"))
        own_share = (result["_molecules"].map(lambda m: len(m & keys) / len(keys))
                     if keys else pd.Series(0.0, index=result.index))
        name_hit = result["_drug_text"].str.contains(_term_pattern(drug), regex=True)
        own = (own_share > 0) | name_hit
        pattern = _any_term_regex([drug] + [k for k in keys if len(k) > 4])
        mentioned = ~own & (result["_primary_text"].str.contains(pattern, regex=True)
                            | result["_secondary_text"].str.contains(pattern, regex=True))
        hit = own | mentioned
        result, score, why = result[hit], score[hit], why[hit]
        own_share, name_hit, own = own_share[hit], name_hit[hit], own[hit]
        score = score + own.map({True: 80, False: 30}) + own_share * 40 + name_hit * 10
        why = why + own.map({True: "Drug", False: "Drug (mentioned)"})
        # A combination query ('nivolumab + ipilimumab') ranks records naming
        # more of its substances anywhere — G-BA often lists only the assessed
        # drug as the substance and the partner in the indication.
        parts = [p.strip() for p in _MOLECULE_SEPARATORS.split(drug) if len(p.strip()) >= 4]
        if len(parts) > 1:
            for part in parts:
                pattern = _term_pattern(_inn_stem(part))
                score = score + (result["_primary_text"].str.contains(pattern, regex=True)
                                 | result["_secondary_text"].str.contains(pattern, regex=True)) * 15

    groups, labels, extra = expand_condition_query(agency, query.get("condition"))
    if groups:
        in_primary = pd.Series(True, index=result.index)
        matched = pd.Series(True, index=result.index)
        for terms in groups:
            pattern = _any_term_regex(terms)
            primary = result["_primary_text"].str.contains(pattern, regex=True)
            secondary = result["_secondary_text"].str.contains(pattern, regex=True)
            matched &= primary | secondary
            in_primary &= primary
        result, score, why = result[matched], score[matched], why[matched]
        in_primary = in_primary[matched]
        score = score + in_primary.map({True: 60, False: 25})
        joiner = why.map(lambda w: " + " if w else "")
        why = why + joiner + in_primary.map({True: "Condition",
                                             False: "Condition (in decision text)"})
    for word in extra:
        pattern = _term_pattern(word)
        score = score + result["_primary_text"].str.contains(pattern, regex=True) * 5 \
                      + result["_secondary_text"].str.contains(pattern, regex=True) * 2

    dated = result["decision_dt"]
    if dated.notna().any():
        lo, hi = dated.min(), dated.max()
        span = max((hi - lo).days, 1)
        score = score + ((dated - lo).dt.days / span * 10).fillna(0)
    result = result.assign(_score=score, _why=why)
    return result.sort_values(["_score", "decision_dt"], ascending=[False, False],
                              na_position="last", kind="stable"), labels
def _query_key(query):
    """A hashable, order-stable form of the submitted search."""
    key = []
    for field in EXPLORER_FIELDS:
        value = query.get(field)
        if isinstance(value, (list, tuple)):
            value = tuple(value)
        elif isinstance(value, str):
            value = value.strip() or None
        key.append((field, value))
    return tuple(key)
@st.cache_data(max_entries=64)
def _search_ranking(agency, version, query_key):
    """Cached result of one search: (row index, score, why, condition labels).
    Only the ranking is cached, not the rows, so a hit costs a slice of the
    already-cached frame. A condition search over the 12k HAS records takes up
    to ~1s, and without this it would repeat on every click in the results."""
    frame = prepare_agency_frame(agency, version)
    ranked, labels = run_agency_search(agency, frame, dict(query_key))
    return (list(ranked.index), ranked["_score"].tolist(), ranked["_why"].tolist(), labels)
def cached_agency_search(agency, version, query):
    frame = prepare_agency_frame(agency, version)
    index, score, why, labels = _search_ranking(agency, version, _query_key(query))
    results = frame.loc[index].assign(_score=score, _why=why)
    return results, labels
def describe_query(agency, query, labels, frame=None):
    """One line saying what was searched, including how the condition was read."""
    parts = []
    if _text(query.get("drug")):
        parts.append(f"drug **{_md_safe(query['drug'].strip())}**")
    if _text(query.get("condition")):
        read_as = (f" (read as: {', '.join(labels)})" if labels else "")
        parts.append(f"condition **{_md_safe(query['condition'].strip())}**{read_as}")
    for field, label in (("areas", "areas"), ("outcomes", "outcomes"),
                         ("evidence", "evidence"),
                         ("procedures", "procedure types")):
        if query.get(field):
            parts.append(f"{label}: {', '.join(query[field])}")
    years = query.get("years")
    full = ((int(frame["year"].min()), int(frame["year"].max()))
            if frame is not None and frame["year"].notna().any() else None)
    if years and len(years) == 2 and tuple(int(y) for y in years) != full:
        parts.append(f"years {years[0]}–{years[1]}")
    if query.get("orphan") and query["orphan"] != "All":
        parts.append(query["orphan"].lower())
    return " · ".join(parts)
def render_agency_record(agency, row, docs, hits):
    cfg = AGENCY_CONFIG[agency]
    color = cfg["outcome_md_colors"].get(row["outcome"], "gray")
    label = _md_safe(str(row["outcome"]).replace("[", "(").replace("]", ")"))
    st.markdown(f"**Outcome:** :{color}[**{label}**]")
    recorded = _text(row.get("decision_simple"))
    if recorded and recorded != row["outcome"]:
        st.caption(f"Recorded as: {_md_safe(recorded)}")

    meta = [f"**Decision date:** {row['date_label']}"]
    published = _parse_date(row.get("publication_date"))
    if not pd.isna(published) and published != _parse_date(row.get("decision_date")):
        meta.append(f"**Published:** {published:%d %b %Y}")
    meta.append(f"**Procedure:** {_md_safe(row.get('procedure_type') or 'Not recorded')}")
    if row.get("orphan") == "Yes":
        meta.append("**Orphan drug**")
    st.markdown(" · ".join(meta))
    areas = ", ".join(row.get("_areas") or ()) or row.get("therapy_area_en")
    condition = _text(row.get("condition"))
    indication = _text(row.get("indication"))
    if indication == "Indication not recorded":
        indication = None
    # Where the table's condition label is just the indication trimmed, the
    # full indication below says it better — don't repeat it.
    if condition and indication and indication.startswith(condition.rstrip("…").rstrip()):
        condition = None
    st.markdown((f"**Condition:** {_md_safe(condition)} · " if condition else "")
                + f"**Therapy area:** {_md_safe(areas)}")
    assessment_type = _text(row.get("assessment_type"))
    if assessment_type and assessment_type.rstrip(";") != row.get("procedure_type"):
        st.caption(f"Procedure as recorded: {_md_safe(_shorten(assessment_type, 200))}")
    if indication:
        st.markdown(f"**Indication (as recorded):** {_md_safe(_shorten(source_citations.display_text(indication), 700))}")

    for label, col in cfg["detail_fields"]:
        value = _text(row.get(col))
        if value:
            st.markdown(f"**{label}:** "
                        f"{_md_safe(source_citations.display_text(value))}")
    icer = format_eur_icer(row)
    if icer:
        st.markdown(f"**Published ICER:** {icer}")
    note = _text(row.get("methodology_note"))
    if note:
        st.caption(_md_safe(note))

    # Silent for anything without coded attributes: every non-NICE
    # record, and the NICE appraisals with no committee discussion
    # indexed.
    nice_attributes.render_panel(row.get("appraisal_id"))

    # Cited source passages, with a link that opens the official PDF at the
    # page the text came from. Silent for NICE and for any record with no
    # indexed passages, so no guard is needed here.
    # Verdicts first, passages underneath: a reader wants the conclusion and
    # its page before the wall of source text.
    endpoint_panel.render(agency, row.get("appraisal_id"))

    source_citations.render_panel(agency, row.get("appraisal_id"))

    record_docs = docs[docs["Appraisal_ID"] == str(row["appraisal_id"])]
    def _doc_date(value):
        when = _parse_date(value)
        return f" · {when:%d %b %Y}" if not pd.isna(when) else (
            f" · {_md_safe(value)}" if _text(value) else "")

    links = [
        f"- {_md_link(d['Document_Type'], d['Source_URL'])}{_doc_date(d['Date'])}"
        for _, d in record_docs.iterrows() if _text(d["Source_URL"])
    ]
    if links:
        st.markdown("**Source documents**")
        st.markdown("\n".join(links))
    elif _text(row.get("official_source_url")):
        st.markdown(_md_link(f"Open the official {agency} decision ↗", row["official_source_url"]))

    if hits is not None and not hits.empty:
        st.markdown("**Same molecule at other HTA bodies**")
        st.markdown(_cross_agency_lines(hits))
def _block_subset(frame, subset):
    if subset == "favourable":
        return frame[frame["favourable"].eq(True)]
    if subset == "unfavourable":
        return frame[frame["favourable"].eq(False)]
    return frame
def render_synthesis_block(block, frame):
    """
    First substantive text among block['cols'], per row, for whichever subset
    of the frame this block covers.

    Row-by-row .iterrows() here used to cost ~1-1.5s per block at the full
    HAS catalogue (12k+ rows) — negligible at the few hundred rows this was
    built against, but five such blocks add seconds to every single rerun
    once the sheet is full-sized. Finding the first non-blank column is a
    columnwise .map() plus a fill-forward across columns in priority order,
    which is the same "first truthy in this order" logic without visiting
    rows in Python; only the rows actually kept (capped at
    SYNTHESIS_MAX_ITEMS for display) are ever turned into row objects.
    """
    subset = _block_subset(frame, block.get("subset", "all"))
    texts = pd.Series(None, index=subset.index, dtype=object)
    for c in block["cols"]:
        if c not in subset.columns:
            continue
        col_text = subset[c].map(_substantive)
        texts = texts.where(texts.notna(), col_text)
    mask = texts.notna()
    total = int(mask.sum())
    if not total:
        if block.get("empty"):
            st.markdown(f"**{block['title']}**")
            st.caption(block["empty"])
        return
    matched = subset.loc[mask]
    matched_texts = texts.loc[mask]
    with st.expander(f"{block['title']} ({total})", expanded=total <= 5):
        if block.get("caption"):
            st.caption(block["caption"])
        for idx in matched.index[:SYNTHESIS_MAX_ITEMS]:
            row = matched.loc[idx]
            text = matched_texts.loc[idx]
            about = _text(row.get("condition")) or row["indication"]
            st.markdown(f"**{_md_safe(row['appraisal_id'])} — {_md_safe(row['inn'])}** · "
                        f"{row['outcome']} · _{_md_safe(_shorten(about, 90))}_\n\n"
                        f"> {_md_safe(_shorten(text, 450))}")
        if total > SYNTHESIS_MAX_ITEMS:
            st.caption(f"Showing the {SYNTHESIS_MAX_ITEMS} most relevant of {total:,} — narrow "
                       f"your search in the Explorer to see a specific set.")
def render_economic_block(frame):
    st.markdown("**Economic evaluation**")
    with_icer = frame[frame["icer_eur_lower"].notna()]
    if len(with_icer):
        table = pd.DataFrame({
            "ID": with_icer["appraisal_id"],
            "Drug (INN)": with_icer["inn"],
            "Indication": with_icer["indication"],
            "ICER (€/QALY)": with_icer.apply(format_eur_icer, axis=1),
            "Outcome": with_icer["outcome"],
        })
        st.dataframe(table, width="stretch", hide_index=True)
        st.caption("Shown in €/QALY as published — not converted to £ or compared with NICE's "
                   "threshold. Neither France nor Germany applies an explicit cost-effectiveness "
                   "threshold.")
        return
    statuses = (frame["ICER_status"].map(_text).dropna().value_counts()
                if "ICER_status" in frame.columns else pd.Series(dtype=int))
    if len(statuses):
        st.caption("No ICER in the current selection. Recorded status: " + "; ".join(
            f"'{s}' ({n})" for s, n in statuses.items()) + ".")
    else:
        st.caption("No ICER in the current selection.")
def _coverage_line(frame, docs):
    n = len(frame)

    def count(col):
        return int(frame[col].map(_substantive).notna().sum()) if col in frame.columns else 0

    with_docs = int(frame["appraisal_id"].astype(str).isin(set(docs["Appraisal_ID"])).sum())
    return (
        f"Coverage in this selection — reasoning {count('decision_reasoning')}/{n} · "
        f"comparator {count('comparator')}/{n} · patient population "
        f"{count('patient_population')}/{n} · source documents {with_docs}/{n} · "
        f"published ICER {int(frame['icer_eur_lower'].notna().sum())}/{n}")
def _integer_axis(fig, axis, max_value):
    """Whole-number ticks — counts of 1–5 otherwise get 0.5 steps."""
    update = fig.update_yaxes if axis == "y" else fig.update_xaxes
    update(rangemode="tozero", tickformat=",d", dtick=1 if max_value <= 8 else None)
CHART_OPTIONS = {
    "G-BA": ["Outcome breakdown (table)", "Outcome mix", "Decisions over time",
             "By therapy area", "Certainty of evidence"],
    "HAS": ["Outcome breakdown (table)", "Outcome mix", "Decisions over time",
            "By therapy area", "SMR & ASMR ratings"],
}
CHART_DEFAULTS = ["Outcome breakdown (table)", "Outcome mix"]
MAX_AREA_BARS = 12
def _style_bars(fig, height):
    """Thin bars with rounded data ends, recessive chrome, legend above."""
    fig.update_layout(height=height, margin={"l": 8, "r": 8, "t": 36, "b": 8},
                      legend={"orientation": "h", "yanchor": "bottom", "y": 1.02,
                              "xanchor": "left", "x": 0, "title": None},
                      barcornerradius=4)
    return fig
def outcome_breakdown_table(agency, frame):
    """Every outcome on the agency's full scale, with its share — the readable
    form for a 12–13-class scale, and the table view behind the charts."""
    counts = frame["outcome"].value_counts()
    order = [o for o in outcome_order(agency, frame) if counts.get(o)]
    n = len(frame)
    return pd.DataFrame({
        "Outcome": order,
        "Group": [outcome_family(agency, o) for o in order],
        "Decisions": [int(counts[o]) for o in order],
        "Share": [counts[o] / n * 100 if n else 0 for o in order],
    })
def render_outcome_table(agency, frame, key):
    table = outcome_breakdown_table(agency, frame)
    st.dataframe(
        table,
        column_config={
            "Decisions": st.column_config.NumberColumn("Decisions", format="%d"),
            "Share": st.column_config.ProgressColumn("Share", format="%.0f%%", min_value=0,
                                                     max_value=100),
        },
        hide_index=True, width="stretch", height=35 * (len(table) + 1) + 3, key=key)
def render_agency_charts(agency, frame):
    """The Analysis section: a checklist of views, each a fixed, readable size."""
    cfg = AGENCY_CONFIG[agency]
    slug = cfg["slug"]
    mode = _theme_mode()
    colors = FAMILY_COLORS[mode]
    families = family_order(agency)
    key = f"w_{slug}_charts"
    st.session_state.setdefault(key, CHART_DEFAULTS)
    shown = st.pills("Show — tick the views you want", CHART_OPTIONS[agency],
                     selection_mode="multi", key=key) or []
    if not shown:
        st.caption("Tick a view above to show it.")
        return
    fam = frame["outcome_family"]

    if "Outcome breakdown (table)" in shown:
        st.markdown("**Outcome breakdown**")
        render_outcome_table(agency, frame, key=f"{slug}_analysis_outcomes")
        st.caption("Group is how the outcome is coloured in the charts below.")

    if "Outcome mix" in shown:
        st.markdown("**Outcome mix**")
        counts = fam.value_counts().reindex(families).fillna(0).astype(int)
        counts = counts[counts > 0]
        data = pd.DataFrame({"Group": counts.index, "Decisions": counts.to_numpy()})
        data["label"] = [f"{v:,}  ({v / len(frame) * 100:.0f}%)" for v in data["Decisions"]]
        fig = px.bar(data, x="Decisions", y="Group", orientation="h", color="Group",
                     color_discrete_map=colors, text="label",
                     category_orders={"Group": families})
        fig.update_traces(textposition="outside", cliponaxis=False, showlegend=False,
                          hovertemplate="%{y}: %{x:,}<extra></extra>")
        fig.update_layout(xaxis_title=None, yaxis_title=None, bargap=0.45,
                          xaxis={"range": [0, data["Decisions"].max() * 1.25]})
        _integer_axis(fig, "x", data["Decisions"].max())
        st.plotly_chart(_style_bars(fig, 44 * len(data) + 70), width="stretch",
                        key=f"{slug}_mix")

    if "Decisions over time" in shown:
        st.markdown(f"**{cfg['noun_title']} over time**")
        dated = frame.dropna(subset=["decision_dt"])
        if dated.empty:
            st.caption("No decision dates recorded in this selection.")
        else:
            by_year = (dated.assign(Year=dated["decision_dt"].dt.year.astype(int),
                                    Group=dated["outcome_family"])
                       .groupby(["Year", "Group"]).size().reset_index(name="Decisions"))
            fig = px.bar(by_year, x="Year", y="Decisions", color="Group",
                         color_discrete_map=colors, category_orders={"Group": families})
            fig.update_layout(barmode="stack", xaxis_title=None, yaxis_title=None,
                              bargap=0.35, xaxis={"dtick": 1 if by_year["Year"].nunique() <= 12
                                                  else 2})
            _integer_axis(fig, "y", by_year.groupby("Year")["Decisions"].sum().max())
            st.plotly_chart(_style_bars(fig, 340), width="stretch", key=f"{slug}_time")

    if "By therapy area" in shown:
        st.markdown("**By therapy area**")
        # Every area a decision is filed under, so bar lengths match the counts
        # the Explorer's therapy-area checklist filters on.
        exploded = (frame.assign(Area=frame["_areas"].map(list), Group=fam)
                    [["Area", "Group"]].explode("Area"))
        area_totals = exploded["Area"].value_counts()
        kept = list(area_totals.head(MAX_AREA_BARS).index)
        exploded["Area"] = exploded["Area"].where(exploded["Area"].isin(kept), "Other areas")
        by_area = exploded.groupby(["Area", "Group"]).size().reset_index(name="Decisions")
        fig = px.bar(by_area, x="Decisions", y="Area", color="Group", orientation="h",
                     color_discrete_map=colors, category_orders={"Group": families})
        fig.update_layout(barmode="stack", yaxis={"categoryorder": "total ascending",
                                                  "title": None},
                          xaxis_title=None, bargap=0.4)
        _integer_axis(fig, "x", by_area.groupby("Area")["Decisions"].sum().max())
        st.plotly_chart(_style_bars(fig, 30 * by_area["Area"].nunique() + 110),
                        width="stretch", key=f"{slug}_areas")
        note = ("A decision filed under several areas counts in each, as it does in the "
                "Explorer's checklist.")
        if len(area_totals) > MAX_AREA_BARS:
            note = (f"The {MAX_AREA_BARS} largest of {len(area_totals)} therapy areas; the rest "
                    f"are grouped as 'Other areas'. " + note)
        st.caption(note)

    if "Certainty of evidence" in shown and agency == "G-BA":
        st.markdown("**Positive ratings by certainty of evidence**")
        rated = frame[frame["certainty"].notna() & frame["favourable"].eq(True)]
        if rated.empty:
            st.caption("No positive rating with a stated certainty in this selection.")
        else:
            levels = [c for c in ("Proof", "Indication", "Hint") if c in set(rated["certainty"])]
            table = pd.crosstab(rated["outcome"], rated["certainty"]).reindex(columns=levels)
            table = table.reindex([o for o in outcome_order(agency, frame) if o in table.index])
            st.dataframe(table.rename_axis(index="Outcome", columns=None),
                         width=150 + 110 * len(levels) + 180)
            st.caption("For mixed resolutions, certainty refers to the subgroup that received "
                       "a positive rating.")

    if "SMR & ASMR ratings" in shown and agency == "HAS":
        smr = frame[frame["_smr"].notna()]
        asmr = frame[frame["_asmr"].notna()]
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**SMR (actual clinical benefit)**")
            if len(smr):
                counts = smr["_smr"].value_counts().reindex(
                    [s for s in SMR_ORDER if s in set(smr["_smr"])]
                    + sorted(set(smr["_smr"]) - set(SMR_ORDER))).reset_index()
                counts.columns = ["SMR", "Decisions"]
                fig = px.bar(counts, x="SMR", y="Decisions", text="Decisions",
                             color_discrete_sequence=[SINGLE_SERIES[mode]])
                fig.update_traces(textposition="outside", cliponaxis=False,
                                  texttemplate="%{text:,}")
                fig.update_layout(xaxis_title=None, yaxis_title=None, bargap=0.45)
                _integer_axis(fig, "y", counts["Decisions"].max())
                st.plotly_chart(_style_bars(fig, 300), width="stretch", key=f"{slug}_smr")
            else:
                st.caption("No SMR ratings in this selection.")
        with c2:
            st.markdown("**ASMR (improvement in clinical benefit)**")
            if len(asmr):
                levels = [lv for lv in _ASMR_ORDER if lv in set(asmr["_asmr"])]
                counts = asmr["_asmr"].value_counts().reindex(levels).reset_index()
                counts.columns = ["ASMR", "Decisions"]
                counts["ASMR"] = "ASMR " + counts["ASMR"]
                ramp = dict(zip(["ASMR I", "ASMR II", "ASMR III", "ASMR IV"], ASMR_RAMP[mode]))
                ramp["ASMR V"] = NEUTRAL_GREY[mode]
                fig = px.bar(counts, x="ASMR", y="Decisions", color="ASMR", text="Decisions",
                             color_discrete_map=ramp)
                fig.update_traces(textposition="outside", cliponaxis=False,
                                  texttemplate="%{text:,}", showlegend=False)
                fig.update_layout(xaxis_title=None, yaxis_title=None, bargap=0.45)
                _integer_axis(fig, "y", counts["Decisions"].max())
                st.plotly_chart(_style_bars(fig, 300), width="stretch", key=f"{slug}_asmr")
                st.caption("ASMR I = major improvement … V = none (grey).")
            else:
                st.caption("No ASMR ratings in this selection.")
def _agency_footer(agency, total):
    cfg = AGENCY_CONFIG[agency]
    st.divider()
    st.caption(
        f"Built with Python & Streamlit | {total:,} {cfg['noun']} sourced from "
        f"{cfg['source_label']} | Preliminary intelligence tool — not a substitute for full "
        f"evidence review or professional market access advice")
def _records_matched(matches, body=None):
    """How many records (rows, not IDs) have at least one match, optionally at one body."""
    subset = matches if body is None else matches[matches["body"] == body]
    return int(subset["_src"].nunique())
@st.cache_data
def _agency_export_bytes(agency, version, index_tuple):
    """
    The .xlsx download for the current filter selection.

    Writing this with openpyxl takes ~9-10s at the full HAS catalogue (12k+
    rows, ~40 columns) — fine once, but the call site used to rebuild it on
    every rerun, including ones that have nothing to do with the filters
    (opening a detail expander, asking the chat a question, switching to
    another agency tab and back). Keying the cache on the filtered row
    index means it's only rebuilt when the actual selection changes, and
    re-fetching prepare_agency_frame here is cheap — it's already cached by
    (agency, version) — so this doesn't need the (large, slower-to-hash)
    filtered frame passed in directly.
    """
    frame = prepare_agency_frame(agency, version)
    subset = frame.loc[list(index_tuple)]
    export = subset.drop(columns=[c for c in subset.columns
                                  if c.startswith("_") or c == "search_blob"])
    buffer = io.BytesIO()
    export.to_excel(buffer, index=False)
    return buffer.getvalue()
def _supports_deferred_download():
    """Streamlit 1.49+ can build a download's bytes only when it's clicked."""
    try:
        import streamlit.elements.widgets.button as _button
        return "Callable" in str(getattr(_button, "DownloadButtonDataType", ""))
    except Exception:
        return False
DEFERRED_DOWNLOADS = _supports_deferred_download()
EAGER_EXPORT_MAX_ROWS = 1500    # older Streamlit: above this, build the file on request
def render_download(agency, version, results):
    """Excel export of exactly the search results, in ranked order. At the full
    HAS catalogue the file takes ~10s to write, so it's never built just to
    draw the page: on click where Streamlit supports it, else on request."""
    cfg = AGENCY_CONFIG[agency]
    slug = cfg["slug"]
    index = tuple(results.index)
    label = f"📥 Download these {len(results):,} as Excel"
    kwargs = {"file_name": f"{slug}_search_results.xlsx", "mime": XLSX_MIME,
              "key": f"download_{slug}"}
    if DEFERRED_DOWNLOADS:
        st.download_button(label, data=lambda: _agency_export_bytes(agency, version, index),
                           **kwargs)
    elif len(results) <= EAGER_EXPORT_MAX_ROWS or st.session_state.get(f"{slug}_export") == index:
        st.download_button(label, data=_agency_export_bytes(agency, version, index), **kwargs)
    elif st.button(f"📥 Prepare Excel file ({len(results):,} rows)", key=f"{slug}_prepare"):
        st.session_state[f"{slug}_export"] = index
        st.rerun()
def _is_narrowed(frame, query):
    years = query.get("years")
    bounds = ((int(frame["year"].min()), int(frame["year"].max()))
              if frame["year"].notna().any() else None)
    return bool(_text(query.get("drug")) or _text(query.get("condition"))
                or query.get("areas") or query.get("outcomes") or query.get("procedures")
                or query.get("orphan") not in (None, "All")
                or (years and bounds and tuple(int(y) for y in years) != bounds))
def render_results_header(agency, frame, results, query, labels, matches):
    cfg = AGENCY_CONFIG[agency]
    n, total = len(results), len(frame)
    if _is_narrowed(frame, query):
        st.markdown(f"### {n:,} of {total:,} {cfg['noun']} found")
        split = ""
        if _text(query.get("drug")) and n:
            mentioned = int(results["_why"].str.startswith("Drug (mentioned)").sum())
            if mentioned:
                split = (f" {n - mentioned:,} are of the drug itself; {mentioned:,} more mention "
                         f"it — as the comparator, a combination partner or a prior therapy "
                         f"(see 'Matched on').")
        st.caption(f"Searched {describe_query(agency, query, labels, frame)}. "
                   + ("Most relevant first." if (_text(query.get("drug"))
                                                 or _text(query.get("condition")))
                      else "Newest first.") + split)
    else:
        st.markdown(f"### All {total:,} {cfg['noun']}, newest first")
        st.caption("No search yet — use the Explorer above to narrow this down.")
    if not n:
        return
    tiles = [(f"{cfg['noun_title']} found", f"{n:,}", None)] + cfg["headlines_fn"](results)
    # Four tiles fit the page width without truncating; the NICE overlap has its
    # own section when an agency already fills the row.
    if len(tiles) < 4:
        tiles.append(("Same molecule at NICE", f"{_records_matched(matches, 'NICE'):,}/{n:,}",
                      "Results whose active substance also appears in the NICE data — see "
                      "'Other HTA bodies'."))
    for col, (label, value, help_text) in zip(st.columns(len(tiles)), tiles):
        col.metric(label, value, help=help_text)
def _section_nav(slug):
    key = f"w_{slug}_section"
    if st.session_state.get(key) not in RESULTS_SECTIONS:
        st.session_state[key] = RESULTS_SECTIONS[0]
    if "required" in inspect.signature(st.segmented_control).parameters:
        choice = st.segmented_control("Results section", RESULTS_SECTIONS, key=key,
                                      required=True, label_visibility="collapsed")
    else:
        choice = st.radio("Results section", RESULTS_SECTIONS, key=key, horizontal=True,
                          label_visibility="collapsed")
    return choice or RESULTS_SECTIONS[0]
def _results_signature(results):
    """Changes whenever the search does, so a stale row selection never carries over."""
    return abs(hash((len(results), tuple(results.index[:200])))) % 10**10
def render_results_decisions(agency, version, results, docs, matches, query):
    cfg = AGENCY_CONFIG[agency]
    slug = cfg["slug"]
    top, bottom = st.columns([3, 2])
    with top:
        st.caption("Click a row to open its full record below.")
    with bottom:
        render_download(agency, version, results)

    table = pd.DataFrame(index=results.index)
    if _text(query.get("drug")) or _text(query.get("condition")):
        table["Matched on"] = results["_why"]
    table["ID"] = results["appraisal_id"]
    table["Drug"] = results["inn"]
    table["Brand"] = results["brand"]
    table["Outcome"] = results["outcome"]
    table["Decided"] = results["decision_dt"]
    table["Condition"] = results["condition"]
    table["Therapy area"] = results["therapy_area_en"]
    if agency == "G-BA":
        table["Certainty"] = results["certainty"].fillna("—")
    else:
        table["SMR"] = results["SMR_rating"].map(lambda v: _shorten(_text(v), 40)
                                                 if _text(v) else "—")
        table["ASMR"] = results["ASMR_rating"].map(lambda v: _shorten(_text(v), 30)
                                                   if _text(v) else "—")
        table["Procedure"] = results["procedure_type"]
    table["Source"] = results["official_source_url"]
    event = st.dataframe(
        table.reset_index(drop=True),
        column_config={
            "Decided": st.column_config.DateColumn("Decided", format="D MMM YYYY"),
            "Condition": st.column_config.TextColumn("Condition", width="medium"),
            "Source": st.column_config.LinkColumn("Source", display_text="Open ↗"),
        },
        hide_index=True, width="stretch", height=420, on_select="rerun",
        selection_mode="single-row", key=f"{slug}_results_{_results_signature(results)}")

    rows = []
    try:
        rows = list(event.selection.rows)
    except Exception:
        pass
    position = rows[0] if rows and rows[0] < len(results) else 0
    row = results.iloc[position]
    st.divider()
    st.markdown(f"#### {_md_safe(row['appraisal_id'])} · {_md_safe(row['inn'])}"
                + (f" ({_md_safe(row['brand'])})" if row["brand"] else ""))
    if not rows:
        st.caption("Showing the top result — click any row in the table to open another.")
    by_record = _matches_by_record(matches[matches["_src"] == row.name])
    render_agency_record(agency, row, docs, by_record.get(row.name))
def render_results_evidence(agency, results, docs):
    cfg = AGENCY_CONFIG[agency]
    st.caption(f"What the {len(results):,} {cfg['noun']} your search found say — up to "
               f"{SYNTHESIS_MAX_ITEMS} per heading, most relevant first. "
               + _coverage_line(results, docs) + ".")
    for block in cfg["synthesis_blocks"]:
        render_synthesis_block(block, results)
    render_economic_block(results)
def render_results_other_bodies(agency, results, matches):
    """Same molecule at NICE and the other agency. Widely appraised molecules
    (pembrolizumab has dozens of NICE TAs) would swamp the table, so by default
    each result shows only its closest-indication matches."""
    cfg = AGENCY_CONFIG[agency]
    per_record = 3
    show_all = False
    if len(matches) and matches.groupby("_src").size().max() > per_record:
        show_all = st.checkbox(
            f"Show every same-molecule match (default: the {per_record} closest indications "
            f"per result)", key=f"w_{cfg['slug']}_all_matches")
    shown = matches if show_all else matches.groupby("_src", sort=False).head(per_record)
    if not len(shown):
        st.info("None of these molecules appear at another HTA body in this workbook yet.")
        return
    source = results.loc[shown["_src"]]
    comparison = pd.DataFrame({
        "Molecule": source["inn"].to_numpy(),
        f"{agency} ID": source["appraisal_id"].to_numpy(),
        f"{agency} outcome": source["outcome"].to_numpy(),
        "Other body": [f"{AGENCY_FLAGS.get(b, '')} {b}" for b in shown["body"]],
        "Other ID": shown["appraisal_id"].to_numpy(),
        "Other decision": shown["decision"].to_numpy(),
        "Other drug / regimen": shown["drug"].to_numpy(),
        "Other indication": shown["indication"].to_numpy(),
        "When": shown["when"].to_numpy(),
        "Link": shown["url"].to_numpy(),
    })
    st.markdown(f"**{_records_matched(matches):,} of {len(results):,}** results have a "
                f"decision on the same active substance at another HTA body.")
    st.dataframe(comparison,
                 column_config={"Link": st.column_config.LinkColumn("Link", display_text="Open ↗")},
                 width="stretch", hide_index=True)
    st.caption(CROSS_AGENCY_CAVEAT)
def render_agency_view(agency):
    cfg = AGENCY_CONFIG[agency]
    slug = cfg["slug"]
    version = data_version()
    frame = prepare_agency_frame(agency, version)

    st.title(cfg["title"])
    if frame.empty:
        st.info(f"No {agency} records in this workbook yet — add rows to the "
                f"'{cfg['sheet']}' sheet and they'll appear here with no code changes.")
        return

    doc_index = load_document_index(version)
    # Index the three page-cited passage sheets for this run. Each is absent
    # from the 6-sheet build of the workbook, in which case load_agency_sheet
    # returns an empty frame and the citation panel renders nothing.
    source_citations.attach(
        load_agency_sheet("GBA_Decision_Passages", version),
        load_agency_sheet("HAS_Economic_Pages", version),
        load_agency_sheet("Economic_Findings", version),
    )
    # The endpoint panel needs the AIS subgroup rows and the same passages,
    # which it parses for each resolution's endpoint-category summary table.
    endpoint_panel.attach(
        load_agency_sheet("GBA_Subgroups", version),
        load_agency_sheet("GBA_Decision_Passages", version),
    )
    # Which appraisals had assessable evidence at all. G-BA reads this
    # from the subgroup verdicts; HAS from whether an SMR grade exists.
    evidence_filter.attach_gba(load_agency_sheet("GBA_Subgroups", version)
                               if agency == "G-BA" else None)
    evidence_filter.attach_has(load_agency_sheet("HAS_Ratings", version)
                               if agency == "HAS" else None)
    docs = doc_index[doc_index["Institution"] == agency]
    total = len(frame)
    all_matches = agency_match_table(agency, version)

    st.markdown(f"*{_md_safe(cfg['subtitle_fn'](frame))} — {_date_span(frame['decision_dt'])}*")
    if total < SMALL_AGENCY_SET:
        st.info(f"Small dataset so far — {total} {cfg['noun']}. Counts, charts and patterns "
                f"below are illustrative until more {agency} decisions are loaded.")
    with st.expander("Scope & methodology — read before comparing with NICE"):
        st.markdown(cfg["methodology_md"])
        documented = int(frame["appraisal_id"].astype(str).isin(set(docs["Appraisal_ID"])).sum())
        st.markdown(
            f"**Coverage.** {total:,} {cfg['noun']} dated {_date_span(frame['decision_dt'])} · "
            f"{len(docs):,} official source documents indexed for {documented:,} of them · the "
            f"same active substance also appears in the NICE data for "
            f"{_records_matched(all_matches, 'NICE'):,} of {total:,}.")
        for note in frame["methodology_note"].map(_text).dropna().unique():
            st.caption(f"Workbook methodology note: {_md_safe(note)}")
        log = load_enrichment_log(data_version())
        if log is not None and "Area" in log.columns:
            area = log["Area"].astype(str)
            rows = log[area.str.contains(cfg["log_keyword"], case=False, regex=False)
                       | area.str.contains("document index", case=False, regex=False)]
            if len(rows):
                st.markdown("**Enrichment log**")
                st.dataframe(rows.astype(str), width="stretch", hide_index=True)

    agency_sidebar_guide(agency)
    query = render_agency_explorer(agency, frame)
    results, labels = cached_agency_search(agency, version, query)
    matches = all_matches[all_matches["_src"].isin(results.index)]

    st.divider()
    render_results_header(agency, frame, results, query, labels, matches)
    if results.empty:
        st.warning("Nothing matches that search. Try a broader condition (e.g. 'lymphoma' "
                   "rather than a subtype), fewer ticked boxes, or a wider year range.")
        _agency_footer(agency, total)
        return

    section = _section_nav(slug)
    if section == "📋 Decisions":
        render_results_decisions(agency, version, results, docs, matches, query)
    elif section == "📊 Analysis":
        render_agency_charts(agency, results)
    elif section == "🧾 Evidence":
        render_results_evidence(agency, results, docs)
    elif section == "🌍 Other HTA bodies":
        render_results_other_bodies(agency, results, matches)
    else:
        render_agency_chat(agency, results, matches)
    _agency_footer(agency, total)
