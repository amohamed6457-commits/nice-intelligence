"""
NICE Technology Appraisal Intelligence — dashboard, with G-BA and HAS views.

Data contract: NICE_v14_updated_2026-09-28.xlsx
  Sheet1               — NICE appraisal rows (adds year_start, year_label, search_blob)
  G-BA_Recent          — G-BA AMNOG added-benefit resolutions (Germany)
  HAS_Recent           — HAS early-access / Transparency Committee decisions (France)
  Appraisal_Documents  — official source documents for every agency, keyed by ID
  Tag_Vocabulary       — single source of truth for every categorical dropdown
  Enrichment_Log       — provenance, surfaced in the Methodology expanders

The sidebar's "HTA body" switch picks the view. The NICE view is the original
dashboard; the G-BA and HAS views are driven by AGENCY_CONFIG, so adding rows
to their sheets needs no code changes.
"""

import io
import os
import re
import warnings
import requests
from collections import Counter
from datetime import date, datetime

import pandas as pd
import plotly.express as px
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    HRFlowable,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

DATA_FILE = "NICE_v14_updated_2026-09-28.xlsx"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def data_version():
    """
    Cache key for the workbook: its path and last-modified time. Every loader
    below takes this as an argument, so replacing the workbook — or pointing
    DATA_FILE at a new one — refreshes every view on the next run instead of
    serving rows cached from the old file.
    """
    try:
        return (DATA_FILE, os.path.getmtime(DATA_FILE))
    except OSError:
        return (DATA_FILE, None)


@st.cache_resource
def _workbook(version):
    """
    Open the workbook once per (path, mtime) and hand back the same ExcelFile
    for every sheet. Each independent pd.read_excel(path, sheet_name=...) call
    re-parses the *entire* file — harmless at a few hundred rows, but at the
    full HAS/G-BA catalogue (12k+ and 1k+ rows, a 40k-row document index) six
    separate full-file parses is where a cold load stops feeling instant.
    Reading every sheet off one shared handle turns six parses of the whole
    file into one open plus six cheap per-sheet parses.

    python-calamine (if installed) is roughly an order of magnitude faster
    at opening a large .xlsx than the default openpyxl engine; if it isn't
    available, this falls back to the default engine and still gets the
    single-open benefit.
    """
    path = version[0]
    try:
        return pd.ExcelFile(path, engine="calamine")
    except (ImportError, ValueError):
        return pd.ExcelFile(path)


def _read_sheet(version, sheet_name):
    return _workbook(version).parse(sheet_name)

st.set_page_config(
    page_title="NICE Intelligence Dashboard",
    page_icon="💊",
    layout="wide",
)

# Keep filters and Explorer inputs when switching HTA body. Streamlit deletes a
# widget's value at the end of any run in which that widget wasn't drawn, so
# NICE → G-BA → NICE would otherwise wipe every NICE input — and, because
# hta_query_ran survives, land on a bare "Please enter an indication" warning
# instead of the results that were on screen. Re-saving keyed values at the top
# of each run is Streamlit's documented way to keep them. Only value widgets use
# the "w_" key prefix: buttons and chat inputs can't be written this way. Their
# defaults are set through session state rather than value=/index= arguments,
# which is what lets this run without widget-state warnings.
for _state_key in [k for k in st.session_state.keys() if str(k).startswith("w_")]:
    st.session_state[_state_key] = st.session_state[_state_key]


# Data loading

@st.cache_data
def load_data(version):
    return _read_sheet(version, "Sheet1")


@st.cache_data
def load_enrichment_log(version):
    try:
        return _read_sheet(version, "Enrichment_Log")
    except Exception:
        return None


@st.cache_data
def load_vocabulary(version):
    """
    Build dropdown options from the Tag_Vocabulary sheet.

    The previous build hardcoded these lists in the selectbox calls, which
    silently drifted from the data: 85 of 137 tagged appraisals carried a
    value no user could select. Because the similarity scorer treats a
    value mismatch as a MISS (not an exclusion), those rows were being
    actively penalised — e.g. every HER2-positive breast appraisal scored
    zero on biomarker because only 'HER2 mutation-positive' was offered.
    Driving the options off the vocabulary sheet makes that drift
    impossible by construction.
    """
    try:
        vocab = _read_sheet(version, "Tag_Vocabulary")
    except Exception:
        return {}

    options = {}
    for field, group in vocab.groupby("Field"):
        values = [
            str(v).strip()
            for v in group["Allowed value"].dropna().tolist()
            if str(v).strip() and str(v).strip().lower() != "nan"
        ]
        values = [v for v in dict.fromkeys(values) if v.lower() != "not specified"]
        options[field] = ["Not specified"] + sorted(values, key=str.lower)
    return options


def vocab_options(vocab, field, fallback):
    return vocab.get(field, fallback)


df = load_data(data_version())
VOCAB = load_vocabulary(data_version())
TOTAL_ROWS = len(df)


def _guidance_status_category(text):
    """
    Bucket the free-text current_guidance_status into something filterable.

    The field is unstructured (~90 distinct one-off strings like 'Replaced
    by TA375' or 'Guidance withdrawn because...'), so this only classifies
    on keyword presence — it doesn't attempt to extract the replacement ID.
    Rows never checked for this stay 'Not checked' rather than being
    assumed current, since that's a materially different claim.
    """
    if pd.isna(text):
        return "Not checked"
    t = str(text).lower()
    if "no separate supersession" in t:
        return "Current"
    if any(k in t for k in ("replaced", "superseded", "withdrawn", "incorporated", "moved to static")):
        return "Replaced / withdrawn"
    return "Current"


df["guidance_status"] = df["current_guidance_status"].apply(_guidance_status_category)


# Similarity scoring

SIMILARITY_WEIGHTS = {
    "therapeutic_area": ("Disease area", 30),
    "mechanism": ("Mechanism of action", 20),
    "line_of_therapy": ("Line of therapy", 15),
    "comparator_type": ("Comparator type", 15),
    "biomarker": ("Biomarker", 10),
    "appraisal_type": ("Appraisal type", 5),
    "orphan_status": ("Orphan status", 5),
    "patient_population_size": ("Population size", 5)
}
RECENCY_MAX_POINTS = 5

FIELD_MAP = {
    "therapeutic_area": "therapeutic_area",
    "mechanism": "mechanism_of_action",
    "line_of_therapy": "line_of_therapy",
    "comparator_type": "comparator_type",
    "biomarker": "biomarker",
    "appraisal_type": "appraisal_type",
    "orphan_status": "orphan_status",
    "patient_population_size": "patient_population_size"
}


def calculate_similarity_score(query, candidate_row, dataset_max_year=None):
    """
    Weighted similarity (0-100) between a hypothetical drug profile and a
    historical appraisal, using structured tags where both sides have them.

    Factors the query didn't specify are skipped entirely rather than
    counted as misses. A small recency component breaks ties inside a
    coarse categorical cluster without inventing precision.

    Returns (score, breakdown, max_possible, same_drug).
    """

    def _val(x):
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return None
        s = str(x).strip()
        return s if s and s.lower() != "not specified" else None

    score = 0
    max_possible = 0
    breakdown = []

    for key, (label, weight) in SIMILARITY_WEIGHTS.items():
        field = FIELD_MAP[key]
        q_val = _val(query.get(field))
        c_val = _val(candidate_row.get(field))

        if q_val is None:
            continue

        if c_val is None:
            breakdown.append(
                {"label": label, "weight": weight, "status": "not_available", "points": 0}
            )
            continue

        max_possible += weight
        if q_val.lower() == c_val.lower():
            score += weight
            breakdown.append(
                {"label": label, "weight": weight, "status": "match", "points": weight}
            )
        else:
            breakdown.append(
                {"label": label, "weight": weight, "status": "no_match", "points": 0}
            )

    if max_possible == 0:
        return None, breakdown, 0, False

    candidate_year = candidate_row.get("year_start")
    if dataset_max_year and pd.notna(candidate_year):
        year_gap = max(dataset_max_year - int(candidate_year), 0)
        recency_points = round(max(RECENCY_MAX_POINTS - (year_gap * 0.3), 0), 1)
        max_possible += RECENCY_MAX_POINTS
        score += recency_points
        breakdown.append(
            {
                "label": "Recency",
                "weight": RECENCY_MAX_POINTS,
                "status": "match" if recency_points >= RECENCY_MAX_POINTS * 0.7 else "no_match",
                "points": recency_points,
            }
        )

    same_drug = False
    query_drug = query.get("_drug_name")
    if query_drug:
        cand = str(candidate_row.get("drug_name", "")).strip().lower()
        q = query_drug.strip().lower()
        same_drug = bool(q) and (q in cand or cand in q)

    pct = round(score / max_possible * 100, 1) if max_possible > 0 else None
    return pct, breakdown, max_possible, same_drug


def has_tag_coverage(similar_df):
    if "mechanism_of_action" not in similar_df.columns:
        return False
    return similar_df["mechanism_of_action"].notna().sum() > 0


# Keyword resolution

SEARCH_SYNONYMS = {
    "non-small cell lung cancer": "lung",
    "non small cell lung cancer": "lung",
    "small cell lung cancer": "lung",
    "multiple myeloma": "myeloma",
    "colorectal cancer": "colorectal",
    "bowel cancer": "colorectal",
    "breast cancer": "breast",
    "ulcerative colitis": "colitis",
    "rheumatoid arthritis": "rheumatoid",
    "crohn's": "crohn",
    "crohns": "crohn",
    "nsclc": "lung",
    "sclc": "lung",
    "crc": "colorectal",
    "tnbc": "breast",
    "mbc": "breast",
    "ibd": "colitis",
    "uc": "colitis",
    "cll": "lymphocytic",
    "cml": "leukaemia",
    "ra": "rheumatoid",
}


def resolve_keyword(keyword):
    """
    Map an indication keyword onto a term that actually appears in the
    data. The previous build used a plain dict lookup, so anything but a
    bare token fell straight through: 'NSCLC' resolved to 'lung' and
    returned 96 rows, but 'Advanced NSCLC' — the app's own placeholder
    text — resolved to nothing and returned zero.

    Longest phrases are tested first so 'small cell lung cancer' isn't
    shadowed by a shorter key.
    """
    if not keyword:
        return "", None
    cleaned = re.sub(r"\s+", " ", keyword.strip().lower())
    if not cleaned:
        return "", None

    if cleaned in SEARCH_SYNONYMS:
        return SEARCH_SYNONYMS[cleaned], cleaned

    for term in sorted(SEARCH_SYNONYMS, key=len, reverse=True):
        if re.search(rf"(?<!\w){re.escape(term)}(?!\w)", cleaned):
            return SEARCH_SYNONYMS[term], term

    return cleaned, None


# Grounded chat over the retrieved comparator set

CHAT_MODEL = "claude-sonnet-5"
CHAT_MAX_TURNS = 8
CHAT_MAX_CONTEXT_ROWS = 60   # safety ceiling only — normally the full retrieved set
CHAT_FIELD_CHARS_FULL = 400  # per-field budget for small sets
CHAT_FIELD_CHARS_TIGHT = 250 # per-field budget once the set gets large

CHAT_SYSTEM_PROMPT = """You are answering questions about a specific, already-retrieved set \
of NICE Technology Appraisal precedents for a market access consultant. You are NOT a general \
NICE or market-access assistant.

Rules, no exceptions:
1. Answer ONLY using the appraisal data provided below. Do not use outside knowledge of NICE, \
drugs, or clinical practice beyond what's in this context, even if you know it.
2. If the retrieved appraisals don't contain enough information to answer, say so plainly — \
never fill the gap with a plausible-sounding guess.
3. Never state a numeric ICER, cost, or percentage that isn't explicitly present in the data below.
4. When you draw a conclusion, name which appraisal(s) it comes from (by TA ID) so it can be \
checked against the source.
5. Keep answers short and direct — a consultant is scanning this, not reading an essay.
6. This is a descriptive summary of retrieved precedent, not a prediction of a NICE decision. \
If asked to predict an outcome, say that's outside what this data can support."""

def classify_historical_icer(row):
    lower = row.get("icer_lower")
    upper = row.get("icer_upper")

    if pd.isna(lower):
        return "not_reported"

    note = str(row.get("icer_evidence_note") or "").lower()
    same_value = pd.isna(upper) or float(upper) == float(lower)

    # Strong evidence that the values are scenarios/sensitivity analyses
    if any(term in note for term in (
        "scenario",
        "sensitivity",
        "alternative survival",
        "exploration",
    )):
        return "scenario_point" if same_value else "scenario_range"

    if same_value:
        return "point_estimate"

    # Multiple published estimates, but not necessarily a scenario range
    if any(term in note for term in (
        "ranged from",
        "probabilistic",
        "deterministic",
        "most plausible",
        " versus ",
    )):
        return "reported_range"

    # We have lower + upper, but cannot safely claim what they represent
    return "range_basis_unconfirmed"


ICER_TYPE_LABELS = {
    "point_estimate": "Point estimate",
    "scenario_point": "Scenario / indicative estimate",
    "scenario_range": "Scenario / sensitivity range",
    "reported_range": "Reported range / multiple estimates",
    "range_basis_unconfirmed": "Reported range — basis not confirmed",
    "not_reported": "Not reported",
}


def format_historical_icer(row, include_type=True):
    kind = classify_historical_icer(row)

    if kind == "not_reported":
        return "Not reported"

    lower = float(row.get("icer_lower"))
    upper = row.get("icer_upper")

    if pd.isna(upper) or float(upper) == lower:
        value = f"£{lower:,.0f}/QALY"
    else:
        value = f"£{lower:,.0f}–£{float(upper):,.0f}/QALY"

    if include_type:
        return f"{ICER_TYPE_LABELS[kind]}: {value}"

    return value

def build_chat_context(similar_df, drug_name, indication, icer_provided, cost_display,
                       threshold, comparator):
    """
    Structured context for the chat model, covering the WHOLE retrieved set
    rather than a fixed top-10 slice.

    The previous 10-row cap was invisible from the UI: a query showing 28
    similar appraisals was answered from the first 10, so any question about
    the wider pattern ("what drives rejections here?") was silently answered
    from a third of the evidence. CHAT_MAX_CONTEXT_ROWS is now a cost ceiling,
    not an editorial choice, and per-field text is tightened as the set grows
    so a large retrieval doesn't blow up the prompt.
    """
    rows = similar_df.head(CHAT_MAX_CONTEXT_ROWS)
    field_chars = CHAT_FIELD_CHARS_FULL if len(rows) <= 20 else CHAT_FIELD_CHARS_TIGHT

    query_label = drug_name if drug_name else "an unnamed hypothetical product"
    icer_line = (f"Submitted ICER {cost_display}/QALY vs a £{threshold:,}/QALY reference "
                 f"threshold." if icer_provided else
                 "No ICER submitted — this is a precedent/analogue search only.")

    # Decision mix up front so pattern questions can be answered from the
    # whole set without the model having to tally rows itself.
    counts = similar_df["decision_simple"].value_counts()
    mix = ", ".join(f"{k}: {v}" for k, v in counts.items())

    truncated = len(similar_df) - len(rows)
    coverage = (f"Retrieved precedent — all {len(rows)} appraisals in this set:"
                if truncated <= 0 else
                f"Retrieved precedent ({len(rows)} of {len(similar_df)}; "
                f"{truncated} omitted for length — say so if a question needs them):")

    blocks = [
        f"Hypothetical query: {query_label} for {indication}. {icer_line} "
        f"Comparator: {comparator or 'not specified'}.",
        f"Decision mix across the full retrieved set ({len(similar_df)}): {mix}.",
        "",
        coverage,
    ]
    for _, row in rows.iterrows():
        parts = [f"[{row['appraisal_id']}] {row['drug_name']} — {row['indication']} "
                 f"— Decision: {row['decision_simple']} ({row.get('year_label', '')})"]
        for label, col in [
            ("Rejection reasoning", "rejection_reasoning"),
            ("Detail", "detailed_reasoning"),
            ("Restriction", "restriction_note"),
            ("ICER position", "icer_evidence_note"),
            ("Committee comment", "original_nice_comment"),
        ]:
            val = row.get(col)
            if pd.notna(val) and str(val).strip().lower() not in ("", "not specified", "not applicable"):
                parts.append(f"  {label}: {str(val)[:field_chars]}")
        if pd.notna(row.get("icer_lower")):
            parts.append(f"  Reported ICER: {format_historical_icer(row)}")
        blocks.append("\n".join(parts))
    return "\n\n".join(blocks)


def ask_chat(api_key, context, history, question, system_prompt=None):
    """
    One grounded turn. Returns (answer_text, error_message) — exactly one
    is None. Uses the raw HTTP API rather than a client library, since it's
    the only Anthropic call this app makes and keeps the dependency list
    unchanged. system_prompt defaults to the NICE prompt; the G-BA and HAS
    views pass their own.
    """
    messages = list(history) + [{"role": "user", "content": question}]
    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": CHAT_MODEL,
                "max_tokens": 1000,
                "system": ((system_prompt or CHAT_SYSTEM_PROMPT)
                           + "\n\n--- Retrieved precedent data ---\n" + context),
                "messages": messages,
            },
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        text = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")
        return text.strip() or "No response text returned.", None
    except requests.exceptions.HTTPError as e:
        if resp.status_code == 401:
            return None, "API key was rejected — check it's correctly set in Streamlit secrets."
        if resp.status_code == 429:
            return None, "Rate limited — wait a moment and try again."
        return None, f"API error ({resp.status_code})."
    except requests.exceptions.RequestException:
        return None, "Couldn't reach the API — check your connection and try again."


# Reasoning synthesis

RESEARCH_QUESTIONS_BY_THEME = {
    "Comparator did not reflect NHS practice": [
        "Does the comparator reflect current NHS practice?",
        "Would the cost-effectiveness result change against the most relevant NHS comparator?",
    ],

    "Treatment effect waning / durability uncertainty": [
        "How durable is the treatment effect over longer follow-up?",
        "How sensitive are cost-effectiveness results to assumptions about treatment-effect waning?",
    ],

    "Population mismatch with retrieved precedent": [
        "Is the available evidence representative of the target patient population?",
        "Do outcomes differ materially across population size or disease-severity groups?",
    ],

    "Immature survival / follow-up evidence": [
        "How do clinical outcomes change with longer follow-up?",
        "Does additional follow-up reduce uncertainty in survival estimates?",
    ],

    "Insufficient clinical-effectiveness evidence": [
        "What additional clinical evidence is needed to establish comparative effectiveness?",
    ],

    "Indirect treatment comparison uncertainty": [
        "How robust are the results to alternative indirect-comparison methods or assumptions?",
    ],

    "Evidential or modelling uncertainty": [
        "Which model assumptions contribute most to decision uncertainty?",
    ],

    "Utility value uncertainty": [
        "How sensitive are results to alternative health-state utility estimates?",
    ],

    "Long-term extrapolation uncertainty": [
        "How sensitive are results to alternative long-term extrapolation assumptions?",
    ],

    "Survival benefit not established": [
        "Is there sufficient evidence to establish a long-term survival benefit?",
    ],

    "Treatment duration assumptions unsupported": [
        "What treatment duration is supported by the available evidence?",
    ],

    "Stopping rule assumptions unsupported": [
        "How do alternative stopping-rule assumptions affect cost effectiveness?",
    ],

    "Cost-effectiveness exceeded acceptable NHS value": [
        "Which assumptions or cost drivers contribute most to the unfavourable cost-effectiveness result?",
    ],

    "Restricted to a narrower subgroup": [
        "Which patient subgroup is most likely to benefit and be cost effective?",
    ],
}

THEME_TAGS = [
    ("cost effectiveness / value for money", "💰", "Cost-effectiveness exceeded acceptable NHS value"),
    ("appropriate use of nhs resources", "💰", "Cost-effectiveness exceeded acceptable NHS value"),
    ("immature", "⚠️", "Immature survival / follow-up evidence"),
    ("insufficient evidence", "⚠️", "Insufficient clinical-effectiveness evidence"),
    ("no direct evidence", "⚠️", "Insufficient clinical-effectiveness evidence"),
    ("indirect comparison", "⚠️", "Indirect treatment comparison uncertainty"),
    ("indirect treatment comparison", "⚠️", "Indirect treatment comparison uncertainty"),
    ("uncertain", "⚠️", "Evidential or modelling uncertainty"),
    ("utility", "⚠️", "Utility value uncertainty"),
    ("comparator", "⚠️", "Comparator did not reflect NHS practice"),
    ("subgroup", "⚠️", "Restricted to a narrower subgroup"),
    ("extrapolation", "⚠️", "Long-term extrapolation uncertainty"),
    ("survival", "⚠️", "Survival benefit not established"),
    ("treatment duration", "⚠️", "Treatment duration assumptions unsupported"),
    ("stopping rule", "⚠️", "Stopping rule assumptions unsupported"),
    # Waning treatment
    ("duration of treatment effect", "⚠️", "Treatment effect waning / durability uncertainty"),
    ("long-term relative treatment effect", "⚠️", "Treatment effect waning / durability uncertainty"),
    ("duration of gene-therapy benefit", "⚠️", "Treatment effect waning / durability uncertainty"),
    ("durability", "⚠️", "Treatment effect waning / durability uncertainty"),
]

FALLBACK_CONCERN = "Specific concerns not itemised in source text — see full guidance"

def research_questions_from_themes(patterns):
    questions = []

    for theme, _count in patterns:
        questions.extend(RESEARCH_QUESTIONS_BY_THEME.get(theme, []))

    return list(dict.fromkeys(questions))

WANING_KEYWORDS = [
    "duration of treatment effect",
    "long-term relative treatment effect",
    "duration of gene-therapy benefit",
    "durability",
]


def extract_severity(text):
    if not text:
        return None

    text = str(text).lower()

    for severity in [
        "moderately severe",
        "severe",
        "moderate",
        "mild",
    ]:
        if severity in text:
            return severity

    return None

def synthesise_themes(rejected_df, query_profile=None, all_comparables=None, indication=None, max_examples=8):
    """Theme -> (count, supporting appraisal_ids) across a set of appraisals."""
    has_detail = "detailed_reasoning" in rejected_df.columns
    texts_with_ids = []
    for _, row in rejected_df.head(max_examples).iterrows():
        text = (
            row.get("detailed_reasoning")
            if has_detail and pd.notna(row.get("detailed_reasoning"))
            else row.get("rejection_reasoning")
        )
        texts_with_ids.append(
            (str(text).lower() if pd.notna(text) else "", row.get("appraisal_id", "?"))
        )

    total = len(texts_with_ids)
    if total == 0:
        return [], 0, {}

    seen_labels = {}
    theme_sources = {}
    for key, emoji, label in THEME_TAGS:
        matching = [aid for t, aid in texts_with_ids if key in t]
        if not matching:
            continue
        if label not in seen_labels or len(matching) > seen_labels[label][1]:
            seen_labels[label] = (emoji, len(matching))
            theme_sources[label] = matching
        elif label in theme_sources:
            theme_sources[label] = list(dict.fromkeys(theme_sources[label] + matching))
    
        # NEW — structured population mismatch
    if query_profile:
        user_population = query_profile.get("patient_population_size")
        if (user_population and user_population != "Not specified" and "patient_population_size" in rejected_df.columns):
            mismatched_ids = []
            for _, row in rejected_df.head(max_examples).iterrows():
                appraisal_population = row.get("patient_population_size")
                if (pd.notna(appraisal_population) 
                    and str(appraisal_population).strip()
                    and str(appraisal_population).lower() != "not specified"
                    and str(appraisal_population).strip().lower()
                    != str(user_population).strip().lower()):
                    mismatched_ids.append(row.get("appraisal_id", "?"))
            if mismatched_ids:
                label = "Population mismatch with retrieved precedent"
                seen_labels[label] = ("⚠️",len(mismatched_ids))
                theme_sources[label] = mismatched_ids
 # NEW — disease severity mismatch
    if indication and all_comparables is not None:
        user_severity = extract_severity(indication)

        if user_severity:
            severity_mismatches = []

            for _, row in all_comparables.head(max_examples).iterrows():
                appraisal_severity = extract_severity(row.get("indication"))

                if (
                    appraisal_severity
                    and appraisal_severity != user_severity
                ):
                    severity_mismatches.append(
                        row.get("appraisal_id", "?")
                    )

            if severity_mismatches:
                label = "Population mismatch with retrieved precedent"

                existing_sources = theme_sources.get(label, [])

                combined_sources = list(dict.fromkeys(
                    existing_sources + severity_mismatches
                ))

                seen_labels[label] = (
                    "⚠️",
                    len(combined_sources)
                )

                theme_sources[label] = combined_sources


    # NEW — treatment-effect waning across ALL retrieved comparables
    if all_comparables is not None:
        waning_matches = []

        for _, row in all_comparables.iterrows():
            detailed = row.get("detailed_reasoning")
            rejection = row.get("rejection_reasoning")

            text = " ".join([
                str(detailed) if pd.notna(detailed) else "",
                str(rejection) if pd.notna(rejection) else "",
            ]).lower()

            if any(keyword in text for keyword in WANING_KEYWORDS):
                waning_matches.append(
                    row.get("appraisal_id", "?")
                )

        if waning_matches:
            label = "Treatment effect waning / durability uncertainty"

            seen_labels[label] = (
                "⚠️",
                len(waning_matches)
            )

            theme_sources[label] = list(
                dict.fromkeys(waning_matches)
            )
    ranked = sorted(seen_labels.items(), key=lambda x: x[1][1], reverse=True)
    return ranked, total, theme_sources


def structure_reasoning_card(row, has_detail_col):
    """Split one rejection into conclusion / concerns / ICER line."""
    conclusion = None
    if has_detail_col and pd.notna(row.get("primary_reason_category")):
        conclusion = row["primary_reason_category"]

    raw = (
        row.get("detailed_reasoning")
        if has_detail_col and pd.notna(row.get("detailed_reasoning"))
        else row.get("rejection_reasoning")
    )
    raw = str(raw) if pd.notna(raw) else ""
    lower = raw.lower()

    if not conclusion:
        if "appropriate use of nhs resources" in lower or "value for money" in lower:
            conclusion = (
                "Not recommended because the incremental health benefit did not "
                "justify the additional NHS cost"
            )
        elif "insufficient" in lower or "no direct evidence" in lower:
            conclusion = "Not recommended due to insufficient clinical-effectiveness evidence"
        else:
            conclusion = "Not recommended (see full committee guidance for stated reason)"

    concerns = []
    if has_detail_col and pd.notna(row.get("secondary_factors")):
        concerns = [f.strip() for f in str(row["secondary_factors"]).split(";") if f.strip()]
    if not concerns:
        for token, label in [
            ("immature", "Immature clinical/survival evidence"),
            ("uncertain", "Evidential or modelling uncertainty"),
            ("utility", "Utility value uncertainty"),
            ("comparator", "Comparator concerns"),
        ]:
            if token in lower:
                concerns.append(label)
        if not concerns:
            concerns.append(FALLBACK_CONCERN)

    # Prefer the curated evidence note over regex-scraping the prose.
    icer_line = None
    note = row.get("icer_evidence_note")

    if pd.notna(row.get("icer_lower")):
        icer_line = format_historical_icer(row)

        if pd.notna(note) and str(note).strip().lower() not in ("", "not specified"):
            icer_line += f" — {str(note).strip()}"

    elif pd.notna(note) and str(note).strip().lower() not in ("", "not specified"):
        icer_line = str(note).strip()

    elif "no publishable numeric icer" in lower or "no publishable icer" in lower:
        icer_line = "No publishable ICER available"

    elif "£" in raw:
        matches = re.findall(r"£[\d,]+(?:\s*per\s*QALY|/QALY)?", raw)
        icer_line = "; ".join(dict.fromkeys(matches[:3])) if matches else None

    return {
        "conclusion": conclusion,
        "concerns": concerns,
        "icer_line": icer_line or "Not reported in source text",
        "raw": raw,
    }


def build_concern_frequency(rejected_df, has_detail_col, max_rows=15):
    cards = [
        (row, structure_reasoning_card(row, has_detail_col))
        for _, row in rejected_df.head(max_rows).iterrows()
    ]
    counter = Counter()
    for _, card in cards:
        for c in set(card["concerns"]):
            counter[c] += 1
    return cards, counter, len(cards)


def split_shared_unique(card_concerns, counter, sample_size):
    """Split one appraisal's concerns into shared vs unique across the set."""
    shared, unique = [], []
    for c in card_concerns:
        if c == FALLBACK_CONCERN:
            continue
        if counter.get(c, 1) > 1:
            shared.append((c, counter[c]))
        else:
            unique.append(c)
    return shared, unique


# PDF report

def generate_assessment_pdf(
    drug_name, indication, estimated_cost, end_of_life, comparator, threshold,
    appraisal_type, total_similar, recommended_count, optimised_count,
    rejected_count, managed_count, terminated_count, approval_rate,
    similar, patterns, warnings_list, verdict, icer_provided, cost_display, research_questions
):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=1.5 * cm, leftMargin=1.5 * cm,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm,
    )

    heading = ParagraphStyle("heading", fontSize=12, spaceAfter=4, spaceBefore=8,
                             fontName="Helvetica-Bold", textColor=colors.HexColor("#2c3e50"))
    body = ParagraphStyle("body", fontSize=9, spaceAfter=3, fontName="Helvetica", leading=12)
    small = ParagraphStyle("small", fontSize=8, spaceAfter=3, fontName="Helvetica",
                           textColor=colors.grey)

    verdict_color = (
        colors.HexColor("#27ae60") if "Likely" in verdict
        else colors.HexColor("#e67e22") if "Borderline" in verdict
        else colors.HexColor("#7f8c8d") if "Precedent" in verdict
        else colors.HexColor("#e74c3c")
    )

    content = []
    pdf_drug_label = drug_name if drug_name else f"Analogue search — {indication}"

    header = Table(
        [[
            Paragraph(f"<b>{pdf_drug_label}</b>", ParagraphStyle(
                "h", fontSize=16, fontName="Helvetica-Bold", textColor=colors.white)),
            Paragraph(
                f"Market Access Intelligence Report<br/><font size=9>{indication}</font>",
                ParagraphStyle("hr", fontSize=11, fontName="Helvetica",
                               textColor=colors.white, alignment=TA_RIGHT)),
        ]],
        colWidths=[9 * cm, 9 * cm],
    )
    header.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#2c3e50")),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))
    content.append(header)
    content.append(Spacer(1, 0.2 * cm))

    content.append(Paragraph(
        f"Generated: {date.today().strftime('%d %B %Y')}   |   "
        f"Comparator: {comparator or 'Not specified'}   |   "
        f"End of Life: {end_of_life}   |   CONFIDENTIAL", small))
    content.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#bdc3c7")))
    content.append(Spacer(1, 0.2 * cm))

    verdict_text = (
        "RISK SIGNAL: LOW - submitted ICER within reference threshold" if "Likely" in verdict
        else "RISK SIGNAL: MODERATE - submitted ICER exceeds threshold" if "Borderline" in verdict
        else "RISK SIGNAL: HIGH COMMERCIAL/PRICING RISK PATTERN" if "Commercial" in verdict
        else "PRECEDENT REFERENCE ONLY - no ICER submitted" if "Precedent" in verdict
        else "RISK SIGNAL: HIGH - submitted ICER substantially exceeds threshold"
    )
    vt = Table([[Paragraph(verdict_text, ParagraphStyle(
        "vb", fontSize=10, fontName="Helvetica-Bold", textColor=colors.white))]],
        colWidths=[18 * cm])
    vt.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), verdict_color),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
    ]))
    content.append(vt)
    content.append(Spacer(1, 0.3 * cm))

    verdict_plain = {
        "High Commercial Risk": "high commercial/pricing risk",
        "Precedent Reference Only": "no ICER submitted — precedent reference only",
        "Likely Recommended": "low risk",
        "Borderline": "moderate risk",
        "Unlikely to be Recommended": "high risk",
    }.get(verdict, verdict.lower())

    content.append(Paragraph("Executive Summary", heading))
    if icer_provided:
        icer_summary = (
            f"The submitted ICER of {cost_display}/QALY sits "
            f"{((estimated_cost / threshold) - 1) * 100:+.0f}% relative to the "
            f"£{threshold:,}/QALY reference threshold, producing an initial risk "
            f"signal of <b>{verdict_plain}</b>."
        )
    else:
        icer_summary = (
            f"No ICER was submitted for this query, so this is a precedent reference "
            f"only (<b>{verdict_plain}</b>) rather than a threshold-based risk signal."
        )
    content.append(Paragraph(
        f"{pdf_drug_label} for {indication} has been benchmarked against {total_similar} NICE "
        f"technology appraisals retrieved by indication keyword match. {icer_summary} "
        f"Within the retrieved set, {approval_rate:.0f}% of appraisals were recommended or "
        f"optimised ({recommended_count} recommended, {optimised_count} optimised, "
        f"{rejected_count} not recommended) — this is descriptive of the retrieved precedent "
        f"only and is not a predicted probability of a NICE decision for this submission.",
        body))
    content.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#bdc3c7")))
    content.append(Spacer(1, 0.2 * cm))

    content.append(Paragraph("Economic Profile & Appraisal Landscape", heading))

    econ_data = [
        ["Parameter", "Value"],
        ["Estimated ICER", f"{cost_display}/QALY" if icer_provided else "Not provided"],
        ["WTP Threshold", f"£{threshold:,}/QALY"],
        ["Position vs Threshold",
         f"{((estimated_cost / threshold) - 1) * 100:+.0f}%" if icer_provided else "N/A"],
        ["Comparator", comparator or "Not specified"],
        ["End of Life", end_of_life],
        # Previously hardcoded to "STA" regardless of the user's selection.
        ["Appraisal Type", appraisal_type],
    ]

    def _pct(n):
        return f"{n / total_similar * 100:.0f}%" if total_similar > 0 else "N/A"

    # Terminated is included so the rows reconcile against Total Similar.
    landscape_data = [
        ["Decision", "Count", "%"],
        ["Recommended", str(recommended_count), _pct(recommended_count)],
        ["Optimised", str(optimised_count), _pct(optimised_count)],
        ["Not Recommended", str(rejected_count), _pct(rejected_count)],
        ["Managed Access", str(managed_count), _pct(managed_count)],
        ["Terminated", str(terminated_count), _pct(terminated_count)],
        ["Total Similar", str(total_similar), "100%"],
        ["Recommendation proportion", f"{approval_rate:.0f}%", ""],
    ]

    ts = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#34495e")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f2f2")]),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
    ])

    econ_table = Table(econ_data, colWidths=[4.5 * cm, 4 * cm])
    econ_table.setStyle(ts)
    land_table = Table(landscape_data, colWidths=[4 * cm, 2 * cm, 2 * cm])
    land_table.setStyle(ts)

    two_col = Table([[econ_table, land_table]], colWidths=[9 * cm, 9 * cm])
    two_col.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]))
    content.append(two_col)
    content.append(Spacer(1, 0.3 * cm))
    content.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#bdc3c7")))

    content.append(Paragraph("Similar NICE Appraisals (Top 10)", heading))
    sim_data = [["Drug", "Decision", "Indication", "Year", "TA ID", "Historical ICER"]]
    for _, row in similar.head(10).iterrows():
        sim_data.append([
            str(row["drug_name"])[:16],
            str(row["decision_simple"]),
            str(row["indication"])[:30],
            str(row.get("year_label", row.get("year", ""))),
            str(row["appraisal_id"]),
            format_historical_icer(row)
            if pd.notna(row.get("icer_lower"))
            else "Not reported",
        ])
    sim_table = Table(
    sim_data,
    colWidths=[
        2.8 * cm,
        2.6 * cm,
        5.2 * cm,
        1.5 * cm,
        1.5 * cm,
        4.4 * cm,
    ])
    sim_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#34495e")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f2f2")]),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
    ]))
    content.append(sim_table)
    content.append(Spacer(1, 0.3 * cm))
    content.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#bdc3c7")))

    content.append(Paragraph("Rejection Risk Analysis & Contextual Considerations", heading))

    if patterns:
        pat_data = [["Rejection Theme", "Freq"]] + [[p, str(c)] for p, c in patterns]
        pat_table = Table(pat_data, colWidths=[6.5 * cm, 1.5 * cm])
        pat_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#c0392b")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.HexColor("#fdf2f2"), colors.white]),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ]))
    else:
        pat_table = Paragraph("No common rejection patterns identified.", body)

    warn_rows = [[Paragraph(f"- {w}", small)] for w in warnings_list] or [
        [Paragraph("No major contextual concerns identified.", small)]
    ]
    warn_table = Table(warn_rows, colWidths=[9 * cm])
    warn_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#fef9e7")),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#f39c12")),
    ]))

    two_col2 = Table([[pat_table, warn_table]], colWidths=[9 * cm, 9 * cm])
    two_col2.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]))
    content.append(two_col2)
    content.append(Spacer(1, 0.3 * cm))
    content.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#bdc3c7")))
    content.append(Paragraph("Research Questions", heading))
    if research_questions:
        question_rows = [
            [Paragraph(f"{i + 1}. {question}", body)]
            for i, question in enumerate(research_questions)
        ]
    else:
        question_rows = [[
            Paragraph(
                "No specific research questions were generated because no matching "
                "risk themes were identified in the retrieved appraisal set.",
                body
            )
        ]]
    question_table = Table(
        question_rows,
        colWidths=[18 * cm]
    )
    question_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#eaf4fb")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#2980b9")),
        ("LINEBELOW", (0, 0), (-1, -2), 0.3, colors.HexColor("#d6eaf8")),
    ]))
    content.append(question_table)
    content.append(Spacer(1, 0.3 * cm))
    content.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#bdc3c7")))
    content.append(Spacer(1, 0.1 * cm))
    content.append(Paragraph(
        "Disclaimer: This report was generated automatically based on historical NICE appraisal "
        "data. It is intended as a preliminary intelligence tool only. Further economic "
        "modelling and expert review is strongly recommended before drawing conclusions or "
        "making submission decisions. Generated by NICE Technology Appraisal Intelligence Tool "
        f"| {date.today().strftime('%d %B %Y')}", small))

    doc.build(content)
    buffer.seek(0)
    return buffer


# ═════════════════════════════════════════════════════════════════════════════
# Multi-agency support — G-BA (Germany) and HAS (France)
#
# G-BA and HAS don't answer NICE's question. The G-BA rates added clinical
# benefit against a comparator it sets itself (AMNOG); HAS rates clinical
# benefit (SMR/ASMR) and decides early access. Neither produces an ICER in
# these records, so their views are built around their own outcome scales
# rather than squeezed into the NICE ICER-vs-threshold frame. Everything below
# is driven by AGENCY_CONFIG, so loading more rows into the G-BA_Recent /
# HAS_Recent sheets needs no code changes.
# ═════════════════════════════════════════════════════════════════════════════

AGENCY_LABELS = {
    "NICE": "🇬🇧 NICE — England",
    "G-BA": "🇩🇪 G-BA — Germany",
    "HAS": "🇫🇷 HAS — France",
}
AGENCY_FLAGS = {"NICE": "🇬🇧", "G-BA": "🇩🇪", "HAS": "🇫🇷"}
AGENCY_SORT = {"NICE": 0, "G-BA": 1, "HAS": 2}
SMALL_AGENCY_SET = 20       # below this, a view carries an "illustrative only" note
SYNTHESIS_MAX_ITEMS = 15    # per evidence-synthesis block before "narrow the selection"
DOC_INDEX_COLUMNS = ["Institution", "Appraisal_ID", "Document_Type",
                     "Document_Title", "Date", "Source_URL", "Notes"]


@st.cache_data
def load_agency_sheet(sheet_name, version):
    """Raw rows for a non-NICE agency. A missing sheet yields an empty frame, so
    an older workbook still opens — the NICE view never depends on these."""
    try:
        return _read_sheet(version, sheet_name)
    except Exception:
        return pd.DataFrame()


@st.cache_data
def load_document_index(version):
    """Appraisal_Documents: one row per official source document, all agencies."""
    try:
        docs = _read_sheet(version, "Appraisal_Documents")
    except Exception:
        return pd.DataFrame(columns=DOC_INDEX_COLUMNS)
    for col in ("Institution", "Appraisal_ID"):
        docs[col] = docs[col].astype(str).str.strip()
    return docs


# ── Text helpers ─────────────────────────────────────────────────────────────

_PLACEHOLDER_VALUES = {"", "nan", "none", "n/a", "na", "-", "—", "not specified", "not reported",
                       "sans objet", "non applicable", "non évalué", "non evalue"}
# Notes that are useful on one record but say nothing across a set — English and
# French forms, since HAS rows may be entered either way.
_NOTE_PREFIXES = ("not applicable", "not assessed", "not reported", "not recorded",
                  "non applicable", "non évalué", "non evalue", "non renseigné",
                  "non renseigne", "non disponible", "sans objet", "n/a")


def _text(value):
    """Stripped string, or None for blank / NaN / bare placeholder values."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    s = str(value).strip()
    return None if s.lower() in _PLACEHOLDER_VALUES else s


def _substantive(value):
    """Like _text, but also drops 'Not applicable — …' / 'Not assessed …' notes.
    Those are worth showing on a single record but are noise in a cross-set
    synthesis ('Comparator: not applicable' three times says nothing)."""
    s = _text(value)
    if s and s.lower().startswith(_NOTE_PREFIXES):
        return None
    return s


def _shorten(text, limit):
    s = str(text)
    return s if len(s) <= limit else s[:limit].rsplit(" ", 1)[0] + "…"


def _md_safe(text):
    """Stop '$' in source text being rendered as LaTeX by st.markdown."""
    return str(text).replace("$", "\\$")


def _md_link(label, url):
    """Markdown link that survives brackets in the label and parentheses in the URL."""
    label = str(label).replace("[", "(").replace("]", ")")
    url = str(url).strip().replace(" ", "%20").replace("(", "%28").replace(")", "%29")
    return f"[{label}]({url})"


def _parse_date(value):
    """
    One workbook date → Timestamp (NaT if unreadable).

    ISO text ('2026-09-03') is read as ISO; anything else day-first, because
    G-BA and HAS dates are European — '03.09.2026' and '03/09/2026' are 3
    September, not 9 March. Real Excel dates pass straight through, and bare
    Excel serial numbers (46268) are converted rather than read as 1970.

    Some cells hold a short history of dates rather than one (e.g. a G-BA
    procedure's original resolution and a later amendment, joined by blank
    lines: '2017-03-16\\n\\n2016-09-15'). Handing that whole blob to the
    parser lets it misread the leftover text as a time/timezone fragment
    (observed: '2017-03-16 20:16:00-15:00', a tz-aware artifact that then
    crashes the column cast against its naive neighbours) — so only the
    first date in the cell is parsed, on the assumption it's the current
    one. The tz strip below is a second safety net for any other cell that
    parses to something tz-aware despite that.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime)):
        result = pd.Timestamp(value)
        return result.tz_localize(None) if result.tzinfo is not None else result
    if isinstance(value, (int, float)):
        if 20000 <= value <= 80000:
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(value))
        return pd.NaT
    s = str(value).strip()
    if not s or s.lower() in _PLACEHOLDER_VALUES:
        return pd.NaT
    first = re.split(r"[\n;]+", s, maxsplit=1)[0].strip()
    if first:
        s = first
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        if re.match(r"^\d{4}-\d{1,2}-\d{1,2}(?:[ T]|$)", s):
            result = pd.to_datetime(s[:10], format="%Y-%m-%d", errors="coerce")
        else:
            result = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if isinstance(result, pd.Timestamp) and result.tzinfo is not None:
        result = result.tz_localize(None)
    return result


def _to_datetime(series):
    return pd.Series([_parse_date(v) for v in series], index=series.index,
                     dtype="datetime64[ns]")


def _date_span(dates):
    d = dates.dropna()
    if d.empty:
        return "decision dates not recorded"
    lo, hi = d.min(), d.max()
    if (lo.year, lo.month) == (hi.year, hi.month):
        return f"{lo:%B %Y}"
    if lo.year == hi.year:
        return f"{lo:%b}–{hi:%b %Y}"
    return f"{lo:%b %Y}–{hi:%b %Y}"


def _decision_text(row):
    return " ".join(
        _text(row.get(c)) or "" for c in ("decision_simple", "decision_raw")
    ).lower()


# ── G-BA outcome scale ───────────────────────────────────────────────────────

GBA_OUTCOME_ORDER = [
    "Major added benefit",
    "Considerable added benefit",
    "Minor added benefit",
    "Non-quantifiable added benefit",
    "Mixed by subgroup",
    "No added benefit proven",
    "Lesser benefit",
    "Unclassified",
]
GBA_POSITIVE = {
    "Major added benefit", "Considerable added benefit",
    "Minor added benefit", "Non-quantifiable added benefit",
}
# Positive extents are an ordered scale, so they take one green ramp (dark =
# major); the rest are distinct hues. Validated for colour-vision separation.
GBA_COLORS = {
    "Major added benefit": "#145a32",
    "Considerable added benefit": "#1e8449",
    "Minor added benefit": "#52be80",
    "Non-quantifiable added benefit": "#2e86c1",
    "Mixed by subgroup": "#f39c12",
    "No added benefit proven": "#e74c3c",
    "Lesser benefit": "#922b21",
    "Unclassified": "#95a5a6",
}
GBA_MD_COLORS = {
    "Major added benefit": "green", "Considerable added benefit": "green",
    "Minor added benefit": "green", "Non-quantifiable added benefit": "blue",
    "Mixed by subgroup": "orange", "No added benefit proven": "red",
    "Lesser benefit": "red",
}


_GBA_EXTENT_PATTERNS = [
    # Whole-word patterns, so 'majority' never reads as 'major'. English as the
    # workbook records it, plus the German forms used in the resolutions.
    ("Lesser benefit", r"\bless(?:er)? benefit\b|\bgeringere[nr]? nutzen\b"),
    ("No added benefit proven",
     r"\bno (?:added|additional) benefit\b|\bnot proven\b|\bnicht belegt\b"
     r"|\bkein zusatznutzen\b"),
    ("Major added benefit", r"\bmajor\b|\berheblich(?:e[nrs]?)?\b"),
    ("Considerable added benefit", r"\bconsiderable\b|\bbeträchtlich(?:e[nrs]?)?\b"),
    ("Non-quantifiable added benefit",
     r"\bnon-quantifiable\b|\bnot quantifiable\b|\bnicht quantifizierbar(?:e[nrs]?)?\b"),
    ("Minor added benefit", r"\bminor\b|\bgering(?:e[nrs]?)?\b"),
]
_GBA_CERTAINTY_PATTERNS = [
    ("Proof", r"\bproof\b|\bbeleg für\b"),
    ("Indication", r"\bindication (?:of|for)\b|\bindication\s*[,)]|:\s*indication\b|\bhinweis\b"),
    ("Hint", r"\bhint\b|\banhaltspunkt\b"),
]


def _clauses(value):
    """Split free text into clauses — one per subgroup finding, usually."""
    s = _text(value)
    if not s:
        return []
    return [c.strip().lower() for c in re.split(r"[;\n|]|(?<=[A-Za-z0-9)])\.\s+", s)
            if c.strip()]


def _extents_in(value):
    found = []
    for clause in _clauses(value):
        for extent, pattern in _GBA_EXTENT_PATTERNS:
            if re.search(pattern, clause) and extent not in found:
                found.append(extent)
    return found


def _gba_extents(row):
    """(extents named in the curated decision fields, extents in the added-benefit text)."""
    decision = []
    for col in ("decision_simple", "decision_raw"):
        for extent in _extents_in(row.get(col)):
            if extent not in decision:
                decision.append(extent)
    return decision, _extents_in(row.get("added_benefit_assessment"))


def gba_outcome(row):
    """
    Bucket a G-BA resolution onto the AMNOG extent-of-added-benefit scale.

    decision_simple is curated free text ('Added benefit: hint,
    non-quantifiable', 'Mixed: no added benefit for most groups; one subgroup
    has a hint'). That reads well but would give every new phrasing its own
    metric card, so each clause is matched against the six statutory extents.
    A resolution naming different extents for different subgroups — or saying
    'mixed' outright — is 'Mixed by subgroup', whether or not the word appears.
    The curated decision fields are preferred; the longer added-benefit text is
    only used when they name no extent. The recorded wording is always shown
    alongside, so the bucketing never hides anything.
    """
    decision, detail = _gba_extents(row)
    extents = decision or detail
    mixed = re.search(r"\bmixed\b|\bby subgroup\b|\bdiffers? by\b", _decision_text(row))
    if mixed or len(extents) > 1:
        return "Mixed by subgroup"
    return extents[0] if extents else "Unclassified"


def gba_certainty(row):
    """
    Certainty of evidence behind a rating: proof / indication / hint.

    Read from the clauses that actually carry a rating (decision fields first),
    so 'indication' in its therapeutic sense elsewhere in the text is never
    taken for the AMNOG certainty level. None where no benefit was found.
    """
    if row.get("outcome") in ("No added benefit proven", "Unclassified"):
        return None
    rated = GBA_POSITIVE | {"Lesser benefit"}
    decision_clauses = (_clauses(row.get("decision_simple"))
                        + _clauses(row.get("decision_raw")))
    detail_clauses = _clauses(row.get("added_benefit_assessment"))

    def first_certainty(clauses, rated_only):
        for clause in clauses:
            if rated_only and not any(
                    re.search(p, clause) for e, p in _GBA_EXTENT_PATTERNS if e in rated):
                continue
            for level, pattern in _GBA_CERTAINTY_PATTERNS:
                if re.search(pattern, clause):
                    return level
        return None

    return (first_certainty(decision_clauses, True)
            or first_certainty(detail_clauses, True)
            or first_certainty(decision_clauses, False))


def gba_favourable(row):
    """
    True where at least one patient subgroup received an added benefit; False
    where every named rating is 'no added benefit' or 'lesser benefit'; None
    where nothing is named.

    For single-outcome rows only the curated decision fields count, so a
    manufacturer's claim quoted in the longer text can't flip the result. Mixed
    rows also read the added-benefit text, which is where the subgroup ratings
    are usually spelled out.
    """
    decision, detail = _gba_extents(row)
    mixed = row.get("outcome") == "Mixed by subgroup"
    named = set(decision)
    if mixed or not named:
        named |= set(detail)
    if named & GBA_POSITIVE:
        return True
    # 'One subgroup has a hint' names a certainty level but not its extent — it
    # could be a hint of benefit or of lesser benefit, so leave it undetermined.
    if mixed and "Lesser benefit" not in named and any(
            re.search(pattern, clause)
            for clause in _clauses(row.get("decision_simple")) + _clauses(row.get("decision_raw"))
            for _, pattern in _GBA_CERTAINTY_PATTERNS):
        return None
    return False if named else None


def gba_subtitle(frame):
    return f"{len(frame):,} AMNOG early benefit assessments (§35a SGB V)"


def gba_headlines(frame):
    n = len(frame)
    positive = int(frame["favourable"].eq(True).sum())
    return [
        ("Added benefit in ≥1 subgroup",
         f"{positive}/{n}" + (f" ({positive / n * 100:.0f}%)" if n else ""),
         "Major, considerable, minor or non-quantifiable added benefit, plus mixed "
         "resolutions where at least one subgroup received a positive rating. Descriptive "
         "of this selection only. Not equivalent to NICE's recommendation proportion: a "
         "medicine without added benefit is still reimbursed in Germany; the rating shapes "
         "its negotiated price."),
    ]


# ── HAS outcome scale ────────────────────────────────────────────────────────

HAS_OUTCOME_ORDER = [
    "Early access granted",
    "Early access renewed",
    "Early access refused",
    "Early access — other",
    "ASMR I", "ASMR II", "ASMR III", "ASMR IV", "ASMR V",
    "SMR sufficient, no ASMR recorded",
    "SMR mixed by population",
    "SMR insufficient",
    "Unclassified",
]
# ASMR I–IV is an ordered scale → one green ramp (dark = major improvement).
HAS_COLORS = {
    "Early access granted": "#17a589",
    "Early access renewed": "#2a78d6",
    "Early access refused": "#e74c3c",
    "Early access — other": "#95a5a6",
    "ASMR I": "#0f3b22",
    "ASMR II": "#176b3a",
    "ASMR III": "#239b56",
    "ASMR IV": "#45c07a",
    "ASMR V": "#7f8c8d",
    "SMR sufficient, no ASMR recorded": "#8e44ad",
    "SMR mixed by population": "#f39c12",
    "SMR insufficient": "#922b21",
    "Unclassified": "#bdc3c7",
}
HAS_MD_COLORS = {
    "Early access granted": "green", "Early access renewed": "blue",
    "Early access refused": "red", "ASMR I": "green", "ASMR II": "green",
    "ASMR III": "green", "ASMR IV": "green", "ASMR V": "gray",
    "SMR sufficient, no ASMR recorded": "violet", "SMR mixed by population": "orange",
    "SMR insufficient": "red",
}
SMR_ORDER = ["Important", "Moderate", "Low", "Mixed", "Insufficient"]
_SMR_PATTERNS = [
    ("Insufficient", r"insuffis|insuffic"),
    ("Important", r"\bimportant\b"),
    ("Moderate", r"\bmodér|\bmoder"),
    ("Low", r"\bfaible\b|\blow\b"),
]
_ASMR_ORDER = ["I", "II", "III", "IV", "V"]
_HAS_EA_NEGATIVE = (r"\brefus|\bnot renewed\b|\bnon[- ]?renouvel|\bnot granted\b"
                    r"|\bdéfavorable\b|\bdefavorable\b|\bunfavou?rable\b|\brejected\b")


def has_is_early_access(row):
    """Decided by assessment_type where it's filled in — a Transparency Committee
    opinion that merely mentions a past early-access period is still an opinion.
    Falls back to the decision wording only when the type is blank."""
    source = _text(row.get("assessment_type")) or _decision_text(row)
    t = source.lower()
    return any(k in t for k in ("early access", "early-access", "accès précoce", "acces precoce"))


def asmr_level(value):
    """'ASMR IV (mineure)' → 'IV'. Where different populations got different
    levels, the best one (the headline counts 'improvement for ≥1 population').
    None where not assessed, not applicable or unparseable."""
    s = _substantive(value)
    if not s:
        return None
    levels = re.findall(r"\b(IV|V|I{1,3})\b", s.upper())
    return min(levels, key=_ASMR_ORDER.index) if levels else None


def smr_level(value):
    """'SMR important' / 'modéré' / 'insuffisant' → a fixed English label;
    'Mixed' where populations differ; None for anything unreadable, so a note
    like 'Non évalué — accès précoce' is never counted as a rating."""
    s = _substantive(value)
    if not s:
        return None
    t = s.lower()
    levels = [label for label, pattern in _SMR_PATTERNS if re.search(pattern, t)]
    if not levels:
        return None
    return levels[0] if len(levels) == 1 else "Mixed"


def has_outcome(row):
    """
    Early-access decisions → granted / renewed / refused; reimbursement opinions
    → ASMR level, 'SMR insufficient', or an SMR-only bucket. Negative wording is
    tested first: 'not renewed' contains 'renewed' and 'défavorable' contains
    'favorable'. Anything unreadable is 'Unclassified' (the recorded wording is
    still shown), rather than every new phrasing becoming its own category.
    """
    t = _decision_text(row)
    if has_is_early_access(row):
        if re.search(_HAS_EA_NEGATIVE, t):
            return "Early access refused"
        if re.search(r"\brenew|\brenouvel", t):
            return "Early access renewed"
        if re.search(r"\bgrant|\bauthori[sz]|\boctro[iy]|\baccord[ée]e?\b|\bfavou?rable\b", t):
            return "Early access granted"
        return "Early access — other"
    smr = smr_level(row.get("SMR_rating"))
    if smr == "Insufficient":
        return "SMR insufficient"
    level = asmr_level(row.get("ASMR_rating"))
    if level:
        return f"ASMR {level}"
    if smr == "Mixed":
        return "SMR mixed by population"
    if smr:
        return "SMR sufficient, no ASMR recorded"
    return "Unclassified"


def has_favourable(row):
    """ASMR V (no improvement) is neutral, not a refusal: the medicine can still
    be reimbursed, it just gains no price premium."""
    outcome = row.get("outcome")
    if outcome in ("Early access granted", "Early access renewed",
                   "ASMR I", "ASMR II", "ASMR III", "ASMR IV"):
        return True
    if outcome in ("Early access refused", "SMR insufficient"):
        return False
    return None


def has_ratings_text(row):
    smr, asmr = _substantive(row.get("SMR_rating")), _substantive(row.get("ASMR_rating"))
    parts = ([f"SMR: {smr}"] if smr else []) + ([f"ASMR: {asmr}"] if asmr else [])
    return " · ".join(parts) or None


def has_subtitle(frame):
    early = int(frame["_early_access"].sum())
    other = len(frame) - early
    parts = ([f"{early} early access"] if early else []) + (
        [f"{other} Transparency Committee / CEESP"] if other else [])
    return f"{len(frame):,} HAS decisions ({', '.join(parts)})"


def has_headlines(frame):
    out = []
    early = frame[frame["_early_access"]]
    if len(early):
        ok = int(early["outcome"].isin(["Early access granted", "Early access renewed"]).sum())
        out.append(("Early access granted or renewed", f"{ok}/{len(early)}",
                    "Early-access decisions in this selection that granted or renewed access."))
    asmr = frame[frame["_asmr"].notna()]
    if len(asmr):
        ok = int(asmr["_asmr"].isin(["I", "II", "III", "IV"]).sum())
        out.append(("ASMR I–IV", f"{ok}/{len(asmr)}",
                    "Reimbursement opinions recognising some improvement in clinical benefit "
                    "(ASMR V = none)."))
    smr = frame[frame["_smr"].notna()]
    if len(smr):
        ok = int((smr["_smr"] != "Insufficient").sum())
        out.append(("Sufficient SMR (≥1 population)", f"{ok}/{len(smr)}",
                    "Opinions rating the actual clinical benefit sufficient for reimbursement "
                    "in at least one population."))
    return out


# ── Methodology text & chat prompts ──────────────────────────────────────────

GBA_METHODOLOGY = """
**What the G-BA decides.** Under AMNOG (§35a SGB V) the G-BA rates the *added clinical
benefit* of a new medicine against an **appropriate comparator therapy (zVT)** that the
G-BA itself specifies — normally after an IQWiG dossier assessment, written statements
and an oral hearing. The resolution feeds the price negotiation with the
GKV-Spitzenverband. It is not a cost-effectiveness judgement.

**How outcomes are shown.** The extent of added benefit is one of major, considerable,
minor, non-quantifiable, no added benefit proven, or lesser benefit. Each positive rating
carries a certainty of evidence (proof, indication or hint), and resolutions are often
split by patient subgroup. The *Outcome* column maps each resolution onto that scale,
using *Mixed by subgroup* where subgroups differ. The wording recorded in the workbook is
always shown alongside it.

**What doesn't carry over from the NICE view.** There is no ICER, willingness-to-pay
threshold or threshold-based risk signal here, because the G-BA doesn't assess
cost-effectiveness in this procedure. The closest headline to NICE's recommendation
proportion is **added benefit in at least one subgroup**, but the two aren't equivalent: a
medicine without added benefit stays reimbursable in Germany. The rating shapes its
negotiated price, which is anchored to the cost of the comparator therapy.
"""

HAS_METHODOLOGY = """
**What HAS decides.** The Transparency Committee (CT) rates **SMR**, the actual clinical
benefit, which informs whether a medicine is listed for reimbursement and at what rate. It
also rates **ASMR**, the improvement in clinical benefit from I (major) to V (none), which
informs the price negotiation with the CEPS. The CT's opinions are advisory: the ministry
decides listing and the health insurance union (UNCAM) sets the reimbursement rate. For
some products the CEESP adds an efficiency opinion with an ICER in €/QALY. HAS also decides
**early access** (accès précoce) for presumed-innovative medicines in serious, rare or
disabling diseases that have no appropriate treatment.

**How outcomes are shown.** Early-access decisions appear as granted, renewed or refused,
and reimbursement opinions by ASMR level (or *SMR insufficient*). Early-access decisions
carry no SMR/ASMR rating or ICER, so those fields read "Not assessed" rather than being
inferred.

**What doesn't carry over from the NICE view.** France has no explicit cost-effectiveness
threshold, so there's no threshold-based risk signal. Where a CEESP ICER exists it's
shown in €/QALY as published — never converted to £ or compared with NICE's threshold.
"""

_AGENCY_CHAT_RULES = """
Rules, no exceptions:
1. Answer ONLY using the {records} provided below. Do not use outside knowledge of {bodies}, \
drugs or clinical practice beyond what's in this context, even if you know it.
2. If the provided records don't contain enough information to answer, say so plainly — \
never fill the gap with a plausible-sounding guess.
3. {economics}
4. When you draw a conclusion, name which record(s) it comes from (by ID, e.g. {example_id}) \
so it can be checked against the source.
5. Keep answers short and direct — a consultant is scanning this, not reading an essay.
6. This is a descriptive summary of past decisions, not a prediction of {prediction}. If \
asked to predict an outcome, say that's outside what this data can support.
7. Where the data lists the same molecule's decision at another HTA body, you may compare \
them, but say that NICE, the G-BA and HAS answer different questions (cost-effectiveness, \
added clinical benefit, clinical benefit / early access), so a different outcome is not a \
contradiction."""

GBA_CHAT_PROMPT = (
    "You are answering questions about a specific, already-selected set of G-BA early "
    "benefit assessments (AMNOG, §35a SGB V) for a market access consultant. You are NOT "
    "a general G-BA, IQWiG or market-access assistant.\n"
    + _AGENCY_CHAT_RULES.format(
        records="assessment data",
        bodies="the G-BA, IQWiG, AMNOG price negotiations",
        economics=("The G-BA rates added clinical benefit against an appropriate comparator "
                   "therapy; it does not assess cost-effectiveness. Never state an ICER, "
                   "price, reimbursement amount or percentage that isn't explicitly present "
                   "in the data below."),
        example_id="D-1305",
        prediction="a G-BA resolution or a negotiated price",
    )
)

HAS_CHAT_PROMPT = (
    "You are answering questions about a specific, already-selected set of HAS decisions "
    "(early access decisions, Transparency Committee opinions with SMR/ASMR ratings and/or "
    "CEESP efficiency opinions) for a market access consultant. You are NOT a general HAS "
    "or market-access assistant.\n"
    + _AGENCY_CHAT_RULES.format(
        records="decision data",
        bodies="HAS, the Transparency Committee, the CEESP, CEPS pricing",
        economics=("Never state an SMR or ASMR level, ICER, price or percentage that isn't "
                   "explicitly present in the data below. Early access decisions carry no "
                   "SMR/ASMR rating; don't infer one."),
        example_id="AP607",
        prediction="a HAS decision or a negotiated price",
    )
)

CROSS_AGENCY_CAVEAT = (
    "Matched on active substance (including combination regimens), not on indication — "
    "check the indication before comparing outcomes. The bodies answer different questions: "
    "NICE weighs cost-effectiveness for the NHS, the G-BA rates added clinical benefit "
    "against a comparator it sets, and HAS rates clinical benefit (SMR/ASMR) and decides "
    "early access. A different outcome is expected, not a contradiction."
)


# ── Agency configuration ─────────────────────────────────────────────────────

AGENCY_CONFIG = {
    "G-BA": {
        "sheet": "G-BA_Recent",
        "slug": "gba",
        "title": "🇩🇪 G-BA Benefit Assessment Intelligence",
        "noun": "assessments",
        "noun_title": "Assessments",
        "detail_title": "📋 Assessment Detail",
        "search_examples": "'breast', 'myasthenia'",
        "source_label": "official G-BA resolutions and IQWiG assessments",
        "log_keyword": "G-BA",
        "extra_columns": ["comparator_therapy", "added_benefit_assessment", "ICER_EUR_per_QALY"],
        "comparator_col": "comparator_therapy",
        "icer_lower_col": "ICER_EUR_per_QALY",
        "icer_upper_col": None,
        "outcome_fn": gba_outcome,
        "favourable_fn": gba_favourable,
        "derived": {"certainty": gba_certainty},
        "outcome_order": GBA_OUTCOME_ORDER,
        "outcome_colors": GBA_COLORS,
        "outcome_md_colors": GBA_MD_COLORS,
        "subtitle_fn": gba_subtitle,
        "headlines_fn": gba_headlines,
        "show_recorded_decision": True,
        "table_extra": [("certainty", "Evidence certainty")],
        "detail_fields": [
            ("Added benefit (as recorded)", "added_benefit_assessment"),
            ("Evidence certainty", "certainty"),
            ("Appropriate comparator therapy (zVT)", "comparator"),
            ("Patient population", "patient_population"),
            ("Reasoning", "decision_reasoning"),
            ("Economic evaluation", "ICER_status"),
        ],
        "synthesis_blocks": [
            {"title": "Appropriate comparator therapy (zVT) set by the G-BA",
             "cols": ["comparator"], "subset": "all",
             "caption": "The comparator the G-BA required for each assessment — the input that "
                        "most shapes a German dossier. Where it differs by subgroup, the "
                        "resolution lists each subgroup's comparator."},
            {"title": "Where an added benefit was found",
             "cols": ["added_benefit_assessment", "decision_reasoning"], "subset": "favourable",
             "caption": "Positive ratings, including mixed resolutions where at least one "
                        "subgroup received one.",
             "empty": "No positive rating in the current selection."},
            {"title": "Where no added benefit was proven",
             "cols": ["added_benefit_assessment", "decision_reasoning"], "subset": "unfavourable",
             "caption": "As recorded in the workbook. For the full rationale, open the G-BA "
                        "'Reasons for decision' (Tragende Gründe) linked in each assessment's "
                        "detail.",
             "empty": "No negative outcome in the current selection."},
            {"title": "Target population size",
             "cols": ["patient_population"], "subset": "all",
             "caption": "Patient numbers as stated in the G-BA resolution."},
        ],
        "chat_fields": [
            ("Assessment type", "assessment_type"),
            ("Appropriate comparator therapy", "comparator"),
            ("Added benefit", "added_benefit_assessment"),
            ("Evidence certainty", "certainty"),
            ("Patient population", "patient_population"),
            ("Reasoning", "decision_reasoning"),
            ("Economic evaluation", "ICER_status"),
        ],
        "chat_prompt": GBA_CHAT_PROMPT,
        "methodology_md": GBA_METHODOLOGY,
    },
    "HAS": {
        "sheet": "HAS_Recent",
        "slug": "has",
        "title": "🇫🇷 HAS Appraisal Intelligence",
        "noun": "decisions",
        "noun_title": "Decisions",
        "detail_title": "📋 Decision Detail",
        "search_examples": "'lymphoma', 'lung'",
        "source_label": "official HAS decisions and Transparency Committee documents",
        "log_keyword": "HAS",
        "extra_columns": ["comparator", "SMR_rating", "ASMR_rating", "ICER_lower_EUR_per_QALY",
                          "ICER_upper_EUR_per_QALY", "current_status"],
        "comparator_col": "comparator",
        "icer_lower_col": "ICER_lower_EUR_per_QALY",
        "icer_upper_col": "ICER_upper_EUR_per_QALY",
        "outcome_fn": has_outcome,
        "favourable_fn": has_favourable,
        "derived": {
            "_early_access": has_is_early_access,
            "_smr": lambda r: smr_level(r.get("SMR_rating")),
            "_asmr": lambda r: asmr_level(r.get("ASMR_rating")),
            "_ratings": has_ratings_text,
        },
        "outcome_order": HAS_OUTCOME_ORDER,
        "outcome_colors": HAS_COLORS,
        "outcome_md_colors": HAS_MD_COLORS,
        "subtitle_fn": has_subtitle,
        "headlines_fn": has_headlines,
        "show_recorded_decision": False,
        "table_extra": [("SMR_rating", "SMR"), ("ASMR_rating", "ASMR")],
        "detail_fields": [
            ("SMR", "SMR_rating"),
            ("ASMR", "ASMR_rating"),
            ("Comparator", "comparator"),
            ("Patient population", "patient_population"),
            ("Reasoning", "decision_reasoning"),
            ("Current status", "current_status"),
            ("Economic evaluation", "ICER_status"),
        ],
        "synthesis_blocks": [
            {"title": "Refusals and insufficient SMR",
             "cols": ["decision_reasoning"], "subset": "unfavourable",
             "caption": "Reasons as recorded from the public decision — see the linked HAS "
                        "documents for the full opinion.",
             "empty": "No refusal or insufficient SMR in the current selection."},
            {"title": "Where HAS granted, renewed or rated an improvement",
             "cols": ["decision_reasoning"], "subset": "favourable",
             "empty": "No favourable decision in the current selection."},
            {"title": "Status notes",
             "cols": ["current_status"], "subset": "all",
             "caption": "Later changes recorded against each decision — e.g. a product leaving "
                        "the early-access scheme."},
            {"title": "Clinical benefit ratings (SMR / ASMR)",
             "cols": ["_ratings"], "subset": "all",
             "empty": "None of the selected records carry SMR/ASMR ratings — early-access "
                      "decisions don't assign them."},
            {"title": "Comparator", "cols": ["comparator"], "subset": "all"},
        ],
        "chat_fields": [
            ("Assessment type", "assessment_type"),
            ("SMR", "SMR_rating"),
            ("ASMR", "ASMR_rating"),
            ("Comparator", "comparator"),
            ("Patient population", "patient_population"),
            ("Reasoning", "decision_reasoning"),
            ("Current status", "current_status"),
            ("Economic evaluation", "ICER_status"),
        ],
        "chat_prompt": HAS_CHAT_PROMPT,
        "methodology_md": HAS_METHODOLOGY,
    },
}


# ── Same-molecule matching across agencies ───────────────────────────────────

_MOLECULE_SEPARATORS = re.compile(
    r"\s*(?:\bin combination with\b|\bcombined with\b|\bfollowed by\b|\bwith or without\b"
    r"|\bwithout\b|\bwith\b|\bplus\b|\band\b|\bor\b|\bthen\b|\+|,|;|/|&|–|—|\s-\s)\s*",
    re.IGNORECASE)
# Words after which a drug-name part stops naming the drug ('Mepolizumab as an
# add-on…', 'Exenatide … for injection').
_DESCRIPTIVE_TAIL = re.compile(r"\b(?:for|as|in|to|after|before|versus|vs)\b")
_DESCRIPTOR_WORDS = {
    "monotherapy", "maintenance", "adjuvant", "neoadjuvant", "subcutaneous", "intravenous",
    "oral", "formulation", "tablet", "tablets", "capsule", "capsules", "injection",
    "infusion", "suspension", "solution", "prolonged", "immediate", "modified", "extended",
    "release", "high", "low", "dose", "first", "second", "third", "line", "therapy",
    "treatment", "regimen", "combination", "alone", "only", "originator", "biosimilar",
    "biosimilars", "an", "a", "the", "nd", "st", "rd", "th",
}
_SALT_WORDS = {
    "hydrochloride", "dihydrochloride", "mesylate", "mesilate", "maleate", "sodium",
    "potassium", "citrate", "tartrate", "besylate", "besilate", "fumarate",
    "hemifumarate", "succinate", "acetate", "phosphate", "sulfate", "sulphate",
    "bromide", "tosylate", "malate",
}
_NON_MOLECULES = {
    "chemotherapy", "chemoradiotherapy", "radiotherapy", "placebo", "best supportive care",
    "standard care", "standard of care", "endocrine", "hormone", "surgery", "platinum",
    "platinum based chemotherapy", "platinum containing chemotherapy", "taxane",
    "aromatase inhibitor", "sulphonylurea", "sulfonylurea",
}
_INDICATION_STOPWORDS = {
    "adults", "adult", "with", "who", "are", "the", "for", "and", "after", "least", "one",
    "more", "prior", "treating", "treatment", "treatments", "treated", "therapy",
    "therapies", "patients", "people", "previously", "disease", "not", "whose", "has",
    "have", "been", "than", "from", "that", "which", "this", "these", "their", "its", "per",
    "all", "any", "other", "such", "where", "when", "including", "only", "also", "both",
    "either", "specified", "used", "use", "had", "aged", "years", "older", "over", "under",
    "can", "cannot", "unsuitable", "suitable", "eligible", "ineligible", "listed",
    "standard", "options", "option", "systemic", "line", "lines", "first", "second", "third",
}


def molecule_set(*names):
    """
    Active substances named in drug-name strings, as letters-only keys.

    NICE names regimens and formulations in many ways — 'Trastuzumab in
    combination with paclitaxel', 'Trastuzumab monotherapy', 'Pembrolizumab
    plus chemotherapy with or without bevacizumab', 'Nivolumab–relatlimab',
    'High-dose imatinib' — and its generic_name column carries an extraction
    note in brackets. So names are split on regimen separators, bracketed text
    and descriptive tails are dropped, formulation words ignored, and each
    remaining part becomes a letters-only key: 'Lutetium-177 vipivotide
    tetraxetan' and 'Lutetium (177Lu) vipivotide tetraxetan' give the same key.
    Matching is on whole keys, so 'trastuzumab' never matches 'trastuzumab
    deruxtecan'.
    """
    keys = set()
    for name in names:
        s = _text(name)
        if not s:
            continue
        s = re.sub(r"\([^)]*\)|\([^)]*$", " ", s.lower())
        for part in _MOLECULE_SEPARATORS.split(s):
            part = _DESCRIPTIVE_TAIL.split(part, maxsplit=1)[0]
            words = [w for w in re.findall(r"[a-z]+", part)
                     if w not in _SALT_WORDS and w not in _DESCRIPTOR_WORDS]
            key = "".join(words)
            if len(key) >= 4 and " ".join(words) not in _NON_MOLECULES:
                keys.add(key)
    return frozenset(keys)


def _indication_tokens(text):
    s = _text(text)
    if not s:
        return set()
    return {w[:5] for w in re.findall(r"[a-z0-9]+", s.lower())
            if len(w) >= 3 and w not in _INDICATION_STOPWORDS}


def indication_overlap(a, b):
    """Rough 0–1 overlap of two indication texts — used only to list the closest
    same-molecule appraisal first, never to include or exclude one."""
    ta, tb = _indication_tokens(a), _indication_tokens(b)
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / len(ta | tb)


# ── Frames ───────────────────────────────────────────────────────────────────

# Columns every agency view reads. Missing ones are added empty, so renaming or
# dropping a column in the workbook degrades gracefully instead of crashing a
# view — including the NICE view, which reads these rows for "Also assessed by".
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
    outcome, favourable, area_broad, search_blob, _molecules, plus any
    agency-specific derived fields. The row index is kept as the record key,
    so two rows sharing an ID (e.g. an initial decision and its renewal) stay
    two records everywhere."""
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
        _text(g) or (_text(d).title() if _text(d) else "Unknown")
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

    frame["outcome"] = frame.apply(cfg["outcome_fn"], axis=1)
    for name, fn in cfg["derived"].items():
        frame[name] = frame.apply(fn, axis=1)
    frame["favourable"] = frame.apply(cfg["favourable_fn"], axis=1)

    frame["area_broad"] = [
        (_text(a) or "Not classified").split(" / ")[0].strip() for a in frame["therapeutic_area"]
    ]
    blob_cols = ("appraisal_id", "inn", "brand", "drug_name", "indication", "therapeutic_area")
    frame["search_blob"] = [
        " | ".join(_text(v) or "" for v in values).lower()
        for values in zip(*(frame[c] for c in blob_cols))
    ]
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


def format_eur_icer(row):
    lo, hi = row.get("icer_eur_lower"), row.get("icer_eur_upper")
    if lo is None or pd.isna(lo):
        return None
    if hi is None or pd.isna(hi) or float(hi) == float(lo):
        return f"€{float(lo):,.0f}/QALY"
    return f"€{float(lo):,.0f}–€{float(hi):,.0f}/QALY"


def outcome_order(agency, frame):
    """Outcome categories present in the data, in the agency's scale order."""
    order = AGENCY_CONFIG[agency]["outcome_order"]
    present = set(frame["outcome"].dropna())
    return [o for o in order if o in present] + sorted(present - set(order), key=str.lower)


# ── Sidebar ──────────────────────────────────────────────────────────────────

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
    st.sidebar.divider()
    return choice


def agency_sidebar_filters(agency, frame):
    cfg = AGENCY_CONFIG[agency]
    slug = cfg["slug"]
    st.sidebar.title("🔍 Filters")

    search = st.sidebar.text_input(
        "Search",
        placeholder="Drug, brand, indication or ID",
        key=f"w_{slug}_search",
        help="Searches INN, brand name, indication, therapy area and procedure ID. "
             "Abbreviations such as NSCLC, CRC or UC are expanded to the matching disease term.",
    )
    areas = sorted(frame["area_broad"].dropna().unique(), key=str.lower)
    selected_areas = st.sidebar.multiselect(
        "Therapy area", areas, key=f"w_{slug}_areas",
        help="Leave empty to include all areas. Grouped on the first part of the workbook "
             "label (e.g. 'Oncology / haematology' → Oncology).",
    )
    selected_outcomes = st.sidebar.multiselect(
        "Outcome", outcome_order(agency, frame), key=f"w_{slug}_outcomes",
        help="Leave empty to include all outcomes.",
    )
    types = (sorted({t for t in frame["assessment_type"].map(_text).dropna()}, key=str.lower)
             if "assessment_type" in frame.columns else [])
    selected_types = []
    if len(types) > 1:
        selected_types = st.sidebar.multiselect(
            "Assessment type", types, key=f"w_{slug}_types",
            help="Leave empty to include all procedure types.")

    years = frame["year"].dropna()
    year_range, full_range = None, None
    if len(years) and int(years.min()) < int(years.max()):
        full_range = (int(years.min()), int(years.max()))
        init_range_state(f"w_{slug}_years", *full_range)
        year_range = st.sidebar.slider(
            "Decision year range", min_value=full_range[0], max_value=full_range[1],
            step=1, key=f"w_{slug}_years",
            help="Calendar year of the decision (publication date where no decision date "
                 "is recorded).")
    elif len(years):
        st.sidebar.caption(
            f"All records are dated {int(years.min())}. A year-range filter appears once the "
            f"data spans more than one year.")

    result = frame
    if search and search.strip():
        needle = re.sub(r"\s+", " ", search.strip().lower())
        resolved, _ = resolve_keyword(search)
        mask = result["search_blob"].str.contains(needle, regex=False)
        if resolved and resolved != needle:
            mask = mask | result["search_blob"].str.contains(resolved, regex=False)
        result = result[mask]
    if selected_areas:
        result = result[result["area_broad"].isin(selected_areas)]
    if selected_outcomes:
        result = result[result["outcome"].isin(selected_outcomes)]
    if selected_types:
        result = result[result["assessment_type"].isin(selected_types)]
    if year_range and tuple(year_range) != full_range:
        result = result[result["year"].between(year_range[0], year_range[1])]
    return result


# ── Record detail & synthesis ────────────────────────────────────────────────

def render_agency_record(agency, row, docs, hits):
    cfg = AGENCY_CONFIG[agency]
    color = cfg["outcome_md_colors"].get(row["outcome"], "gray")
    label = _md_safe(str(row["outcome"]).replace("[", "(").replace("]", ")"))
    st.markdown(f"**Outcome:** :{color}[**{label}**]")
    recorded = _text(row.get("decision_simple"))
    if recorded and recorded != row["outcome"]:
        st.caption(f"Recorded as: {_md_safe(recorded)}")

    meta = [f"**Decision date:** {row['date_label']}"]
    published = _text(row.get("publication_date"))
    if published and published != _text(row.get("decision_date")):
        meta.append(f"**Published:** {_md_safe(published)}")
    meta.append(f"**Assessment type:** "
                f"{_md_safe(_text(row.get('assessment_type')) or 'Not recorded')}")
    st.markdown(" · ".join(meta))

    for label, col in cfg["detail_fields"]:
        value = _text(row.get(col))
        if value:
            st.markdown(f"**{label}:** {_md_safe(value)}")
    icer = format_eur_icer(row)
    if icer:
        st.markdown(f"**Published ICER:** {icer}")
    note = _text(row.get("methodology_note"))
    if note:
        st.caption(_md_safe(note))

    record_docs = docs[docs["Appraisal_ID"] == str(row["appraisal_id"])]
    links = [
        f"- {_md_link(d['Document_Type'], d['Source_URL'])}"
        + (f" · {d['Date']}" if _text(d["Date"]) else "")
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
            st.markdown(f"**{row['appraisal_id']} — {row['inn']}** · "
                        f"_{_md_safe(_shorten(row['indication'], 90))}_\n\n"
                        f"> {_md_safe(_shorten(text, 450))}")
        if total > SYNTHESIS_MAX_ITEMS:
            st.caption(f"Showing the first {SYNTHESIS_MAX_ITEMS} of {total} — narrow the "
                       f"selection with the sidebar to see the rest.")


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


# ── Charts ───────────────────────────────────────────────────────────────────

def _integer_axis(fig, axis, max_value):
    """Whole-number ticks — counts of 1–5 otherwise get 0.5 steps."""
    update = fig.update_yaxes if axis == "y" else fig.update_xaxes
    update(rangemode="tozero", tickformat=",d", dtick=1 if max_value <= 8 else None)


def render_agency_charts(agency, frame, full_frame):
    cfg = AGENCY_CONFIG[agency]
    colors = cfg["outcome_colors"]
    order = outcome_order(agency, full_frame)

    left, right = st.columns(2)
    with left:
        st.markdown("**Outcome breakdown**")
        fig = px.pie(frame, names="outcome", color="outcome", color_discrete_map=colors,
                     hole=0.4, category_orders={"outcome": order})
        st.plotly_chart(fig, width="stretch", key=f"{cfg['slug']}_outcome_pie")
    with right:
        st.markdown(f"**{cfg['noun_title']} over time**")
        dated = frame.dropna(subset=["decision_dt"])
        if dated.empty:
            st.caption("No decision dates recorded in this selection.")
        else:
            by_year = dated["decision_dt"].dt.year.nunique() >= 2
            periods = dated["decision_dt"].dt.to_period("Y" if by_year else "M")
            counts = (
                dated.assign(_period=periods,
                             period=periods.dt.strftime("%Y" if by_year else "%b %Y"))
                .groupby(["_period", "period", "outcome"]).size()
                .reset_index(name="count").sort_values("_period")
            )
            fig = px.bar(counts, x="period", y="count", color="outcome",
                         color_discrete_map=colors,
                         category_orders={"outcome": order,
                                          "period": list(dict.fromkeys(counts["period"]))})
            fig.update_layout(barmode="stack", xaxis_title=None, yaxis_title="Count",
                              legend_title=None, bargap=0.5)
            _integer_axis(fig, "y", counts.groupby("period")["count"].sum().max())
            st.plotly_chart(fig, width="stretch", key=f"{cfg['slug']}_over_time")

    st.markdown("**Outcome by therapy area**")
    by_area = frame.groupby(["area_broad", "outcome"]).size().reset_index(name="count")
    fig = px.bar(by_area, x="count", y="area_broad", color="outcome", orientation="h",
                 color_discrete_map=colors, category_orders={"outcome": order})
    fig.update_layout(barmode="stack", yaxis={"categoryorder": "total ascending"},
                      yaxis_title=None, xaxis_title="Count", legend_title=None, bargap=0.45,
                      height=max(260, 48 * by_area["area_broad"].nunique() + 140))
    _integer_axis(fig, "x", by_area.groupby("area_broad")["count"].sum().max())
    st.plotly_chart(fig, width="stretch", key=f"{cfg['slug']}_by_area")

    if agency == "G-BA":
        rated = frame[frame["certainty"].notna() & frame["favourable"].eq(True)]
        if len(rated):
            st.markdown("**Positive ratings by certainty of evidence**")
            levels = [c for c in ("Proof", "Indication", "Hint") if c in set(rated["certainty"])]
            table = pd.crosstab(rated["outcome"], rated["certainty"]).reindex(columns=levels)
            table = table.reindex([o for o in order if o in table.index])
            st.dataframe(table.rename_axis(index="Outcome", columns=None),
                         width=150 + 110 * len(levels) + 180)
            st.caption("For mixed resolutions, certainty refers to the subgroup that received "
                       "a positive rating.")

    if agency == "HAS":
        smr = frame[frame["_smr"].notna()]
        asmr = frame[frame["_asmr"].notna()]
        if len(smr) or len(asmr):
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**SMR (actual clinical benefit)**")
                if len(smr):
                    counts = smr["_smr"].value_counts().reindex(
                        [s for s in SMR_ORDER if s in set(smr["_smr"])]
                        + sorted(set(smr["_smr"]) - set(SMR_ORDER))).reset_index()
                    counts.columns = ["SMR", "Count"]
                    fig = px.bar(counts, x="SMR", y="Count", color_discrete_sequence=["#2a78d6"])
                    _integer_axis(fig, "y", counts["Count"].max())
                    st.plotly_chart(fig, width="stretch", key="has_smr")
                else:
                    st.caption("No SMR ratings in this selection.")
            with c2:
                st.markdown("**ASMR (improvement in clinical benefit)**")
                if len(asmr):
                    levels = [lv for lv in ("I", "II", "III", "IV", "V") if lv in set(asmr["_asmr"])]
                    counts = asmr["_asmr"].value_counts().reindex(levels).reset_index()
                    counts.columns = ["ASMR", "Count"]
                    counts["ASMR"] = "ASMR " + counts["ASMR"]
                    fig = px.bar(counts, x="ASMR", y="Count", color="ASMR",
                                 color_discrete_map=colors)
                    fig.update_layout(showlegend=False)
                    _integer_axis(fig, "y", counts["Count"].max())
                    st.plotly_chart(fig, width="stretch", key="has_asmr")
                else:
                    st.caption("No ASMR ratings in this selection.")


# ── Grounded chat over the current selection ─────────────────────────────────

def _anthropic_api_key():
    try:
        return st.secrets.get("ANTHROPIC_API_KEY")
    except Exception:
        # No secrets.toml at all is the normal state before chat is configured.
        return None


def _matches_by_record(matches):
    """{record row index: its same-molecule matches}, from the long match table."""
    return {src: group for src, group in matches.groupby("_src", sort=False)}


def build_agency_chat_context(agency, frame, matches):
    """Same shape as the NICE chat context: an outcome mix for pattern
    questions, then every selected record (up to the cost ceiling), each with
    the same molecule's decisions at other agencies so cross-country questions
    can be answered from the data rather than from memory."""
    cfg = AGENCY_CONFIG[agency]
    by_record = _matches_by_record(matches)
    rows = frame.head(CHAT_MAX_CONTEXT_ROWS)
    field_chars = CHAT_FIELD_CHARS_FULL if len(rows) <= 20 else CHAT_FIELD_CHARS_TIGHT
    mix = ", ".join(f"{k}: {v}" for k, v in frame["outcome"].value_counts().items())
    truncated = len(frame) - len(rows)
    coverage = (f"Selected {agency} records — all {len(rows)}:" if truncated <= 0 else
                f"Selected {agency} records ({len(rows)} of {len(frame)}; {truncated} omitted "
                f"for length — say so if a question needs them):")

    blocks = [
        f"{agency} {cfg['noun']} currently selected: {len(frame)}. Outcome mix: {mix}.",
        "",
        coverage,
    ]
    for idx, row in rows.iterrows():
        brand = f" ({row['brand']})" if row["brand"] else ""
        parts = [f"[{row['appraisal_id']}] {row['inn']}{brand} — {row['indication']} — "
                 f"Outcome: {row['outcome']} (recorded as: "
                 f"{_text(row.get('decision_simple')) or 'not recorded'}; decision date "
                 f"{row['date_label']})"]
        for label, col in cfg["chat_fields"]:
            value = _text(row.get(col))
            if value:
                parts.append(f"  {label}: {value[:field_chars]}")
        icer = format_eur_icer(row)
        if icer:
            parts.append(f"  Published ICER: {icer}")
        hits = by_record.get(idx)
        if hits is not None:
            for _, h in hits.head(4).iterrows():
                parts.append(
                    f"  Same molecule at {h['body']}: [{h['appraisal_id']}] {h['drug']} — "
                    f"{_shorten(h['indication'], 140)} — {h['decision']} ({h['when']})")
        blocks.append("\n".join(parts))
    return "\n\n".join(blocks)


def render_agency_chat(agency, frame, matches):
    cfg = AGENCY_CONFIG[agency]
    slug = cfg["slug"]
    st.divider()
    st.subheader("💬 Ask about this selection")

    api_key = _anthropic_api_key()
    if not api_key:
        st.info("Chat isn't configured yet — needs an ANTHROPIC_API_KEY in this app's "
                "Streamlit secrets.")
        return

    # The selection IS the grounding data, so a different selection starts a
    # fresh conversation — the NICE chat resets on a new query for the same reason.
    history_key, signature_key = f"chat_history_{slug}", f"chat_signature_{slug}"
    signature = tuple(frame.index)
    if st.session_state.get(signature_key) != signature:
        st.session_state[history_key] = []
        st.session_state[signature_key] = signature
    history = st.session_state[history_key]

    st.caption(
        f"Ask about the {len(frame)} {agency} {cfg['noun']} currently selected by the sidebar "
        f"filters — e.g. \"which comparators came up most?\" or \"how did NICE decide on the "
        f"same molecules?\". Answers are grounded only in this selection and will say so if "
        f"something isn't covered, rather than guessing.")
    turns_used = len(history) // 2
    st.caption(f"{turns_used}/{CHAT_MAX_TURNS} questions used for this selection.")

    for msg in history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if turns_used >= CHAT_MAX_TURNS:
        st.warning("Question limit reached for this selection. Change the filters to start "
                   "a new conversation.")
        return

    question = st.chat_input(f"Ask about these {agency} {cfg['noun']}...",
                             key=f"chat_input_{slug}")
    if question:
        history.append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.markdown(question)
        context = build_agency_chat_context(agency, frame, matches)
        with st.chat_message("assistant"):
            with st.spinner(f"Checking the selected {agency} {cfg['noun']}..."):
                answer, error = ask_chat(api_key, context, history[:-1], question,
                                         system_prompt=cfg["chat_prompt"])
            if error:
                st.error(error)
                history.pop()  # don't count a failed turn
            else:
                st.markdown(answer)
                history.append({"role": "assistant", "content": answer})


# ── The agency page ──────────────────────────────────────────────────────────

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


def render_agency_view(agency):
    cfg = AGENCY_CONFIG[agency]
    version = data_version()
    frame = prepare_agency_frame(agency, version)

    st.title(cfg["title"])
    if frame.empty:
        st.info(f"No {agency} records in this workbook yet — add rows to the "
                f"'{cfg['sheet']}' sheet and they'll appear here with no code changes.")
        return

    doc_index = load_document_index(version)
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
            f"**Coverage.** {total} {cfg['noun']} dated {_date_span(frame['decision_dt'])} · "
            f"{len(docs)} official source documents indexed for {documented} of them · the "
            f"same active substance also appears in the NICE data for "
            f"{_records_matched(all_matches, 'NICE')} of {total}.")
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

    st.divider()

    filtered = agency_sidebar_filters(agency, frame)
    matches = all_matches[all_matches["_src"].isin(filtered.index)]

    # Metrics — one card per outcome on this agency's scale, like the NICE view
    st.metric(f"Total {cfg['noun_title']}", len(filtered))
    outcomes = outcome_order(agency, frame)
    for col, outcome in zip(st.columns(len(outcomes)), outcomes):
        col.metric(outcome, int((filtered["outcome"] == outcome).sum()))
    n = len(filtered)
    headlines = cfg["headlines_fn"](filtered) + [(
        "Same molecule appraised by NICE", f"{_records_matched(matches, 'NICE')}/{n}",
        "Selected records whose active substance also appears in the NICE data — see "
        "'Same molecule at other HTA bodies' below.")]
    for col, (label, value, help_text) in zip(st.columns(len(headlines)), headlines):
        col.metric(label, value, help=help_text)

    st.divider()

    st.download_button(
        "📥 Download Filtered Results",
        data=_agency_export_bytes(agency, version, tuple(filtered.index)),
        file_name=f"{cfg['slug']}_filtered.xlsx",
        mime=XLSX_MIME,
        key=f"download_{cfg['slug']}",
    )
    st.caption(f"Showing {n:,} of {total:,} {cfg['noun']}")

    if filtered.empty:
        st.info("No records match the current filters.")
        _agency_footer(agency, total)
        return

    # Outcome sits next to the drug: the indication text is long enough to push
    # anything after it off-screen.
    table = pd.DataFrame({
        "ID": filtered["appraisal_id"],
        "Drug (INN)": filtered["inn"],
        "Brand": filtered["brand"],
        "Outcome": filtered["outcome"],
        "Decision date": filtered["decision_dt"],
        "Indication": filtered["indication"],
        "Therapy Area": filtered["therapeutic_area"],
    })
    # The recorded wording is always worth seeing next to an 'Unclassified' bucket.
    if cfg["show_recorded_decision"] or filtered["outcome"].eq("Unclassified").any():
        table["Decision (as recorded)"] = filtered["decision_simple"]
    for col, label in cfg["table_extra"]:
        if filtered[col].map(_substantive).notna().any():
            table[label] = filtered[col].map(lambda v: _text(v) or "—")
    table["Source"] = filtered["official_source_url"]
    st.dataframe(
        table,
        column_config={
            "Decision date": st.column_config.DateColumn("Decision date", format="D MMM YYYY"),
            "Indication": st.column_config.TextColumn("Indication", width="large"),
            "Source": st.column_config.LinkColumn("Source", display_text="Open ↗"),
        },
        width="stretch", hide_index=True,
    )

    # Detail
    by_record = _matches_by_record(matches)
    st.divider()
    st.subheader(cfg["detail_title"])
    labels = pd.Series(
        [f"{i} ({b})" if b else i for i, b in zip(filtered["inn"], filtered["brand"])],
        index=filtered.index)
    options = sorted(labels.unique(), key=str.lower)
    selected = st.selectbox("Select a drug", options, key=f"select_{cfg['slug']}_drug")
    chosen = filtered[labels == selected].sort_values("decision_dt", ascending=False)
    for idx, row in chosen.iterrows():
        with st.expander(f"{_md_safe(row['appraisal_id'])} — "
                         f"{_md_safe(_shorten(row['indication'], 110))}",
                         expanded=len(chosen) == 1):
            render_agency_record(agency, row, docs, by_record.get(idx))

    # Analysis
    st.divider()
    st.subheader("📊 Analysis")
    render_agency_charts(agency, filtered, frame)

    # Evidence synthesis
    st.divider()
    st.subheader("🔎 Evidence Synthesis")
    st.markdown(f"*What the {n} {cfg['noun']} in the current selection say. Narrow it with "
                f"the sidebar search (e.g. {cfg['search_examples']}) to synthesise one "
                f"indication.*")
    st.caption(_coverage_line(filtered, docs))
    for block in cfg["synthesis_blocks"]:
        render_synthesis_block(block, filtered)
    render_economic_block(filtered)

    # Same molecule across agencies. Widely appraised molecules (pembrolizumab
    # has dozens of NICE TAs) would swamp the table, so by default each record
    # shows only its closest-indication matches.
    st.divider()
    st.subheader("🌍 Same Molecule at Other HTA Bodies")
    per_record = 3
    show_all = False
    if len(matches) and matches.groupby("_src").size().max() > per_record:
        show_all = st.checkbox(
            f"Show every same-molecule match (default: the {per_record} closest indications "
            f"per record)", key=f"w_{cfg['slug']}_all_matches")
    shown = matches if show_all else matches.groupby("_src", sort=False).head(per_record)
    if len(shown):
        source = filtered.loc[shown["_src"]]
        comparison = pd.DataFrame({
            "Molecule": source["inn"].to_numpy(),
            f"{agency} ID": source["appraisal_id"].to_numpy(),
            f"{agency} outcome": source["outcome"].to_numpy(),
            "Other body": [f"{AGENCY_FLAGS.get(b, '')} {b}" for b in shown["body"]],
            "Other ID": shown["appraisal_id"].to_numpy(),
            "Other drug / regimen": shown["drug"].to_numpy(),
            "Other decision": shown["decision"].to_numpy(),
            "Other indication": shown["indication"].to_numpy(),
            "When": shown["when"].to_numpy(),
            "Link": shown["url"].to_numpy(),
        })
        st.markdown(f"**{_records_matched(matches)} of {n}** selected {cfg['noun']} have an "
                    f"appraisal of the same active substance at another HTA body.")
        st.dataframe(
            comparison,
            column_config={"Link": st.column_config.LinkColumn("Link", display_text="Open ↗")},
            width="stretch", hide_index=True,
        )
        st.caption(CROSS_AGENCY_CAVEAT)
    else:
        st.info("None of the selected molecules appear at another HTA body in this workbook yet.")

    render_agency_chat(agency, filtered, matches)
    _agency_footer(agency, total)


# ── HTA body switcher ────────────────────────────────────────────────────────
# Rendered first so it sits at the top of the sidebar. The NICE page below is
# the original script; the other agencies render their own page and stop the
# run here, so none of the NICE code executes for them.

HTA_BODY = render_agency_switcher()
if HTA_BODY != "NICE":
    render_agency_view(HTA_BODY)
    st.stop()


# Header

st.title("💊 NICE Technology Appraisal Intelligence")
st.markdown(f"*{TOTAL_ROWS:,} pharmaceutical appraisals — complete NICE database*")

tagged_total = int(df["line_of_therapy"].notna().sum())
with st.expander("Scope & coverage — read before benchmarking"):
    tagged_areas = (
        df[df["line_of_therapy"].notna()]["therapeutic_area"]
        .value_counts()
        .drop(labels=["Other / Multiple therapy areas"], errors="ignore")
        .head(5).index.tolist()
    )
    st.markdown(f"""
This tool operates at two levels of depth.

**Precedent browsing** works across all **{TOTAL_ROWS:,}** appraisals — search,
filter, decision history, rejection reasoning, and links to NICE guidance.

**Weighted similarity benchmarking** requires structured tags (line of therapy,
mechanism of action, biomarker, comparator type). These are currently populated
for **{tagged_total}** appraisals, concentrated in {', '.join(tagged_areas)}.
Outside those areas, the tool falls back to indication-keyword retrieval and says so.
Where a row has been through this tagging process, the specific basis for each tag
is available — look for "How this row's similarity tags were assigned" wherever
that appraisal appears.

**Curated reliability approach:** rather than relying solely on automated AI extraction
at scale, this dataset uses a curated review layer for key appraisal fields and retains
source/provenance information where available. Automated extraction is treated as a
starting point, not verification in itself. The aim is to prioritise traceability and
auditability over maximising the number of fields extracted automatically.

**Disease area (Therapy Area filter):** this classification covers all
{TOTAL_ROWS:,} rows and predates the tagging pass above — it isn't independently
documented in this dataset. Treat it as a useful first-pass filter rather than a
verified clinical classification, and cross-check against the indication text or
the linked NICE guidance for anything you're relying on.

**Guidance status:** {int((df['guidance_status'] == 'Replaced / withdrawn').sum())}
appraisals have an explicit note that they were later replaced or withdrawn, and
{int((df['guidance_status'] == 'Current').sum())} are confirmed still current.
The remaining {int((df['guidance_status'] == 'Not checked').sum())} were never
checked for supersession — the sidebar filter hides only the confirmed-replaced
set by default, since "not checked" and "confirmed current" are different claims.

**Published ICERs** exist for **{int(df['icer_lower'].notna().sum())}** appraisals.
Most modern NICE appraisals withhold cost-effectiveness results under confidential
commercial arrangements. Where that applies, the numeric field is deliberately left
empty and the qualitative position is recorded separately — a NICE threshold is
never encoded as if it were an observed ICER.
    """)
    log = load_enrichment_log(data_version())
    if log is not None:
        st.markdown("**Enrichment log**")
        # Cast to string: the log mixes counts and words ('Present') in one
        # column, which Arrow cannot serialise as a single type.
        st.dataframe(log.astype(str), width="stretch", hide_index=True)

st.divider()


# Sidebar filters

st.sidebar.title("🔍 Filters")

search = st.sidebar.text_input(
    "Search",
    placeholder="Drug, brand, indication or TA ID",
    help="Searches generic name, brand name, indication and appraisal ID.",
    key="w_nice_search",
)

therapeutic_areas = sorted(df["therapeutic_area"].dropna().astype(str).unique())
selected_areas = st.sidebar.multiselect(
    "Therapy / disease area", therapeutic_areas,
    help="Leave empty to include all areas.",
    key="w_nice_areas",
)

decisions = sorted(df["decision_simple"].dropna().astype(str).unique(), key=str.lower)
selected_decisions = st.sidebar.multiselect(
    "Decision", decisions, help="Leave empty to include all decisions.",
    key="w_nice_decisions",
)

st.session_state.setdefault("w_nice_hide_replaced", True)
hide_replaced = st.sidebar.checkbox(
    "Hide replaced / withdrawn guidance",
    key="w_nice_hide_replaced",
    help=(
        f"{int((df['guidance_status'] == 'Replaced / withdrawn').sum())} appraisals have an "
        f"explicit replacement or withdrawal note and are hidden by default. "
        f"{int((df['guidance_status'] == 'Not checked').sum())} appraisals were never checked "
        f"for supersession — these stay visible either way, since 'not checked' isn't the same "
        f"claim as 'confirmed current.'"
    ),
)

year_min = int(df["year_start"].min())
year_max = int(df["year_start"].max())
init_range_state("w_nice_years", year_min, year_max)
year_range = st.sidebar.slider(
    "Appraisal year range", min_value=year_min, max_value=year_max,
    step=1, key="w_nice_years",
    help="NICE fiscal years, mapped to their start year (2024/25 shows as 2024).",
)

st.sidebar.caption(
    f"Showing {year_range[0]}–{year_range[1]}. "
    "Use the handles to select a span such as the last five years."
)

filtered_df = df.copy()
if search:
    needle = re.sub(r"\s+", " ", search.strip().lower())
    filtered_df = filtered_df[
        filtered_df["search_blob"].str.contains(needle, case=False, na=False, regex=False)
    ]
if selected_areas:
    filtered_df = filtered_df[filtered_df["therapeutic_area"].isin(selected_areas)]
if selected_decisions:
    filtered_df = filtered_df[filtered_df["decision_simple"].isin(selected_decisions)]
if hide_replaced:
    filtered_df = filtered_df[filtered_df["guidance_status"] != "Replaced / withdrawn"]
filtered_df = filtered_df[
    filtered_df["year_start"].between(year_range[0], year_range[1])
]


# Metrics — one card per decision category so they reconcile

st.metric("Total Appraisals", len(filtered_df))
metric_cols = st.columns(len(decisions))
for col, decision in zip(metric_cols, decisions):
    with col:
        st.metric(decision, int((filtered_df["decision_simple"] == decision).sum()))

st.divider()

export_cols = [c for c in filtered_df.columns if c != "search_blob"]
buffer = io.BytesIO()
filtered_df[export_cols].to_excel(buffer, index=False)
st.download_button(
    "📥 Download Filtered Results",
    data=buffer.getvalue(),
    file_name="nice_filtered.xlsx",
    mime=XLSX_MIME,
)

st.caption(f"Showing {len(filtered_df):,} of {TOTAL_ROWS:,} appraisals")

table_cols = ["appraisal_id", "drug_name", "brand_name", "indication",
              "therapeutic_area", "decision_simple", "year_label", "url"]
st.dataframe(
    filtered_df[table_cols].rename(columns={
        "appraisal_id": "TA ID",
        "drug_name": "Drug (INN)",
        "brand_name": "Brand",
        "indication": "Indication",
        "therapeutic_area": "Therapy Area",
        "decision_simple": "Decision",
        "year_label": "Year",
        "url": "NICE Link",
    }),
    width="stretch", hide_index=True,
)

st.divider()
st.subheader("📋 Drug Detail")
drug_options = sorted(filtered_df["drug_name"].dropna().unique().tolist())
if drug_options:
    selected_drug = st.selectbox("Select a drug", drug_options)
    render_nice_cross_agency(filtered_df[filtered_df["drug_name"] == selected_drug])
    for _, row in filtered_df[filtered_df["drug_name"] == selected_drug].iterrows():
        with st.expander(f"{row['appraisal_id']} - {row['indication']}"):
            c1, c2, c3 = st.columns(3)
            c1.metric("Decision", row["decision_simple"])
            c2.metric("Year", row["year_label"])
            c3.metric("Appraisal Type", row["appraisal_type"])
            brand = row.get("brand_name")
            if pd.notna(brand) and str(brand).strip().lower() != "not specified":
                st.caption(f"Brand name: {brand}")

            status_note = row.get("current_guidance_status")
            if row["guidance_status"] == "Replaced / withdrawn" and pd.notna(status_note):
                st.warning(f"Guidance status: {status_note}")
            elif row["guidance_status"] == "Current":
                st.caption("Guidance status: current — no replacement or withdrawal recorded.")
            else:
                st.caption("Guidance status: not checked for supersession in this dataset.")

            tag_basis = row.get("tag_basis")
            if pd.notna(tag_basis):
                with st.expander("How this row's similarity tags were assigned"):
                    st.caption(tag_basis)

            if pd.notna(row.get("rejection_reasoning")):
                st.caption(str(row["rejection_reasoning"])[:300])
            st.markdown(f"[View NICE Guidance]({row['url']})")
else:
    st.info("No appraisals match the current filters.")


# Analysis charts

st.divider()
st.subheader("📊 Analysis")

COLORS = {
    "Recommended": "#2ecc71",
    "Not Recommended": "#e74c3c",
    "Managed Access": "#f39c12",
    "Optimised": "#3498db",
    "Terminated": "#95a5a6",
    "Only in Research": "#9b59b6",
}

if filtered_df.empty:
    st.info("No data to chart with the current filters.")
else:
    col_left, col_right = st.columns(2)
    with col_left:
        st.markdown("**Decision Breakdown**")
        st.plotly_chart(
            px.pie(filtered_df, names="decision_simple", color="decision_simple",
                   color_discrete_map=COLORS, hole=0.4),
            width="stretch",
        )

    with col_right:
        st.markdown("**Approvals Over Time**")
        # Sorted on numeric year_start, not the label string — otherwise
        # publication-dated rows sort alphabetically ahead of the fiscal years.
        yearly = (
            filtered_df[filtered_df["decision_simple"].isin(
                ["Recommended", "Not Recommended", "Managed Access", "Optimised"])]
            .groupby(["year_start", "year_label", "decision_simple"])
            .size().reset_index(name="count")
            .sort_values("year_start")
        )
        fig2 = px.line(yearly, x="year_label", y="count", color="decision_simple",
                       color_discrete_map=COLORS, markers=True)
        fig2.update_layout(
            xaxis_tickangle=-45,
            xaxis={"categoryorder": "array",
                   "categoryarray": yearly.sort_values("year_start")["year_label"].unique()},
        )
        st.plotly_chart(fig2, width="stretch")

    st.markdown("**Top 15 Indications**")
    top_ind = filtered_df["indication"].value_counts().head(15).reset_index()
    top_ind.columns = ["Indication", "Count"]
    fig3 = px.bar(top_ind, x="Count", y="Indication", orientation="h",
                  color_discrete_sequence=["#3498db"])
    fig3.update_layout(yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig3, width="stretch")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**STA vs MTA**")
        tc = filtered_df["appraisal_type"].value_counts().reset_index()
        tc.columns = ["Type", "Count"]
        st.plotly_chart(px.pie(tc, values="Count", names="Type", hole=0.4),
                        width="stretch")
    with c2:
        st.markdown("**Decision by Appraisal Type**")
        td = filtered_df.groupby(["appraisal_type", "decision_simple"]).size().reset_index(name="count")
        st.plotly_chart(
            px.bar(td, x="appraisal_type", y="count", color="decision_simple",
                   color_discrete_map=COLORS, barmode="stack"),
            width="stretch",
        )


# HTA Evidence Explorer

st.divider()
st.subheader("🔎 HTA Evidence Explorer")
st.markdown(f"*Structured retrieval and synthesis of comparable NICE appraisals — "
            f"{TOTAL_ROWS:,} decisions indexed*")

col_a, col_b = st.columns(2)
with col_a:
    drug_name = st.text_input(
        "Drug Name (optional)", placeholder="e.g. Adagrasib",
        help="Leave blank to search for analogues by indication alone, without a named "
             "product in mind.",
        key="w_nice_drug_name")
    indication = st.text_input("Indication", placeholder="e.g. Advanced NSCLC",
                               key="w_nice_indication")
    st.session_state.setdefault("w_nice_icer_provided", True)
    icer_provided = st.checkbox(
        "I have an ICER estimate",
        help="Uncheck to search precedent first and see the ICER range achieved, "
             "without committing to a figure upfront.",
        key="w_nice_icer_provided")
    if icer_provided:
        st.session_state.setdefault("w_nice_icer_low", 50000)
        st.session_state.setdefault("w_nice_icer_high", 0)
        cost_col1, cost_col2 = st.columns(2)
        with cost_col1:
            estimated_cost_low = st.number_input(
                "Estimated ICER (£/QALY)", min_value=0, max_value=500000,
                step=5000, key="w_nice_icer_low")
        with cost_col2:
            estimated_cost_high = st.number_input(
                "Upper estimate (optional)", min_value=0, max_value=500000,
                step=5000, key="w_nice_icer_high",
                help="Leave at 0 for a single figure. Set above the lower estimate to "
                     "submit a range instead of one exact number.")
    else:
        estimated_cost_low = estimated_cost_high = None
        st.caption("No ICER entered — results will show comparable precedent and the "
                   "ICER range they achieved, without a risk-signal comparison.")
with col_b:
    st.session_state.setdefault("w_nice_eol", "No")
    end_of_life = st.radio("End of Life Indication?", ["Yes", "No", "Not specified"],
                           key="w_nice_eol")
    comparator = st.text_input("Main Comparator", placeholder="e.g. Docetaxel",
                               key="w_nice_comparator")
    appraisal_type = st.radio("Appraisal Type", ["STA", "MTA", "Not specified"],
                              key="w_nice_appraisal_type")
    keyword = st.text_input("Indication keyword for benchmarking",
                            placeholder="e.g. lung, breast, immunology, diabetes",
                            key="w_nice_keyword")
    st.caption("Abbreviations and phrases both work — 'NSCLC', 'Advanced NSCLC' and "
               "'non-small cell lung cancer' all resolve to the same retrieval set. "
               "Leave blank to search using the Indication field above instead.")

with st.expander("Advanced profile (improves similarity matching where tagged data is available)"):
    st.caption(
        f"Options are read from the Tag_Vocabulary sheet, so they always match the "
        f"tagged data exactly. {tagged_total} of {TOTAL_ROWS:,} appraisals carry these "
        f"tags — currently {', '.join(tagged_areas)}. Filling these in sharpens the "
        f"similarity score for those indications; elsewhere the tool falls back to "
        f"indication-keyword matching."
    )
    p1, p2 = st.columns(2)
    with p1:
        line_of_therapy_input = st.selectbox(
            "Line of therapy",
            vocab_options(VOCAB, "line_of_therapy",
                          ["Not specified", "First line", "Second line", "Third line+"]),
            key="w_nice_line_of_therapy")
        mechanism_input = st.selectbox(
            "Mechanism of action",
            vocab_options(VOCAB, "mechanism_of_action", ["Not specified"]),
            key="w_nice_mechanism")
        patient_population_size_input = st.selectbox(
            "Patient population size",
            vocab_options(VOCAB,"patient_population_size",["Not specified"]),
            key="w_nice_population_size")
    with p2:
        biomarker_input = st.selectbox(
            "Biomarker", vocab_options(VOCAB, "biomarker", ["Not specified"]),
            key="w_nice_biomarker")
        comparator_type_input = st.selectbox(
            "Comparator type",
            vocab_options(VOCAB, "comparator_type", ["Not specified"]),
            key="w_nice_comparator_type")

if "hta_query_ran" not in st.session_state:
    st.session_state.hta_query_ran = False
if "chat_signature" not in st.session_state:
    st.session_state.chat_signature = None
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

if st.button("Retrieve Comparable Appraisals", type="primary"):
    st.session_state.hta_query_ran = True
    # A genuinely new query resets the chat — its grounding data has changed.
    new_signature = (drug_name, indication, keyword, mechanism_input,
                      line_of_therapy_input, biomarker_input, comparator_type_input, patient_population_size_input)
    if new_signature != st.session_state.chat_signature:
        st.session_state.chat_history = []
        st.session_state.chat_signature = new_signature

if st.session_state.hta_query_ran:
    if not indication:
        st.warning("Please enter an indication.")
        st.stop()

    # 'Not specified' EOL status defaults to the standard threshold — flagged
    # explicitly rather than silently assuming end-of-life status either way.
    if end_of_life == "Yes":
        threshold = 50000
    else:
        threshold = 30000
    eol_unspecified = (end_of_life == "Not specified")

    # Single source of truth for how the submitted ICER (or its absence) is
    # displayed and compared. icer_provided (set by the checkbox above) gates
    # every verdict/threshold section below; estimated_cost stays a plain
    # number (the midpoint of a range, if one was given) so the existing
    # comparison math needs no further changes — only the sections that
    # render it to the user need to check icer_provided first.
    cost_is_range = bool(estimated_cost_high and estimated_cost_high > estimated_cost_low) \
        if icer_provided else False
    if icer_provided:
        estimated_cost = (
            (estimated_cost_low + estimated_cost_high) / 2 if cost_is_range
            else estimated_cost_low
        )
        cost_display = (f"£{estimated_cost_low:,}–£{estimated_cost_high:,}" if cost_is_range
                        else f"£{estimated_cost:,}")
    else:
        estimated_cost = 0
        cost_display = "Not provided"

    keyword_search, matched_synonym = resolve_keyword(keyword)
    used_indication_fallback = False
    if not keyword_search and indication:
        # The 'Indication' field is descriptive only and was never wired to
        # search — leaving 'keyword' blank meant benchmarking against the
        # ENTIRE database with no filter, which reads as "broken" to anyone
        # who assumed the Indication field itself was doing the searching.
        keyword_search, matched_synonym = resolve_keyword(indication)
        used_indication_fallback = True

    if keyword_search:
        similar = df[df["indication"].str.contains(
            keyword_search, case=False, na=False, regex=False)]
        if used_indication_fallback:
            st.caption(
                f"No benchmarking keyword entered — used the Indication field "
                f"('{indication.strip()}') instead, interpreted as '{keyword_search}'.")
        elif matched_synonym and matched_synonym != keyword_search:
            st.caption(f"Interpreted '{keyword.strip()}' as '{keyword_search}' for retrieval.")
    else:
        similar = df

    inferred_area = None
    if len(similar) > 0:
        area_mode = similar["therapeutic_area"].dropna()
        if len(area_mode) > 0:
            inferred_area = area_mode.mode().iloc[0]

    query_profile = {
        "_drug_name": drug_name,
        "therapeutic_area": inferred_area,
        "mechanism_of_action": mechanism_input,
        "line_of_therapy": line_of_therapy_input,
        "comparator_type": comparator_type_input,
        "biomarker": biomarker_input,
        "orphan_status": None,
        "appraisal_type": appraisal_type,
        "patient_population_size": patient_population_size_input
    }
    # Fields that represent genuinely optional profile depth. therapeutic_area
    # is auto-inferred (never a user choice) and appraisal_type is a required
    # radio that's always STA/MTA — including either here made this check
    # permanently True, so scoring silently activated even when the user
    # never touched the Advanced profile at all.
    OPTIONAL_PROFILE_FIELDS = (
        "mechanism_of_action", "line_of_therapy", "comparator_type",
        "biomarker", "orphan_status","patient_population_size"
    )
    query_has_tags = any(
        query_profile.get(k) and query_profile.get(k) != "Not specified"
        for k in OPTIONAL_PROFILE_FIELDS
    )

    if len(similar) > 0:
        dataset_max_year = int(similar["year_start"].max()) if similar["year_start"].notna().any() else None
        results = similar.apply(
            lambda r: calculate_similarity_score(query_profile, r, dataset_max_year), axis=1)
        similar = similar.copy()
        similar["_similarity_score"] = [r[0] for r in results]
        similar["_similarity_breakdown"] = [r[1] for r in results]
        similar["_similarity_max"] = [r[2] for r in results]
        similar["_same_drug"] = [r[3] for r in results]
        if similar["_similarity_score"].notna().any():
            similar = similar.sort_values("_similarity_score", ascending=False,
                                          na_position="last")

    total_similar = len(similar)
    recommended_count = int((similar["decision_simple"] == "Recommended").sum())
    optimised_count = int((similar["decision_simple"] == "Optimised").sum())
    rejected_count = int((similar["decision_simple"] == "Not Recommended").sum())
    managed_count = int((similar["decision_simple"] == "Managed Access").sum())
    terminated_count = int((similar["decision_simple"] == "Terminated").sum())
    approval_rate = ((recommended_count + optimised_count) / total_similar * 100
                     if total_similar else 0)
    termination_rate = (terminated_count / total_similar * 100) if total_similar else 0

    if total_similar == 0:
        st.warning(
            f"No appraisals matched '{keyword.strip()}'. Try a broader term — "
            f"'lung', 'breast', 'colorectal', 'myeloma' — or leave the keyword blank "
            f"to benchmark against the full database."
        )
        st.stop()

    st.markdown("---")
    assessment_label = drug_name if drug_name else f"analogue search — {indication}"
    st.markdown(f"### Assessment for {assessment_label}")

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("WTP Threshold", f"£{threshold:,}")
    r2.metric("Your ICER" if icer_provided else "ICER Status", cost_display)
    r3.metric("Similar Appraisals", total_similar)
    r4.metric("Recommendation Proportion (retrieved set)", f"{approval_rate:.0f}%")

    s1, s2, s3, s4, s5 = st.columns(5)
    s1.metric("Recommended", recommended_count)
    s2.metric("Optimised", optimised_count)
    s3.metric("Not Recommended", rejected_count)
    s4.metric("Managed Access", managed_count)
    s5.metric("Terminated", terminated_count)

    patterns = []
    scoring_active = query_has_tags and has_tag_coverage(similar)

    st.markdown("**Similar appraisals found:**")
    if not query_has_tags:
        st.caption(
            "Ranked by indication keyword match only — fill in the Advanced profile "
            "above to enable weighted similarity scoring against tagged appraisals.")
    elif not has_tag_coverage(similar):
        st.caption(
            f"Advanced profile fields were entered, but none of the retrieved appraisals "
            f"in this indication carry structured tags. Weighted scoring covers "
            f"{tagged_total} tagged appraisals ({', '.join(tagged_areas)}). "
            f"Showing keyword-matched results instead.")
    else:
        st.caption(
            "Ranked by weighted similarity score where tags are available on both sides. "
            "Expand a row below to see which factors matched and what each was worth.")

    rename_map = {
        "drug_name": "Drug", "brand_name": "Brand", "decision_simple": "Decision",
        "indication": "Indication", "year_label": "Year", "appraisal_id": "Appraisal ID",
    }

    if scoring_active:
        disp = similar.head(10).copy()
        disp["Similarity"] = disp["_similarity_score"].apply(
            lambda x: f"{x:.0f}%" if pd.notna(x) else "—")
        # The percentage is computed over whatever COULD be scored, so a
        # sparsely-tagged appraisal can outrank a well-matched one on a much
        # smaller denominator. Showing the denominator keeps that visible
        # rather than letting a thin match masquerade as a strong one.
        disp["Scored on"] = disp["_similarity_max"].apply(
            lambda x: f"{x:.0f}/100 pts" if pd.notna(x) and x else "—")
        disp["Drug"] = disp.apply(
            lambda r: (r["drug_name"] + " 🔁") if r.get("_same_drug") else r["drug_name"], axis=1)
        st.dataframe(
            disp[["Drug", "brand_name", "decision_simple", "indication",
                  "year_label", "appraisal_id", "Similarity",
                  "Scored on"]].rename(columns=rename_map),
            width="stretch", hide_index=True)
        st.caption(
            "**Scored on** is how many of the 100 possible points were actually "
            "assessable for that appraisal. A high percentage on a low denominator "
            "means few factors were comparable — not a strong match. Compare "
            "appraisals with similar denominators.")
        if disp["_same_drug"].any():
            st.caption("🔁 = same drug name as your query, appraised previously under a "
                       "different TA number.")

        st.markdown("**Similarity breakdown by appraisal:**")
        st.caption(
            "Points earned / points possible per factor. Factors your query didn't "
            "specify are omitted entirely — not counted as a miss.")
        for _, row in similar[similar["_similarity_score"].notna()].head(5).iterrows():
            tag = " 🔁 same drug, prior appraisal" if row.get("_same_drug") else ""
            with st.expander(
                f"{row['drug_name']} ({row['appraisal_id']}) — "
                f"{row['_similarity_score']:.0f}% overall similarity "
                f"(scored on {row['_similarity_max']:.0f}/100 possible points){tag}"
            ):
                for f in row["_similarity_breakdown"]:
                    if f["status"] == "not_available":
                        st.markdown(f"⬜ **{f['label']}** — not tagged for this appraisal (excluded)")
                    else:
                        mark = ("✅" if f["points"] >= f["weight"] * 0.99
                                else "🟡" if f["points"] > 0 else "❌")
                        st.markdown(f"{mark} **{f['label']}**: {f['points']:g}/{f['weight']}")
                if pd.notna(row.get("tag_basis")):
                    st.caption(f"Tagging basis: {row['tag_basis']}")
    else:
        disp = similar.head(10).copy()
        # _same_drug is computed for every row regardless of scoring mode —
        # only the scored branch was checking it. Without this, querying an
        # already-approved drug silently includes its own real NICE decision
        # in the "similar precedent" table with no indication that's what
        # happened, which can make the recommendation-proportion figure look
        # more reassuring than the actual comparator evidence supports.
        disp["Drug"] = disp.apply(
            lambda r: (r["drug_name"] + " 🔁") if r.get("_same_drug") else r["drug_name"], axis=1)
        st.dataframe(
            disp[["Drug", "brand_name", "decision_simple", "indication",
                  "year_label", "appraisal_id"]].rename(columns=rename_map).head(10),
            width="stretch", hide_index=True)
        if disp["_same_drug"].any():
            st.caption("🔁 = same drug name as your query, appraised previously under a "
                       "different TA number — this is that drug's own precedent, not an "
                       "independent comparator.")

    # ── Era trend ──────────────────────────────────────────
    if total_similar >= 3:
        era_df = similar.dropna(subset=["year_start"]).copy()
        if len(era_df) >= 3:
            # Bin edges and labels now agree; the previous build labelled the
            # (1998, 2010] bin as "2005-2010" while it held rows from 2001.
            era_df["_era"] = pd.cut(
                era_df["year_start"],
                bins=[1999, 2010, 2016, 2022, 2030],
                labels=["2000–2010", "2011–2016", "2017–2022", "2023–2026"],
            )
            era_stats = (
                era_df.groupby("_era", observed=True)
                .agg(n=("decision_simple", "size"),
                     rate=("decision_simple",
                           lambda s: s.isin(["Recommended", "Optimised"]).sum() / len(s) * 100))
                .dropna()
            )
            if len(era_stats) > 0:
                st.markdown("**Recommendation rate by era (retrieved set):**")
                fig_era = px.bar(
                    x=era_stats.index.astype(str), y=era_stats["rate"],
                    labels={"x": "Era", "y": "Recommendation rate (%)"},
                    text=[f"{r:.0f}% (n={int(n)})"
                          for r, n in zip(era_stats["rate"], era_stats["n"])],
                    color_discrete_sequence=["#3498db"])
                fig_era.update_traces(textposition="outside")
                fig_era.update_layout(height=300, margin=dict(t=20, b=10), yaxis_range=[0, 105])
                st.plotly_chart(fig_era, width="stretch")
                st.caption(
                    "Recommendation rate = Recommended + Optimised as a share of retrieved "
                    "appraisals in that era; 'n' is the count behind each bar. Treat small-n "
                    "eras as indicative only — earlier eras may not reflect current NICE "
                    "methods, managed access frameworks, or NHS treatment pathways.")

    # ── Economic evidence ──────────────────────────────────
    st.markdown("**Economic evidence in this retrieved set**")

    with_icer = similar[similar["icer_lower"].notna()]
    has_note = similar["icer_evidence_note"].notna() & (
        similar["icer_evidence_note"].astype(str).str.strip().str.lower() != "not specified")
    note_only = similar[has_note & similar["icer_lower"].isna()]

    e1, e2, e3 = st.columns(3)
    e1.metric("Public numeric ICER", len(with_icer))
    e2.metric("Qualitative / confidential only", len(note_only))
    e3.metric("No economic detail extracted",
              total_similar - len(with_icer) - len(note_only))

    if len(with_icer) > 0:
        with_icer = with_icer.copy()

        with_icer["_icer_type"] = with_icer.apply(classify_historical_icer,axis=1)
        # Aggregate statistics must only use genuine point estimates
        point_icers = with_icer[with_icer["_icer_type"] == "point_estimate"]
        approved_point_icers = point_icers[
            point_icers["decision_simple"].isin(["Recommended", "Optimised"])]

        if len(approved_point_icers) > 0:
            i1, i2, i3 = st.columns(3)

            i1.metric(
                "Lowest point-estimate ICER",
                f"£{approved_point_icers['icer_lower'].min():,.0f}")
            i2.metric(
                "Highest point-estimate ICER",
                f"£{approved_point_icers['icer_lower'].max():,.0f}")

            avg = point_icers["icer_lower"].mean()

            if icer_provided:
                i3.metric(
                    "Your ICER vs point-estimate mean",
                    f"{'Above' if estimated_cost > avg else 'Below'} (£{avg:,.0f})")
            else:
                i3.metric(
                    "Mean point-estimate ICER",
                    f"£{avg:,.0f}")
        else:
            st.caption(
                "No confirmed point-estimate ICERs are available for aggregate benchmarking "
                "in this retrieved set."
            )

        icer_display = with_icer.copy()

        icer_display["ICER"] = icer_display.apply(
            lambda r: format_historical_icer(r, include_type=False),
            axis=1)

        icer_display["ICER Type"] = icer_display.apply(
            lambda r: ICER_TYPE_LABELS[classify_historical_icer(r)],
            axis=1)

        icer_cols = [
            "appraisal_id",
            "drug_name",
            "indication",
            "ICER",
            "ICER Type",
            "decision_simple",
            "economic_source_document_type",]

        st.dataframe(
            icer_display[icer_cols].rename(columns={
                "appraisal_id": "TA ID",
                "drug_name": "Drug",
                "indication": "Indication",
                "decision_simple": "Decision",
                "economic_source_document_type": "Source Document",
            }),
            width="stretch",
            hide_index=True)
        st.caption(
            "Published figures are not consistently labelled as company base-case, "
            "EAG-corrected, or committee-preferred. Treat as indicative published values, "
            "not confirmed accepted ICERs.")

    if len(note_only) > 0:
        with st.expander(
            f"Appraisals with a qualitative economic position only ({len(note_only)})"
        ):
            st.caption(
                "Cost-effectiveness results withheld under confidential commercial "
                "arrangements. The qualitative position is recorded; no numeric ICER is "
                "imputed, and a stated NICE threshold is never treated as an observed ICER.")
            for _, row in note_only.head(10).iterrows():
                st.markdown(f"**{row['appraisal_id']} — {row['drug_name']}**")
                st.markdown(f"> {row['icer_evidence_note']}")

    # ── Rejection reasoning ────────────────────────────────
    patterns = []
    rejected_similar = similar[similar["decision_simple"] == "Not Recommended"]

    if len(rejected_similar) > 0:
        ranked_themes, theme_sample, theme_sources = synthesise_themes(rejected_similar, query_profile = query_profile, all_comparables=similar,indication=indication)
        patterns = [(label, count) for label, (emoji, count) in ranked_themes]
        if ranked_themes:
            st.markdown(f"**Common themes across {theme_sample} rejected comparable appraisals:**")
            st.dataframe(
                pd.DataFrame([{"Theme": f"{emoji} {label}",
                               "Frequency":  (f"{count}/{len(similar)}"
                        if label in (
                            "Treatment effect waning / durability uncertainty",
                            "Population mismatch with retrieved precedent",
                        )
                        else f"{count}/{theme_sample}"
                    )}
                              for label, (emoji, count) in ranked_themes]),
                width="stretch", hide_index=True)

            for label, (emoji, count) in ranked_themes:
                with st.expander(f"Where '{label}' was raised"):
                    for aid in theme_sources.get(label, []):
                        qr = similar[similar["appraisal_id"] == aid]
                        quote = None
                        if len(qr) > 0:
                            source_row = qr.iloc[0]

                            for col in [
                                "original_nice_comment",
                                "detailed_reasoning",
                                "rejection_reasoning",
                            ]:
                                value = source_row.get(col)

                                if pd.notna(value):
                                    value = str(value).strip()

                                    if value.lower() not in (
                                        "",
                                        "not applicable",
                                        "not specified",
                                        "nan",
                                    ):
                                        quote = value
                                        break

                            if quote and len(quote) > 220:
                                quote = quote[:220].rsplit(" ", 1)[0] + "…"
                        st.markdown(f"**{aid}**")
                        st.markdown(f"> {quote}" if quote else
                                    "_No verbatim committee text available — see full guidance._")
            st.caption(
                "Synthesised from committee reasoning across the rejected appraisals shown "
                "below. A theme count reflects how many of these specific appraisals raised "
                "that concern — not a general base rate for the indication.")

        has_detail_col = "primary_reason_category" in rejected_similar.columns
        with_reasoning = rejected_similar[rejected_similar["rejection_reasoning"].notna()]

        if len(with_reasoning) > 0:
            all_cards, concern_counter, comparison_size = build_concern_frequency(
                with_reasoning, has_detail_col)

            st.markdown("**Individual rejected appraisals:**")
            if comparison_size > 1:
                st.caption(
                    f"Each appraisal's concerns are compared against the other "
                    f"{comparison_size - 1} rejected appraisal(s) in this set, so you can "
                    f"see what's a shared pattern versus specific to that drug.")

            for row, card in all_cards[:5]:
                with st.expander(f"{row['drug_name']} - {row['indication']} ({row['year_label']})"):
                    st.markdown("**Committee conclusion**")
                    st.write(card["conclusion"])

                    if comparison_size > 1:
                        shared, unique = split_shared_unique(
                            card["concerns"], concern_counter, comparison_size)
                        if shared:
                            st.markdown("**Shared concerns** _(also raised in other rejected "
                                        "appraisals here)_")
                            for c, freq in shared:
                                st.markdown(f"- {c} — shared with {freq}/{comparison_size} appraisals")
                        if unique:
                            st.markdown("**Unique to this appraisal**")
                            for c in unique:
                                st.markdown(f"- {c}")
                        if not shared and not unique:
                            st.write("Specific concerns not itemised in source text — "
                                     "see full guidance below.")
                    else:
                        st.markdown("**Key evidence concerns**")
                        for c in card["concerns"]:
                            st.markdown(f"- {c}")

                    st.markdown("**Reported ICER**")
                    st.write(card["icer_line"])

                    with st.expander("Show full source text"):
                        st.write(card["raw"])

                    if pd.notna(row.get("url")):
                        st.markdown(f"[View NICE guidance — full committee discussion]({row['url']})")

    # ── Optimised set ──────────────────────────────────────
    optimised_similar = similar[similar["decision_simple"] == "Optimised"]
    if len(optimised_similar) > 0:
        has_restrictions = (
            "restriction_type" in optimised_similar.columns
            and optimised_similar["restriction_type"].notna().any()
        )
        st.markdown(f"**Optimised appraisals in this retrieved set ({len(optimised_similar)}):**")

        if has_restrictions:
            tagged_restr = optimised_similar[optimised_similar["restriction_type"].notna()]
            untagged_restr = len(optimised_similar) - len(tagged_restr)
            st.caption(
                "Recommended only within a restricted population or under specific "
                "conditions. Restriction type is extracted from committee guidance where "
                "available." + (
                    f" {untagged_restr} appraisal(s) below aren't yet tagged — review those "
                    f"directly in NICE guidance." if untagged_restr else ""
                )
            )
            for _, row in optimised_similar.head(10).iterrows():
                rtype = row.get("restriction_type")
                label = (
                    f"{row['drug_name']} — {rtype}"
                    if pd.notna(rtype) and rtype not in ("Not applicable", "Not specified")
                    else f"{row['drug_name']} — restriction not yet extracted"
                )
                with st.expander(f"{label} ({row['appraisal_id']})"):
                    st.markdown(f"**Indication:** {row['indication']}")
                    note = row.get("restriction_note")
                    if pd.notna(note) and str(note).strip().lower() not in (
                        "not applicable", "not specified"):
                        st.write(note)
                    else:
                        st.caption("No structured restriction detail extracted for this "
                                   "appraisal — see full guidance below.")
                    if pd.notna(row.get("url")):
                        st.markdown(f"[View NICE guidance]({row['url']})")
        else:
            # No tagged rows in this retrieved set at all — same fallback as before.
            st.caption(
                "Recommended only within a restricted population or under specific "
                "conditions. Structured restriction-type data is not yet extracted for "
                "this indication — review the specific restrictions directly in NICE "
                "guidance.")
            st.dataframe(
                optimised_similar[["drug_name", "indication", "year_label", "appraisal_id", "url"]]
                .head(10).rename(columns={
                    "drug_name": "Drug", "indication": "Indication", "year_label": "Year",
                    "appraisal_id": "Appraisal ID", "url": "NICE Link"}),
                width="stretch", hide_index=True)

    # ── Evidence gaps ──────────────────────────────────────
    nonroutine = similar[similar["decision_simple"].isin(
        ["Not Recommended", "Terminated", "Managed Access"])]
    if len(nonroutine) >= 3:
        gap_themes, gap_sample, _ = synthesise_themes(
            nonroutine, query_profile=query_profile, all_comparables=similar,
            indication=indication, max_examples=15)
        if gap_themes:
            st.markdown("**Evidence gaps suggested by historical precedent:**")
            st.caption(
                f"Based on {gap_sample} non-routine appraisals in this set. This does not "
                f"mean NICE will raise the same issues for this submission — it indicates "
                f"areas that have historically required careful justification.")
            for label, (emoji, count) in gap_themes:
                st.markdown(f"☑ {label}")

    # ── Evidence completeness ──────────────────────────────
    st.markdown("### Evidence Completeness")
    st.caption("How much of this assessment rests on solid data versus a thin or "
               "keyword-only match — read this before the sections below.")

    similarity_conf = 9 if scoring_active else (4 if query_has_tags else 2)
    icer_coverage = int(similar["icer_lower"].notna().sum())
    icer_conf = round(min(icer_coverage / max(total_similar, 1), 1.0) * 10)
    detail_coverage = int(similar["detailed_reasoning"].notna().sum())
    reasoning_conf = (round(min(detail_coverage / max(rejected_count, 1), 1.0) * 10)
                      if rejected_count > 0 else 5)
    sample_conf = round(min(total_similar / 10, 1.0) * 10)

    def _bar(n):
        return "█" * n + "░" * (10 - n)

    avg_conf = (similarity_conf + sample_conf + icer_conf + reasoning_conf) / 4
    confidence_label = "High" if avg_conf >= 7 else "Moderate" if avg_conf >= 4 else "Low"

    st.markdown(f"**Similarity match quality** `{_bar(similarity_conf)}` {similarity_conf}/10")
    st.markdown(f"**Clinical precedent (sample size)** `{_bar(sample_conf)}` {sample_conf}/10 "
                f"— {total_similar} appraisals retrieved")
    st.markdown(f"**Economic precedent (published ICER)** `{_bar(icer_conf)}` {icer_conf}/10 "
                f"— {icer_coverage}/{total_similar} with a reported ICER")
    if rejected_count > 0:
        st.markdown(f"**Committee reasoning detail** `{_bar(reasoning_conf)}` "
                    f"{reasoning_conf}/10 — {detail_coverage}/{rejected_count} rejections "
                    f"with structured detail")
    else:
        st.markdown("**Committee reasoning detail** — no rejected appraisals in this set")
    st.markdown(f"**Overall confidence: {confidence_label}**")

    # ── Evidence summary ───────────────────────────────────
    st.markdown("### Evidence Summary")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"""
**Retrieved appraisal set (by indication keyword match):**
- {total_similar} appraisals identified
- {recommended_count} recommended
- {optimised_count} optimised
- {rejected_count} not recommended
- {managed_count} managed access
- {terminated_count} terminated
- Recommendation proportion within this set: {approval_rate:.0f}%
        """)
        st.caption("Descriptive only — the proportion of retrieved appraisals that were "
                   "recommended, not a predicted probability for this drug.")
    with c2:
        icer_lines = (
            f"- Submitted ICER: {cost_display}/QALY\n"
            f"- WTP reference threshold: £{threshold:,}/QALY\n"
            f"- Position vs threshold: {((estimated_cost / threshold) - 1) * 100:+.0f}%\n"
            if icer_provided else
            f"- ICER: not provided — analogue/precedent search only\n"
            f"- WTP reference threshold (for reference): £{threshold:,}/QALY\n"
        )
        st.markdown(f"""
**Your submitted profile (hypothetical):**
{icer_lines}- Comparator: {comparator or 'Not specified'}
- Appraisal type: {appraisal_type}
        """)
        st.caption("These are the figures you entered, not historical or verified NICE values.")

    # ── Contextual considerations ──────────────────────────
    # Built completely BEFORE rendering. The previous build appended the
    # high-termination warning after the render loop, so it never appeared
    # in the app but was still passed into the PDF — the report carried a
    # risk flag the dashboard didn't show.
    st.markdown("**Contextual considerations:**")
    warnings_list = []
    context_facts = []

    yrs = similar["year_start"].dropna()
    if len(yrs) > 0:
        oldest, newest = int(yrs.min()), int(yrs.max())
        span = newest - oldest
        if span >= 10:
            warnings_list.append(
                f"Retrieved appraisals span {oldest} to {newest} ({span} years). NICE methods, "
                f"treatment pathways, comparator prices, and clinical practice have likely "
                f"changed materially over that period.")
        else:
            context_facts.append(f"Retrieved appraisals span {oldest} to {newest} ({span} years).")

    GENERIC_DRUGS = ["omeprazole", "lansoprazole", "metformin", "atorvastatin", "amlodipine",
                     "ramipril", "lisinopril", "simvastatin", "docetaxel", "paclitaxel",
                     "carboplatin", "cisplatin", "capecitabine", "oxaliplatin", "gemcitabine"]
    if comparator and any(g in comparator.lower() for g in GENERIC_DRUGS):
        warnings_list.append(
            f"{comparator} is now a low-cost generic. Historical ICERs using it as a "
            f"comparator may understate the true incremental cost burden versus current "
            f"NHS pricing.")

    if total_similar < 5:
        warnings_list.append(
            f"Small retrieval set — only {total_similar} similar appraisal(s) found. Treat "
            f"any pattern drawn from this set with caution.")
    else:
        context_facts.append(f"Retrieval set size: {total_similar} appraisals.")

    if not keyword_search:
        warnings_list.append(
            "No indication could be resolved from your search terms — benchmarking against "
            "the full database rather than a targeted indication match.")
    elif used_indication_fallback:
        context_facts.append(
            "No benchmarking keyword was entered — the Indication field was used for retrieval "
            "instead.")

    icer_pct = icer_coverage / total_similar * 100
    if icer_pct < 20:
        warnings_list.append(
            f"Only {icer_coverage} of {total_similar} retrieved appraisals ({icer_pct:.0f}%) "
            f"have a publicly reported ICER — most modern appraisals keep this commercially "
            f"confidential, so ICER-based benchmarking here is necessarily thin.")
    else:
        context_facts.append(
            f"Published ICER available for {icer_coverage} of {total_similar} retrieved "
            f"appraisals ({icer_pct:.0f}%).")

    if total_similar >= 3:
        area_counts = similar["therapeutic_area"].dropna().value_counts()
        if len(area_counts) > 0:
            top_pct = area_counts.iloc[0] / total_similar * 100
            if top_pct >= 60:
                context_facts.append(
                    f"{top_pct:.0f}% of retrieved appraisals are in {area_counts.index[0]} — "
                    f"precedent is concentrated in this area.")

    if termination_rate > 50:
        warnings_list.append(
            f"High termination rate: {terminated_count} of {total_similar} similar appraisals "
            f"({termination_rate:.0f}%) were terminated without a submitted evidence package. "
            f"This often signals manufacturers were unable to agree a commercially viable "
            f"price with NICE — a submitted ICER below threshold does not on its own overcome "
            f"that pattern.")

    for w in warnings_list:
        st.warning(w)
    if context_facts:
        st.markdown("_Additional context:_")
        for c in context_facts:
            st.markdown(f"- {c}")
    if not warnings_list and not context_facts:
        st.info("No major contextual concerns identified.")

    # ── Risk signal ────────────────────────────────────────
    if termination_rate == 100 and total_similar >= 2:
        signal = "high_commercial_risk"
    elif termination_rate > 75 and total_similar >= 3:
        signal = "high_commercial_risk"
    elif not icer_provided:
        signal = "no_icer"
    elif estimated_cost <= threshold:
        signal = "low"
    elif estimated_cost <= threshold * 1.5:
        signal = "moderate"
    else:
        signal = "high"

    verdict = {
        "high_commercial_risk": "High Commercial Risk",
        "no_icer": "Precedent Reference Only",
        "low": "Likely Recommended",
        "moderate": "Borderline",
        "high": "Unlikely to be Recommended",
    }[signal]

    st.markdown("**Historical precedent review**")
    st.caption(
        "A descriptive signal based on your submitted ICER versus the reference threshold and "
        "retrieved precedent — not a prediction of the committee's decision. A recommendation "
        "cannot be inferred without the full evidence package, model structure, "
        "committee-preferred assumptions, and any confidential commercial arrangement.")

    if signal == "high_commercial_risk":
        st.error(f"""
Position: High commercial/pricing risk pattern in historical precedent

{termination_rate:.0f}% of retrieved appraisals in this indication were terminated
without a submitted evidence package, regardless of where an ICER might land.
This pattern is more often associated with pricing/commercial disagreement than
with the cost-effectiveness case itself.

Possible next steps:
- Investigate Highly Specialised Technologies pathway eligibility, if applicable
- Seek early NICE scientific advice before a formal submission
- Model list price vs net price scenarios explicitly
- Consider a patient access scheme or managed access route
- Assess commercial viability of UK launch independent of the ICER position
        """)
    elif signal == "no_icer":
        st.info(f"""
Position: No ICER submitted — this is a precedent reference, not a risk signal.

{approval_rate:.0f}% of the {total_similar} retrieved appraisals were recommended or
optimised. Review the ICER range under "Economic evidence" above for a sense of what
comparable submissions achieved — enter an estimate above to get a threshold-based
risk signal instead.

Possible next steps:
- Use the reported ICER range from comparable precedent as a starting benchmark
- Identify which comparator(s) NICE accepted in this indication
- Review the restriction conditions on any Optimised appraisals in this set — these
  often shape what a viable submission actually looks like
        """)
    elif signal == "low":
        st.success(f"""
Position: Submitted ICER is at or below the {'end-of-life' if end_of_life == 'Yes' else 'standard'} reference threshold of £{threshold:,}/QALY.

Possible next steps:
- Stress-test the clinical evidence base versus {comparator or 'the stated comparator'}
- {'Confirm end-of-life criteria are met and evidenced explicitly' if end_of_life == 'Yes' else 'Consider whether CDF/managed access is a fallback if evidence is still maturing'}
- Anticipate that a confidential commercial arrangement is often expected even below threshold
- Note: {optimised_count} appraisal(s) here were approved only with conditions — review what those were
        """)
    elif signal == "moderate":
        st.warning(f"""
Position: Submitted ICER exceeds the reference threshold by {((estimated_cost / threshold) - 1) * 100:.0f}%.

Possible next steps:
- Review the principal drivers of incremental cost and QALY gain in the model
- Test alternative, evidence-supported assumptions (survival extrapolation, utilities, retreatment)
- Model price or commercial-arrangement scenarios that would bring the ICER within range
- Assess whether {comparator or 'the stated comparator'} reflects current NHS practice
- Explore Cancer Drugs Fund / managed access as a contingency route
        """)
    else:
        st.error(f"""
Position: Submitted ICER exceeds the reference threshold by {((estimated_cost / threshold) - 1) * 100:.0f}%.

Possible next steps:
- Review the principal drivers of incremental cost and QALY gain — is the model biased toward a favourable case?
- Re-examine whether the clinical evidence is mature enough to support the QALY estimate
- Test alternative retreatment, extrapolation, and utility assumptions for sensitivity
- Assess whether the comparator reflects current NHS practice and pricing
- Seek NICE scientific advice before a formal submission
        """)

    st.caption(
        "This assessment is a preliminary, evidence-retrieval-based signal only. It does not "
        "constitute a prediction of a NICE committee decision and should not replace full "
        "economic modelling, evidence review, or professional market access advice.")

    st.markdown("---")
    research_questions = research_questions_from_themes(patterns)
    pdf_buffer = generate_assessment_pdf(
        drug_name, indication, estimated_cost, end_of_life, comparator, threshold,
        appraisal_type, total_similar, recommended_count, optimised_count,
        rejected_count, managed_count, terminated_count, approval_rate,
        similar, patterns, warnings_list, verdict, icer_provided, cost_display, research_questions
    )
    pdf_filename_base = drug_name if drug_name else re.sub(r"\s+", "_", indication.strip())[:40]
    st.download_button(
        "📥 Download PDF Report",
        data=pdf_buffer,
        file_name=f"{pdf_filename_base.replace(' ', '_')}_market_access_report.pdf",
        mime="application/pdf",
        type="primary",
    )

    # ── Chat over this retrieved set ─────────────────────
    st.markdown("---")
    st.markdown("### 💬 Ask about this precedent set")

    try:
        api_key = st.secrets.get("ANTHROPIC_API_KEY")
    except Exception:
        # Raised when no secrets.toml exists at all, not just when the key
        # is absent from it — this is the expected state before secrets
        # are configured, so it must degrade gracefully, not crash.
        api_key = None

    if not api_key:
        st.info(
            "Chat isn't configured yet — needs an ANTHROPIC_API_KEY in this app's Streamlit "
            "secrets."
        )
    else:
        st.caption(
            "Ask questions about the appraisals retrieved above — e.g. \"what's driving the "
            "rejections here?\" Answers are grounded only in this specific retrieved set and "
            "will say so if something isn't covered by it, rather than guessing."
        )

        turns_used = len(st.session_state.chat_history) // 2
        st.caption(f"{turns_used}/{CHAT_MAX_TURNS} questions used for this query.")

        for msg in st.session_state.chat_history:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        if turns_used >= CHAT_MAX_TURNS:
            st.warning(
                "Question limit reached for this query. Run a new search above to reset it."
            )
        else:
            user_question = st.chat_input("Ask a question about these results...")
            if user_question:
                st.session_state.chat_history.append(
                    {"role": "user", "content": user_question})
                with st.chat_message("user"):
                    st.markdown(user_question)

                context = build_chat_context(
                    similar, drug_name, indication, icer_provided, cost_display,
                    threshold, comparator)
                with st.chat_message("assistant"):
                    with st.spinner("Checking the retrieved precedent..."):
                        answer, error = ask_chat(
                            api_key, context, st.session_state.chat_history[:-1], user_question)
                    if error:
                        st.error(error)
                        st.session_state.chat_history.pop()  # don't count a failed turn
                    else:
                        st.markdown(answer)
                        st.session_state.chat_history.append(
                            {"role": "assistant", "content": answer})

st.divider()
st.caption(
    f"Built with Python & Streamlit | {TOTAL_ROWS:,} appraisals sourced from NICE Technology "
    f"Appraisals | Preliminary intelligence tool — not a substitute for full economic "
    f"modelling or professional market access advice"
)
