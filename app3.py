"""
NICE Technology Appraisal Intelligence — dashboard.

Data contract: NICE_v12_clean.xlsx
  Sheet1          — appraisal rows (adds year_start, year_label, search_blob)
  Tag_Vocabulary  — single source of truth for every categorical dropdown
  Enrichment_Log  — provenance, surfaced in the Methodology expander
"""

import io
import re
from collections import Counter
from datetime import date

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

DATA_FILE = "NICE_v12_clean.xlsx"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(
    page_title="NICE Intelligence Dashboard",
    page_icon="💊",
    layout="wide",
)


# Data loading

@st.cache_data
def load_data():
    return pd.read_excel(DATA_FILE, sheet_name="Sheet1")


@st.cache_data
def load_vocabulary():
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
        vocab = pd.read_excel(DATA_FILE, sheet_name="Tag_Vocabulary")
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


df = load_data()
VOCAB = load_vocabulary()
TOTAL_ROWS = len(df)


# Similarity scoring

SIMILARITY_WEIGHTS = {
    "therapeutic_area": ("Disease area", 30),
    "mechanism": ("Mechanism of action", 20),
    "line_of_therapy": ("Line of therapy", 15),
    "comparator_type": ("Comparator type", 15),
    "biomarker": ("Biomarker", 10),
    "appraisal_type": ("Appraisal type", 5),
    "orphan_status": ("Orphan status", 5),
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


# Reasoning synthesis

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
]

FALLBACK_CONCERN = "Specific concerns not itemised in source text — see full guidance"


def synthesise_themes(rejected_df, max_examples=8):
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
    if pd.notna(note) and str(note).strip().lower() not in ("", "not specified"):
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
    similar, patterns, warnings_list, verdict,
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
        else colors.HexColor("#e74c3c")
    )

    content = []

    header = Table(
        [[
            Paragraph(f"<b>{drug_name}</b>", ParagraphStyle(
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
        "Likely Recommended": "low risk",
        "Borderline": "moderate risk",
        "Unlikely to be Recommended": "high risk",
    }.get(verdict, verdict.lower())

    content.append(Paragraph("Executive Summary", heading))
    content.append(Paragraph(
        f"{drug_name} for {indication} has been benchmarked against {total_similar} NICE "
        f"technology appraisals retrieved by indication keyword match. The submitted ICER of "
        f"£{estimated_cost:,}/QALY sits {((estimated_cost / threshold) - 1) * 100:+.0f}% "
        f"relative to the £{threshold:,}/QALY reference threshold, producing an initial risk "
        f"signal of <b>{verdict_plain}</b>. Within the retrieved set, "
        f"{approval_rate:.0f}% of appraisals were recommended or optimised "
        f"({recommended_count} recommended, {optimised_count} optimised, "
        f"{rejected_count} not recommended) — this is descriptive of the retrieved precedent "
        f"only and is not a predicted probability of a NICE decision for this submission.",
        body))
    content.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#bdc3c7")))
    content.append(Spacer(1, 0.2 * cm))

    content.append(Paragraph("Economic Profile & Appraisal Landscape", heading))

    econ_data = [
        ["Parameter", "Value"],
        ["Estimated ICER", f"£{estimated_cost:,}/QALY"],
        ["WTP Threshold", f"£{threshold:,}/QALY"],
        ["Position vs Threshold", f"{((estimated_cost / threshold) - 1) * 100:+.0f}%"],
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
    sim_data = [["Drug", "Decision", "Indication", "Year", "TA ID"]]
    for _, row in similar.head(10).iterrows():
        sim_data.append([
            str(row["drug_name"])[:18],
            str(row["decision_simple"]),
            str(row["indication"])[:38],
            str(row.get("year_label", row.get("year", ""))),
            str(row["appraisal_id"]),
        ])
    sim_table = Table(sim_data, colWidths=[3.5 * cm, 3.2 * cm, 7 * cm, 2 * cm, 1.8 * cm])
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

    content.append(Paragraph("Strategic Recommendations", heading))
    if "Likely" in verdict:
        recs = [
            f"Ensure robust clinical evidence vs {comparator or 'comparator'}",
            "Prepare for commercial negotiation - confidential discount likely required",
            "Evidence end-of-life criteria clearly" if end_of_life == "Yes"
            else "Consider CDF route if evidence is immature",
            f"{optimised_count} similar drugs approved with conditions - prepare for optimisation",
            "Commission full probabilistic sensitivity analysis before submission",
        ]
    elif "Borderline" in verdict:
        recs = [
            "Explore Cancer Drugs Fund / managed access route as primary strategy",
            f"Strengthen clinical evidence package vs {comparator or 'comparator'}",
            f"Consider price reduction to bring ICER below £{threshold:,}/QALY",
            "Conduct full PSA to quantify uncertainty range",
            "Engage NICE scientific advice before formal submission",
        ]
    else:
        recs = [
            "Substantial price reduction required before submission",
            "Re-examine QALY estimates - are utility values robust?",
            "Consider alternative indication with stronger clinical evidence",
            "Engage NICE scientific advice before submission",
            "Review managed access options as interim route to market",
        ]

    rec_table = Table([[Paragraph(f"{i + 1}. {r}", body)] for i, r in enumerate(recs)],
                      colWidths=[18 * cm])
    rec_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#eaf4fb")),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#2980b9")),
        ("LINEBELOW", (0, 0), (-1, -2), 0.3, colors.HexColor("#d6eaf8")),
    ]))
    content.append(rec_table)
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


# Header

st.title("💊 NICE Technology Appraisal Intelligence")
st.markdown(f"*{TOTAL_ROWS:,} pharmaceutical appraisals — complete NICE database*")

tagged_total = int(df["line_of_therapy"].notna().sum())
with st.expander("Scope & coverage — read before benchmarking"):
    tagged_areas = (
        df[df["line_of_therapy"].notna()]["therapeutic_area"]
        .value_counts().head(5).index.tolist()
    )
    st.markdown(f"""
This tool operates at two levels of depth.

**Precedent browsing** works across all **{TOTAL_ROWS:,}** appraisals — search,
filter, decision history, rejection reasoning, and links to NICE guidance.

**Weighted similarity benchmarking** requires structured tags (line of therapy,
mechanism of action, biomarker, comparator type). These are currently populated
for **{tagged_total}** appraisals, concentrated in {', '.join(tagged_areas[:3])}
— specifically lung, breast, and colorectal cancer. Outside those, the tool falls
back to indication-keyword retrieval and says so.

**Published ICERs** exist for **{int(df['icer_lower'].notna().sum())}** appraisals.
Most modern NICE appraisals withhold cost-effectiveness results under confidential
commercial arrangements. Where that applies, the numeric field is deliberately left
empty and the qualitative position is recorded separately — a NICE threshold is
never encoded as if it were an observed ICER.
    """)
    try:
        log = pd.read_excel(DATA_FILE, sheet_name="Enrichment_Log")
        st.markdown("**Enrichment log**")
        # Cast to string: the log mixes counts and words ('Present') in one
        # column, which Arrow cannot serialise as a single type.
        st.dataframe(log.astype(str), width="stretch", hide_index=True)
    except Exception:
        pass

st.divider()


# Sidebar filters

st.sidebar.title("🔍 Filters")

search = st.sidebar.text_input(
    "Search",
    placeholder="Drug, brand, indication or TA ID",
    help="Searches generic name, brand name, indication and appraisal ID.",
)

therapeutic_areas = sorted(df["therapeutic_area"].dropna().astype(str).unique())
selected_areas = st.sidebar.multiselect(
    "Therapy / disease area", therapeutic_areas,
    help="Leave empty to include all areas.",
)

decisions = sorted(df["decision_simple"].dropna().astype(str).unique(), key=str.lower)
selected_decisions = st.sidebar.multiselect(
    "Decision", decisions, help="Leave empty to include all decisions."
)

year_min = int(df["year_start"].min())
year_max = int(df["year_start"].max())
year_range = st.sidebar.slider(
    "Appraisal year range", min_value=year_min, max_value=year_max,
    value=(year_min, year_max), step=1,
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
    for _, row in filtered_df[filtered_df["drug_name"] == selected_drug].iterrows():
        with st.expander(f"{row['appraisal_id']} - {row['indication']}"):
            c1, c2, c3 = st.columns(3)
            c1.metric("Decision", row["decision_simple"])
            c2.metric("Year", row["year_label"])
            c3.metric("Appraisal Type", row["appraisal_type"])
            brand = row.get("brand_name")
            if pd.notna(brand) and str(brand).strip().lower() != "not specified":
                st.caption(f"Brand name: {brand}")
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
    drug_name = st.text_input("Drug Name", placeholder="e.g. Adagrasib")
    indication = st.text_input("Indication", placeholder="e.g. Advanced NSCLC")
    estimated_cost = st.number_input("Estimated Cost (£/QALY)", min_value=0,
                                     max_value=500000, value=50000, step=5000)
with col_b:
    end_of_life = st.radio("End of Life Indication?", ["Yes", "No"], index=1)
    comparator = st.text_input("Main Comparator", placeholder="e.g. Docetaxel")
    appraisal_type = st.radio("Appraisal Type", ["STA", "MTA"])
    keyword = st.text_input("Indication keyword for benchmarking",
                            placeholder="e.g. lung, breast, colorectal")
    st.caption("Abbreviations and phrases both work — 'NSCLC', 'Advanced NSCLC' and "
               "'non-small cell lung cancer' all resolve to the same retrieval set.")

with st.expander("Advanced profile (improves similarity matching where tagged data is available)"):
    st.caption(
        f"Options are read from the Tag_Vocabulary sheet, so they always match the "
        f"tagged data exactly. {tagged_total} of {TOTAL_ROWS:,} appraisals carry these "
        f"tags — currently lung, breast and colorectal cancer. Filling these in sharpens "
        f"the similarity score for those indications; elsewhere the tool falls back to "
        f"indication-keyword matching."
    )
    p1, p2 = st.columns(2)
    with p1:
        line_of_therapy_input = st.selectbox(
            "Line of therapy",
            vocab_options(VOCAB, "line_of_therapy",
                          ["Not specified", "First line", "Second line", "Third line+"]))
        mechanism_input = st.selectbox(
            "Mechanism of action",
            vocab_options(VOCAB, "mechanism_of_action", ["Not specified"]))
    with p2:
        biomarker_input = st.selectbox(
            "Biomarker", vocab_options(VOCAB, "biomarker", ["Not specified"]))
        comparator_type_input = st.selectbox(
            "Comparator type",
            vocab_options(VOCAB, "comparator_type", ["Not specified"]))

if st.button("Retrieve Comparable Appraisals", type="primary"):
    if not (drug_name and indication):
        st.warning("Please enter a drug name and indication.")
        st.stop()

    threshold = 50000 if end_of_life == "Yes" else 30000
    keyword_search, matched_synonym = resolve_keyword(keyword)

    if keyword_search:
        similar = df[df["indication"].str.contains(
            keyword_search, case=False, na=False, regex=False)]
        if matched_synonym and matched_synonym != keyword_search:
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
    }
    query_has_tags = any(
        v and v != "Not specified"
        for k, v in query_profile.items()
        if k not in ("therapeutic_area", "_drug_name")
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
    st.markdown(f"### Assessment for {drug_name}")

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("WTP Threshold", f"£{threshold:,}")
    r2.metric("Your ICER", f"£{estimated_cost:,}")
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
            f"{tagged_total} tagged appraisals (lung, breast and colorectal cancer). "
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
        approved_icer = with_icer[
            with_icer["decision_simple"].isin(["Recommended", "Optimised"])]
        if len(approved_icer) > 0:
            if len(approved_icer) < 3:
                st.warning(
                    f"Only {len(approved_icer)} recommended/optimised appraisal(s) here have "
                    f"a public numeric ICER. Treat as an isolated historical data point, not "
                    f"a benchmark range.")
            i1, i2, i3 = st.columns(3)
            i1.metric("Lowest reported ICER", f"£{approved_icer['icer_lower'].min():,.0f}")
            i2.metric("Highest reported ICER", f"£{approved_icer['icer_lower'].max():,.0f}")
            avg = with_icer["icer_lower"].mean()
            i3.metric("Your ICER vs mean",
                      f"{'Above' if estimated_cost > avg else 'Below'} (£{avg:,.0f})")

        icer_cols = ["appraisal_id", "drug_name", "indication", "icer_lower", "icer_upper",
                     "decision_simple", "economic_source_document_type"]
        st.dataframe(
            with_icer[icer_cols].rename(columns={
                "appraisal_id": "TA ID", "drug_name": "Drug", "indication": "Indication",
                "icer_lower": "ICER Lower", "icer_upper": "ICER Upper",
                "decision_simple": "Decision",
                "economic_source_document_type": "Source Document"}),
            width="stretch", hide_index=True)
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
    rejected_similar = similar[similar["decision_simple"] == "Not Recommended"]

    if len(rejected_similar) > 0:
        ranked_themes, theme_sample, theme_sources = synthesise_themes(rejected_similar)
        patterns = [(label, count) for label, (emoji, count) in ranked_themes]
        if ranked_themes:
            st.markdown(f"**Common themes across {theme_sample} rejected comparable appraisals:**")
            st.dataframe(
                pd.DataFrame([{"Theme": f"{emoji} {label}",
                               "Frequency": f"{count}/{theme_sample}"}
                              for label, (emoji, count) in ranked_themes]),
                width="stretch", hide_index=True)

            for label, (emoji, count) in ranked_themes:
                with st.expander(f"Where '{label}' was raised"):
                    for aid in theme_sources.get(label, []):
                        qr = rejected_similar[rejected_similar["appraisal_id"] == aid]
                        quote = None
                        if len(qr) > 0 and pd.notna(qr.iloc[0].get("original_nice_comment")):
                            quote = str(qr.iloc[0]["original_nice_comment"]).strip()
                            if len(quote) > 220:
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
        st.markdown(f"**Optimised appraisals in this retrieved set ({len(optimised_similar)}):**")
        st.caption(
            "Recommended only within a restricted population or under specific conditions. "
            "Structured restriction-type data is not yet extracted — review the specific "
            "restrictions directly in NICE guidance.")
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
        gap_themes, gap_sample, _ = synthesise_themes(nonroutine, max_examples=15)
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
        st.markdown(f"""
**Your submitted profile (hypothetical):**
- Submitted ICER: £{estimated_cost:,}/QALY
- WTP reference threshold: £{threshold:,}/QALY
- Position vs threshold: {((estimated_cost / threshold) - 1) * 100:+.0f}%
- Comparator: {comparator or 'Not specified'}
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

    if not keyword:
        warnings_list.append(
            "No indication keyword entered — benchmarking against the full database rather "
            "than a targeted indication match.")

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
    elif estimated_cost <= threshold:
        signal = "low"
    elif estimated_cost <= threshold * 1.5:
        signal = "moderate"
    else:
        signal = "high"

    verdict = {
        "high_commercial_risk": "High Commercial Risk",
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
    pdf_buffer = generate_assessment_pdf(
        drug_name, indication, estimated_cost, end_of_life, comparator, threshold,
        appraisal_type, total_similar, recommended_count, optimised_count,
        rejected_count, managed_count, terminated_count, approval_rate,
        similar, patterns, warnings_list, verdict,
    )
    st.download_button(
        "📥 Download PDF Report",
        data=pdf_buffer,
        file_name=f"{drug_name.replace(' ', '_')}_market_access_report.pdf",
        mime="application/pdf",
        type="primary",
    )

st.divider()
st.caption(
    f"Built with Python & Streamlit | {TOTAL_ROWS:,} appraisals sourced from NICE Technology "
    f"Appraisals | Preliminary intelligence tool — not a substitute for full economic "
    f"modelling or professional market access advice"
)
