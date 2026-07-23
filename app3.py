import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
from collections import Counter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, HRFlowable
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from datetime import date

st.set_page_config(
    page_title="NICE Intelligence Dashboard",
    page_icon="💊",
    layout="wide"
)

@st.cache_data
def load_data():
    df = pd.read_excel(r"NICE_v10_final.xlsx")
    return df

df = load_data()

SIMILARITY_WEIGHTS = {
    "therapeutic_area":  ("Disease area",       30),
    "mechanism":         ("Mechanism of action", 20),
    "line_of_therapy":   ("Line of therapy",     15),
    "comparator_type":   ("Comparator type",     15),
    "biomarker":         ("Biomarker",           10),
    "appraisal_type":    ("Appraisal type",       5),
    "orphan_status":     ("Orphan status",        5),
}
RECENCY_MAX_POINTS = 5

def calculate_similarity_score(query, candidate_row, dataset_max_year=None):
    """
    Weighted similarity score (0-100) between a hypothetical drug profile
    (query) and a candidate historical appraisal row, using the structured
    tag fields where both sides have them populated.

    Categorical tags alone are coarse — several genuinely different drugs
    can share every tagged attribute and legitimately tie. To break ties
    within a cluster without inventing false precision, a small recency
    component is added: appraisals closer to the most recent one in the
    dataset score marginally higher, since that reflects current NICE
    methods and treatment pathway rather than an arbitrary tiebreak.

    Returns (score, breakdown, max_possible, same_drug) where breakdown is
    an ordered list of dicts — one per factor considered — each with label,
    weight, status ("match"/"no_match"/"not_available"), and points earned.
    same_drug flags an exact drug-name match, surfaced separately in the UI
    rather than folded silently into the percentage.
    """
    def _val(x):
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return None
        s = str(x).strip()
        return s if s and s.lower() != "not specified" else None

    field_map = {
        "therapeutic_area": ("therapeutic_area", "therapeutic_area"),
        "mechanism":         ("mechanism_of_action", "mechanism_of_action"),
        "line_of_therapy":   ("line_of_therapy", "line_of_therapy"),
        "comparator_type":   ("comparator_type", "comparator_type"),
        "biomarker":         ("biomarker", "biomarker"),
        "appraisal_type":    ("appraisal_type", "appraisal_type"),
        "orphan_status":     ("orphan_status", "orphan_status"),
    }

    score = 0
    max_possible = 0
    breakdown = []

    for key, (label, weight) in SIMILARITY_WEIGHTS.items():
        q_field, c_field = field_map[key]
        q_val = _val(query.get(q_field))
        c_val = _val(candidate_row.get(c_field))

        if q_val is None:
            # Query didn't specify this factor at all — not counted, not shown as a miss
            continue

        if c_val is None:
            breakdown.append({"label": label, "weight": weight, "status": "not_available", "points": 0})
            continue

        max_possible += weight
        if q_val.lower() == c_val.lower():
            score += weight
            breakdown.append({"label": label, "weight": weight, "status": "match", "points": weight})
        else:
            breakdown.append({"label": label, "weight": weight, "status": "no_match", "points": 0})

    if max_possible == 0:
        return None, breakdown, 0, False

    # ── Recency tiebreaker ────────────────────────────────
    # Always applied when a year is available on the candidate side, so it
    # can genuinely separate otherwise-tied candidates within a coarse
    # categorical cluster. Shown in the breakdown like any other factor.
    candidate_year_str = str(candidate_row.get("year", ""))[:4]
    if dataset_max_year and candidate_year_str.isdigit():
        year_gap = max(dataset_max_year - int(candidate_year_str), 0)
        recency_points = max(RECENCY_MAX_POINTS - (year_gap * 0.3), 0)
        recency_points = round(recency_points, 1)
        max_possible += RECENCY_MAX_POINTS
        score += recency_points
        breakdown.append({
            "label": "Recency", "weight": RECENCY_MAX_POINTS,
            "status": "match" if recency_points >= RECENCY_MAX_POINTS * 0.7 else "no_match",
            "points": recency_points
        })

    same_drug = False
    query_drug = query.get("_drug_name")
    if query_drug:
        candidate_drug = str(candidate_row.get("drug_name", ""))
        same_drug = query_drug.strip().lower() in candidate_drug.strip().lower() or candidate_drug.strip().lower() in query_drug.strip().lower()

    pct = round(score / max_possible * 100, 1) if max_possible > 0 else None
    return pct, breakdown, max_possible, same_drug


    pct = round(score / max_possible * 100)
    return pct, matched, max_possible

def has_tag_coverage(similar_df):
    """Whether this retrieved set has any tagged rows worth scoring against."""
    if "mechanism_of_action" not in similar_df.columns:
        return False
    return similar_df["mechanism_of_action"].notna().sum() > 0


    patterns = []
    for text in similar_df["rejection_patterns"].dropna():
        if text:
            patterns.extend(text.split(" | "))
    return Counter(patterns).most_common(5)

# Keyword tags used to synthesise cross-appraisal themes from free-text reasoning.
# Each tag maps to an emoji/label pair used in the "common themes" summary.
THEME_TAGS = [
    ("cost effectiveness / value for money", "💰", "Cost-effectiveness exceeded acceptable NHS value"),
    ("appropriate use of nhs resources",      "💰", "Cost-effectiveness exceeded acceptable NHS value"),
    ("immature",                              "⚠️", "Immature survival / follow-up evidence"),
    ("insufficient evidence",                 "⚠️", "Insufficient clinical-effectiveness evidence"),
    ("no direct evidence",                    "⚠️", "Insufficient clinical-effectiveness evidence"),
    ("indirect comparison",                   "⚠️", "Indirect treatment comparison uncertainty"),
    ("indirect treatment comparison",         "⚠️", "Indirect treatment comparison uncertainty"),
    ("uncertain",                              "⚠️", "Evidential or modelling uncertainty"),
    ("utility",                                "⚠️", "Utility value uncertainty"),
    ("comparator",                             "⚠️", "Comparator did not reflect NHS practice"),
    ("subgroup",                               "⚠️", "Restricted to a narrower subgroup"),
    ("extrapolation",                          "⚠️", "Long-term extrapolation uncertainty"),
    ("survival",                               "⚠️", "Survival benefit not established"),
    ("treatment duration",                     "⚠️", "Treatment duration assumptions unsupported"),
    ("stopping rule",                          "⚠️", "Stopping rule assumptions unsupported"),
]


def synthesise_themes(rejected_df, max_examples=8):
    """
    Scan free-text reasoning across a set of rejected appraisals and return
    theme -> (count, total) so the UI can show 'X/Y appraisals raised this'.
    Also returns theme -> list of appraisal_ids that support it, so every
    synthesised claim can be traced back to its source appraisals.
    Falls back gracefully if detailed_reasoning isn't populated for a row.
    """
    has_detail = "detailed_reasoning" in rejected_df.columns
    rows = rejected_df.head(max_examples)
    texts_with_ids = []
    for _, row in rows.iterrows():
        text = row.get("detailed_reasoning") if has_detail and pd.notna(row.get("detailed_reasoning")) else row.get("rejection_reasoning")
        texts_with_ids.append((str(text).lower() if pd.notna(text) else "", row.get("appraisal_id", "?")))

    total = len(texts_with_ids)
    if total == 0:
        return [], 0, {}

    seen_labels = {}
    theme_sources = {}
    for label_key, emoji, label_text in THEME_TAGS:
        matching_ids = [aid for t, aid in texts_with_ids if label_key in t]
        count = len(matching_ids)
        if count > 0:
            if label_text not in seen_labels or count > seen_labels[label_text][1]:
                seen_labels[label_text] = (emoji, count)
                theme_sources[label_text] = matching_ids
            elif label_text in theme_sources:
                theme_sources[label_text] = list(dict.fromkeys(theme_sources[label_text] + matching_ids))

    ranked = sorted(seen_labels.items(), key=lambda x: x[1][1], reverse=True)
    return ranked, total, theme_sources

def structure_reasoning_card(row, has_detail_col):
    """
    Split a rejection reasoning row into scannable components:
    conclusion / key concerns / reported ICER, instead of one raw paragraph.
    Returns a dict for rendering.
    """
    conclusion = None
    concerns = []
    icer_line = None

    if has_detail_col and pd.notna(row.get("primary_reason_category")):
        conclusion = row["primary_reason_category"]

    raw = row.get("detailed_reasoning") if has_detail_col and pd.notna(row.get("detailed_reasoning")) else row.get("rejection_reasoning")
    raw = str(raw) if pd.notna(raw) else ""

    if not conclusion:
        if "appropriate use of nhs resources" in raw.lower() or "value for money" in raw.lower():
            conclusion = "Not recommended because the incremental health benefit did not justify the additional NHS cost"
        elif "insufficient" in raw.lower() or "no direct evidence" in raw.lower():
            conclusion = "Not recommended due to insufficient clinical-effectiveness evidence"
        else:
            conclusion = "Not recommended (see full committee guidance for stated reason)"

    if has_detail_col and pd.notna(row.get("secondary_factors")):
        concerns = [f.strip() for f in str(row["secondary_factors"]).split(";") if f.strip()]
    if not concerns:
        lower = raw.lower()
        if "immature" in lower:
            concerns.append("Immature clinical/survival evidence")
        if "uncertain" in lower:
            concerns.append("Evidential or modelling uncertainty")
        if "utility" in lower:
            concerns.append("Utility value uncertainty")
        if "comparator" in lower:
            concerns.append("Comparator concerns")
        if not concerns:
            concerns.append("Specific concerns not itemised in source text — see full guidance")

    if "no publishable numeric icer" in raw.lower() or "no publishable icer" in raw.lower():
        icer_line = "No publishable ICER available"
    elif "£" in raw:
        import re
        matches = re.findall(r"£[\d,]+(?:\s*per\s*QALY|/QALY)?", raw)
        icer_line = "; ".join(dict.fromkeys(matches[:3])) if matches else None

    return {
        "conclusion": conclusion,
        "concerns": concerns,
        "icer_line": icer_line or "Not reported in source text",
        "raw": raw,
    }

def build_concern_frequency(rejected_df, has_detail_col, max_rows=15):
    """
    Build a set of structured cards for up to max_rows rejected appraisals,
    then count how often each distinct concern phrase appears across the set.
    Returns (list_of_cards_with_row, concern_counter, sample_size).
    """
    cards = []
    for _, row in rejected_df.head(max_rows).iterrows():
        card = structure_reasoning_card(row, has_detail_col)
        cards.append((row, card))

    concern_counter = Counter()
    for _, card in cards:
        for c in set(card["concerns"]):  # count each distinct concern once per appraisal
            concern_counter[c] += 1

    return cards, concern_counter, len(cards)

def split_shared_unique(card_concerns, concern_counter, sample_size):
    """
    Given one appraisal's concern list and the frequency counter across the
    comparison set, split into (shared, unique) with a 'X/Y appraisals' tag
    for shared items. The generic "not itemised" fallback is excluded from
    this comparison — it reflects missing source detail, not a genuine
    shared clinical/economic concern, so counting it as "shared" would be
    misleading.
    """
    FALLBACK = "Specific concerns not itemised in source text — see full guidance"
    shared, unique = [], []
    for c in card_concerns:
        if c == FALLBACK:
            continue
        freq = concern_counter.get(c, 1)
        if freq > 1:
            shared.append((c, freq))
        else:
            unique.append(c)
    return shared, unique

def generate_assessment_pdf(
    drug_name, indication, estimated_cost, qalys,
    end_of_life, comparator, threshold,
    total_similar, recommended_count, optimised_count,
    rejected_count, managed_count, approval_rate,
    similar, patterns, warnings_list, verdict
):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=1.5*cm, leftMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)

    # ── Styles ───────────────────────────────────────────
    heading_style  = ParagraphStyle("heading",  fontSize=12, spaceAfter=4, spaceBefore=8,
                                    fontName="Helvetica-Bold",
                                    textColor=colors.HexColor("#2c3e50"))
    body_style     = ParagraphStyle("body",     fontSize=9,  spaceAfter=3,
                                    fontName="Helvetica", leading=12)
    small_style    = ParagraphStyle("small",    fontSize=8,  spaceAfter=3,
                                    fontName="Helvetica", textColor=colors.grey)
    title_style    = ParagraphStyle("title",    fontSize=18, spaceAfter=4,
                                    fontName="Helvetica-Bold", alignment=TA_LEFT)
    subtitle_style = ParagraphStyle("subtitle", fontSize=10, spaceAfter=2,
                                    fontName="Helvetica", textColor=colors.grey)
    verdict_color  = (colors.HexColor("#27ae60") if "Likely" in verdict
                      else colors.HexColor("#e67e22") if "Borderline" in verdict
                      else colors.HexColor("#e74c3c"))
    verdict_style  = ParagraphStyle("verdict", fontSize=11, spaceAfter=4,
                                    fontName="Helvetica-Bold", textColor=verdict_color)

    content = []

    # ── Header bar ───────────────────────────────────────
    header_data = [[
        Paragraph(f"<b>{drug_name}</b>", ParagraphStyle("h", fontSize=16,
                  fontName="Helvetica-Bold", textColor=colors.white)),
        Paragraph(f"Market Access Intelligence Report<br/>"
                  f"<font size=9>{indication}</font>", 
                  ParagraphStyle("hr", fontSize=11, fontName="Helvetica",
                  textColor=colors.white, alignment=TA_RIGHT))
    ]]
    header_table = Table(header_data, colWidths=[9*cm, 9*cm])
    header_table.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,-1), colors.HexColor("#2c3e50")),
        ("TOPPADDING",   (0,0), (-1,-1), 10),
        ("BOTTOMPADDING",(0,0), (-1,-1), 10),
        ("LEFTPADDING",  (0,0), (-1,-1), 10),
        ("RIGHTPADDING", (0,0), (-1,-1), 10),
    ]))
    content.append(header_table)
    content.append(Spacer(1, 0.2*cm))

    # Meta line
    meta = f"Generated: {date.today().strftime('%d %B %Y')}   |   Comparator: {comparator or 'Not specified'}   |   End of Life: {end_of_life}   |   CONFIDENTIAL"
    content.append(Paragraph(meta, small_style))
    content.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#bdc3c7")))
    content.append(Spacer(1, 0.2*cm))

    # ── Verdict banner ───────────────────────────────────
    verdict_text = (
        "RISK SIGNAL: LOW - submitted ICER within reference threshold"
        if "Likely" in verdict
        else "RISK SIGNAL: MODERATE - submitted ICER exceeds threshold"
        if "Borderline" in verdict
        else "RISK SIGNAL: HIGH COMMERCIAL/PRICING RISK PATTERN"
        if "Commercial" in verdict
        else "RISK SIGNAL: HIGH - submitted ICER substantially exceeds threshold"
    )
    verdict_data = [[Paragraph(verdict_text, ParagraphStyle("vb", fontSize=10,
                    fontName="Helvetica-Bold", textColor=colors.white))]]
    verdict_table = Table(verdict_data, colWidths=[18*cm])
    verdict_table.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,-1), verdict_color),
        ("TOPPADDING",   (0,0), (-1,-1), 6),
        ("BOTTOMPADDING",(0,0), (-1,-1), 6),
        ("LEFTPADDING",  (0,0), (-1,-1), 10),
    ]))
    content.append(verdict_table)
    content.append(Spacer(1, 0.3*cm))

    # ── Executive Summary ────────────────────────────────
    content.append(Paragraph("Executive Summary", heading_style))
    content.append(Paragraph(
        f"{drug_name} for {indication} has been benchmarked against {total_similar} NICE "
        f"technology appraisals retrieved by indication keyword match. The submitted ICER of "
        f"£{estimated_cost:,}/QALY sits {((estimated_cost/threshold)-1)*100:+.0f}% relative to the "
        f"£{threshold:,}/QALY reference threshold, producing an initial risk signal of "
        f"<b>{verdict_text.lower()}</b>. "
        f"Within the retrieved set, {approval_rate:.0f}% of appraisals were recommended or "
        f"optimised ({recommended_count} recommended, {optimised_count} optimised, "
        f"{rejected_count} not recommended) — this is descriptive of the retrieved precedent "
        f"only and is not a predicted probability of a NICE decision for this submission.",
        body_style
    ))
    content.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#bdc3c7")))
    content.append(Spacer(1, 0.2*cm))

    # ── Two column layout: Economic + Landscape ──────────
    content.append(Paragraph("Economic Profile & Appraisal Landscape", heading_style))

    econ_data = [
        ["Parameter",             "Value"],
        ["Estimated ICER",        f"£{estimated_cost:,}/QALY"],
        ["WTP Threshold",         f"£{threshold:,}/QALY"],
        ["Position vs Threshold", f"{((estimated_cost/threshold)-1)*100:+.0f}%"],
        ["QALYs Gained",          str(qalys)],
        ["Comparator",            comparator or "Not specified"],
        ["End of Life",           end_of_life],
        ["Appraisal Type",        "STA"],
    ]

    landscape_data = [
        ["Decision",        "Count", "%"],
        ["Recommended",     str(recommended_count), f"{recommended_count/total_similar*100:.0f}%" if total_similar > 0 else "N/A"],
        ["Optimised",       str(optimised_count),   f"{optimised_count/total_similar*100:.0f}%"   if total_similar > 0 else "N/A"],
        ["Not Recommended", str(rejected_count),    f"{rejected_count/total_similar*100:.0f}%"    if total_similar > 0 else "N/A"],
        ["Managed Access",  str(managed_count),     f"{managed_count/total_similar*100:.0f}%"     if total_similar > 0 else "N/A"],
        ["Total Similar",   str(total_similar),     "100%"],
        ["Recommendation proportion", f"{approval_rate:.0f}%", ""],
    ]

    ts = TableStyle([
        ("BACKGROUND",    (0,0), (-1,0),  colors.HexColor("#34495e")),
        ("TEXTCOLOR",     (0,0), (-1,0),  colors.white),
        ("FONTNAME",      (0,0), (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,-1), 8),
        ("ALIGN",         (1,0), (-1,-1), "CENTER"),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [colors.white, colors.HexColor("#f2f2f2")]),
        ("GRID",          (0,0), (-1,-1), 0.3, colors.grey),
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 5),
    ])

    econ_table = Table(econ_data, colWidths=[4.5*cm, 4*cm])
    econ_table.setStyle(ts)

    land_table = Table(landscape_data, colWidths=[4*cm, 2*cm, 2*cm])
    land_table.setStyle(ts)

    two_col = Table([[econ_table, land_table]], colWidths=[9*cm, 9*cm])
    two_col.setStyle(TableStyle([
        ("VALIGN",      (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING",(0,0), (-1,-1), 5),
    ]))
    content.append(two_col)
    content.append(Spacer(1, 0.3*cm))
    content.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#bdc3c7")))

    # ── Similar Appraisals ───────────────────────────────
    content.append(Paragraph("Similar NICE Appraisals (Top 10)", heading_style))
    similar_top = similar[["appraisal_id", "drug_name", "indication", "decision_simple", "year"]].head(10)
    sim_data = [["Drug", "Decision", "Indication", "Year", "TA ID"]]
    for _, row in similar_top.iterrows():
        sim_data.append([
            str(row["drug_name"])[:18],
            str(row["decision_simple"]),
            str(row["indication"])[:38],
            str(row["year"]),
            str(row["appraisal_id"])
        ])
    sim_table = Table(sim_data, colWidths=[3.5*cm, 3.2*cm, 7*cm, 2*cm, 1.8*cm])
    sim_table.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0),  colors.HexColor("#34495e")),
        ("TEXTCOLOR",     (0,0), (-1,0),  colors.white),
        ("FONTNAME",      (0,0), (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,-1), 7.5),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [colors.white, colors.HexColor("#f2f2f2")]),
        ("GRID",          (0,0), (-1,-1), 0.3, colors.grey),
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 4),
    ]))
    content.append(sim_table)
    content.append(Spacer(1, 0.3*cm))
    content.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#bdc3c7")))

    # ── Rejection Analysis + Contextual side by side ─────
    content.append(Paragraph("Rejection Risk Analysis & Contextual Considerations", heading_style))

    if patterns:
        pat_data = [["Rejection Theme", "Freq"]]
        for pattern, count in patterns:
            pat_data.append([pattern, f"{count}"])
        pat_table = Table(pat_data, colWidths=[6.5*cm, 1.5*cm])
        pat_table.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,0),  colors.HexColor("#c0392b")),
            ("TEXTCOLOR",     (0,0), (-1,0),  colors.white),
            ("FONTNAME",      (0,0), (-1,0),  "Helvetica-Bold"),
            ("FONTSIZE",      (0,0), (-1,-1), 8),
            ("ROWBACKGROUNDS",(0,1), (-1,-1), [colors.HexColor("#fdf2f2"), colors.white]),
            ("GRID",          (0,0), (-1,-1), 0.3, colors.grey),
            ("TOPPADDING",    (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
            ("LEFTPADDING",   (0,0), (-1,-1), 5),
        ]))
    else:
        pat_table = Paragraph("No common rejection patterns identified.", body_style)

    warn_content = []
    if warnings_list:
        for w in warnings_list:
            warn_content.append(Paragraph(f"- {w}", small_style))
    else:
        warn_content.append(Paragraph("No major contextual concerns identified.", small_style))

    warn_table = Table([[w] for w in warn_content], colWidths=[9*cm])
    warn_table.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,-1), colors.HexColor("#fef9e7")),
        ("TOPPADDING",   (0,0), (-1,-1), 3),
        ("BOTTOMPADDING",(0,0), (-1,-1), 3),
        ("LEFTPADDING",  (0,0), (-1,-1), 5),
        ("BOX",          (0,0), (-1,-1), 0.5, colors.HexColor("#f39c12")),
    ]))

    two_col2 = Table([[pat_table, warn_table]], colWidths=[9*cm, 9*cm])
    two_col2.setStyle(TableStyle([
        ("VALIGN",      (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING",(0,0), (-1,-1), 5),
    ]))
    content.append(two_col2)
    content.append(Spacer(1, 0.3*cm))
    content.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#bdc3c7")))

    # ── Strategic Recommendations ────────────────────────
    content.append(Paragraph("Strategic Recommendations", heading_style))

    if "Likely" in verdict:
        recs = [
            f"Ensure robust clinical evidence vs {comparator or 'comparator'}",
            "Prepare for commercial negotiation - confidential discount likely required",
            f"{'Evidence end-of-life criteria clearly' if end_of_life == 'Yes' else 'Consider CDF route if evidence is immature'}",
            f"{optimised_count} similar drugs approved with conditions - prepare for optimisation",
            "Commission full probabilistic sensitivity analysis before submission"
        ]
    elif "Borderline" in verdict:
        recs = [
            "Explore Cancer Drugs Fund / managed access route as primary strategy",
            f"Strengthen clinical evidence package vs {comparator or 'comparator'}",
            f"Consider price reduction to bring ICER below £{threshold:,}/QALY",
            "Conduct full PSA to quantify uncertainty range",
            "Engage NICE scientific advice before formal submission"
        ]
    else:
        recs = [
            "Substantial price reduction required before submission",
            "Re-examine QALY estimates - are utility values robust?",
            "Consider alternative indication with stronger clinical evidence",
            "Engage NICE scientific advice before submission",
            "Review managed access options as interim route to market"
        ]

    rec_data = [[Paragraph(f"{i+1}. {rec}", body_style)] for i, rec in enumerate(recs)]
    rec_table = Table(rec_data, colWidths=[18*cm])
    rec_table.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,-1), colors.HexColor("#eaf4fb")),
        ("TOPPADDING",   (0,0), (-1,-1), 3),
        ("BOTTOMPADDING",(0,0), (-1,-1), 3),
        ("LEFTPADDING",  (0,0), (-1,-1), 8),
        ("BOX",          (0,0), (-1,-1), 0.5, colors.HexColor("#2980b9")),
        ("LINEBELOW",    (0,0), (-1,-2), 0.3, colors.HexColor("#d6eaf8")),
    ]))
    content.append(rec_table)
    content.append(Spacer(1, 0.3*cm))
    content.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#bdc3c7")))
    content.append(Spacer(1, 0.1*cm))
    content.append(Paragraph(
        "Disclaimer: This report was generated automatically based on historical NICE appraisal data. "
        "It is intended as a preliminary intelligence tool only. Further economic modelling and expert "
        "review is strongly recommended before drawing conclusions or making submission decisions. "
        f"Generated by NICE Technology Appraisal Intelligence Tool | {date.today().strftime('%d %B %Y')}",
        small_style
    ))

    doc.build(content)
    buffer.seek(0)
    return buffer

# ── Title ─────────────────────────────────────────────────
st.title("💊 NICE Technology Appraisal Intelligence")
st.markdown("*1,435 pharmaceutical appraisals - complete NICE database*")
st.divider()

# ── Sidebar ───────────────────────────────────────────────
st.sidebar.title("🔍 Filters")
search = st.sidebar.text_input("Search Drug Name")
decisions_vals = df["decision_simple"].dropna().astype(str).unique().tolist()
decisions = ["All"] + sorted(decisions_vals, key=lambda x: x.lower())
selected_decision = st.sidebar.selectbox("Decision", decisions)
years = ["All"] + sorted(df["year"].dropna().unique().tolist(), reverse=True)
selected_year = st.sidebar.selectbox("Year", years)

filtered_df = df.copy()
if search:
    filtered_df = filtered_df[filtered_df["drug_name"].str.contains(search, case=False, na=False)]
if selected_decision != "All":
    filtered_df = filtered_df[filtered_df["decision_simple"] == selected_decision]
if selected_year != "All":
    filtered_df = filtered_df[filtered_df["year"] == selected_year]

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("Total Appraisals", len(filtered_df))
with col2:
    st.metric("Recommended", len(filtered_df[filtered_df["decision_simple"] == "Recommended"]))
with col3:
    st.metric("Optimised", len(filtered_df[filtered_df["decision_simple"] == "Optimised"]))
with col4:
    st.metric("Not Recommended", len(filtered_df[filtered_df["decision_simple"] == "Not Recommended"]))
with col5:
    st.metric("Managed Access", len(filtered_df[filtered_df["decision_simple"] == "Managed Access"]))

st.divider()

buffer = io.BytesIO()
filtered_df.to_excel(buffer, index=False)
st.download_button(
    label="📥 Download Filtered Results",
    data=buffer.getvalue(),
    file_name="nice_filtered.xlsx",
    mime="application/vnd.ms-excel"
)

st.caption(f"Showing {len(filtered_df)} of {len(df)} appraisals")
st.dataframe(
    filtered_df[["appraisal_id", "drug_name", "indication", "decision_simple", "year", "url"]].rename(columns={
        "appraisal_id":    "TA ID",
        "drug_name":       "Drug Name",
        "indication":      "Indication",
        "decision_simple": "Decision",
        "year":            "Year",
        "url":             "NICE Link"
    }),
    use_container_width=True
)

st.divider()
st.subheader("📋 Drug Detail")
drug_options = filtered_df["drug_name"].dropna().unique().tolist()
if drug_options:
    selected_drug = st.selectbox("Select a drug", sorted(drug_options))
    drug_rows = filtered_df[filtered_df["drug_name"] == selected_drug]
    for _, row in drug_rows.iterrows():
        with st.expander(f"{row['appraisal_id']} - {row['indication']}"):
            col_d1, col_d2, col_d3 = st.columns(3)
            with col_d1:
                st.metric("Decision", row["decision_simple"])
            with col_d2:
                st.metric("Year", row["year"])
            with col_d3:
                st.metric("Appraisal Type", row["appraisal_type"])
            if pd.notna(row.get("rejection_reasoning")):
                st.caption(str(row["rejection_reasoning"])[:300])
            st.markdown(f"[View NICE Guidance]({row['url']})")

st.divider()
st.subheader("📊 Analysis")

COLORS = {
    "Recommended":      "#2ecc71",
    "Not Recommended":  "#e74c3c",
    "Managed Access":   "#f39c12",
    "Optimised":        "#3498db",
    "Terminated":       "#95a5a6",
    "Only in Research": "#9b59b6"
}

col_left, col_right = st.columns(2)
with col_left:
    st.markdown("**Decision Breakdown**")
    fig1 = px.pie(filtered_df, names="decision_simple", color="decision_simple",
                  color_discrete_map=COLORS, hole=0.4)
    st.plotly_chart(fig1, use_container_width=True)

with col_right:
    st.markdown("**Approvals Over Time**")
    yearly = filtered_df[
        filtered_df["decision_simple"].isin(["Recommended", "Not Recommended", "Managed Access", "Optimised"])
    ].groupby(["year", "decision_simple"]).size().reset_index(name="count")
    fig2 = px.line(yearly, x="year", y="count", color="decision_simple",
                   color_discrete_map=COLORS, markers=True)
    fig2.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig2, use_container_width=True)

st.markdown("**Top 15 Indications**")
top_indications = filtered_df["indication"].value_counts().head(15).reset_index()
top_indications.columns = ["Indication", "Count"]
fig3 = px.bar(top_indications, x="Count", y="Indication", orientation="h",
              color_discrete_sequence=["#3498db"])
fig3.update_layout(yaxis={"categoryorder": "total ascending"})
st.plotly_chart(fig3, use_container_width=True)

col_c1, col_c2 = st.columns(2)
with col_c1:
    st.markdown("**STA vs MTA**")
    type_counts = filtered_df["appraisal_type"].value_counts().reset_index()
    type_counts.columns = ["Type", "Count"]
    fig4 = px.pie(type_counts, values="Count", names="Type", hole=0.4,
                  color_discrete_sequence=["#3498db", "#e74c3c"])
    st.plotly_chart(fig4, use_container_width=True)

with col_c2:
    st.markdown("**Decision by Appraisal Type**")
    type_decision = filtered_df.groupby(["appraisal_type", "decision_simple"]).size().reset_index(name="count")
    fig5 = px.bar(type_decision, x="appraisal_type", y="count", color="decision_simple",
                  color_discrete_map=COLORS, barmode="stack")
    st.plotly_chart(fig5, use_container_width=True)

st.divider()
st.subheader("🔎 HTA Evidence Explorer")
st.markdown("*Structured retrieval and synthesis of comparable NICE appraisals — 1,439 decisions indexed*")

col_a, col_b = st.columns(2)
with col_a:
    drug_name      = st.text_input("Drug Name", placeholder="e.g. Adagrasib")
    indication     = st.text_input("Indication", placeholder="e.g. Advanced NSCLC")
    estimated_cost = st.number_input("Estimated Cost (£/QALY)", min_value=0, max_value=500000, value=50000, step=5000)
    qalys          = st.number_input("Estimated QALYs Gained", min_value=0.0, max_value=5.0, value=0.5, step=0.1)

with col_b:
    end_of_life    = st.radio("End of Life Indication?", ["Yes", "No"])
    comparator     = st.text_input("Main Comparator", placeholder="e.g. Docetaxel")
    appraisal_type = st.radio("Appraisal Type", ["STA", "MTA"])
    keyword        = st.text_input("Indication keyword for benchmarking", placeholder="e.g. lung, breast, diabetes")
    st.caption("Tips: use 'colitis' for UC, 'crohn' for Crohn's, 'myeloma' for myeloma, 'lung' for NSCLC")

with st.expander("Advanced profile (improves similarity matching where tagged data is available)"):
    st.caption(
        "Currently only lung cancer appraisals (48 of 1,439) have this level of tagging. "
        "Filling these in sharpens the similarity score for lung indications; for other "
        "indications the tool will fall back to indication-keyword matching only."
    )
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        line_of_therapy_input = st.selectbox(
            "Line of therapy",
            ["Not specified", "First line", "Second line", "Third line+"]
        )
        mechanism_input = st.selectbox(
            "Mechanism of action",
            ["Not specified", "PD-1/PD-L1 inhibitor", "PD-1/CTLA-4 combination",
             "Tyrosine kinase inhibitor", "Small-molecule targeted inhibitor",
             "Targeted therapy combination", "Monoclonal antibody — targeted",
             "Antibody–drug conjugate", "Bispecific T-cell engager"]
        )
    with col_p2:
        biomarker_input = st.selectbox(
            "Biomarker",
            ["Not specified", "PD-L1 positive", "EGFR mutation-positive",
             "EGFR T790M mutation-positive", "EGFR exon 20 insertion-positive",
             "ALK-positive", "ROS1-positive", "RET fusion-positive",
             "KRAS G12C mutation-positive", "MET alteration-positive",
             "MET exon 14 skipping-positive", "BRAF V600 mutation-positive",
             "BRAF V600E mutation-positive", "HER2 mutation-positive"]
        )
        comparator_type_input = st.selectbox(
            "Comparator type",
            ["Not specified", "Chemotherapy", "Best supportive care",
             "Another targeted therapy", "Another immunotherapy",
             "Immunotherapy plus chemotherapy", "Standard care / surveillance",
             "Mixed active comparators"]
        )

if st.button("Retrieve Comparable Appraisals", type="primary"):
    if drug_name and indication:
        threshold = 50000 if end_of_life == "Yes" else 30000

        search_synonyms = {
            "crohns": "crohn", "crohn's": "crohn", "uc": "colitis",
            "ibd": "colitis", "nsclc": "lung", "sclc": "lung",
            "ra": "rheumatoid", "multiple myeloma": "myeloma",
            "cll": "lymphocytic", "cml": "leukaemia"
        }
        keyword_search = search_synonyms.get(keyword.lower(), keyword) if keyword else keyword
        similar = df[df["indication"].str.contains(keyword_search, case=False, na=False)] if keyword_search else df

        # ── Weighted similarity scoring ───────────────────
        # Infer the query's therapeutic area from the retrieved set itself
        # (every row already matched the indication keyword, so the modal
        # therapeutic_area of that set is the query's own area — this isn't
        # a separate user input, it's read off the retrieval already done).
        inferred_therapeutic_area = None
        if len(similar) > 0 and "therapeutic_area" in similar.columns:
            area_mode = similar["therapeutic_area"].dropna()
            if len(area_mode) > 0:
                inferred_therapeutic_area = area_mode.mode().iloc[0]

        query_profile = {
            "_drug_name":          drug_name,
            "therapeutic_area":    inferred_therapeutic_area,
            "mechanism_of_action": mechanism_input,
            "line_of_therapy":     line_of_therapy_input,
            "comparator_type":     comparator_type_input,
            "biomarker":           biomarker_input,
            "orphan_status":       None,   # not collected as a direct input; skipped gracefully
            "appraisal_type":      appraisal_type,
        }
        query_has_tags = any(
            v and v != "Not specified" for k, v in query_profile.items() if k not in ("therapeutic_area", "_drug_name")
        )

        if len(similar) > 0:
            dataset_max_year = None
            try:
                year_nums_all = pd.to_numeric(similar["year"].dropna().astype(str).str[:4], errors="coerce").dropna()
                if len(year_nums_all) > 0:
                    dataset_max_year = int(year_nums_all.max())
            except Exception:
                pass

            sim_results = similar.apply(
                lambda row: calculate_similarity_score(query_profile, row, dataset_max_year), axis=1
            )
            similar = similar.copy()
            similar["_similarity_score"]     = [r[0] for r in sim_results]
            similar["_similarity_breakdown"] = [r[1] for r in sim_results]
            similar["_same_drug"]            = [r[3] for r in sim_results]
            if similar["_similarity_score"].notna().any():
                similar = similar.sort_values("_similarity_score", ascending=False, na_position="last")

        total_similar     = len(similar)
        recommended_count = len(similar[similar["decision_simple"] == "Recommended"])
        optimised_count   = len(similar[similar["decision_simple"] == "Optimised"])
        rejected_count    = len(similar[similar["decision_simple"] == "Not Recommended"])
        managed_count     = len(similar[similar["decision_simple"] == "Managed Access"])
        terminated_count  = len(similar[similar["decision_simple"] == "Terminated"])
        approval_rate     = (recommended_count + optimised_count) / total_similar * 100 if total_similar > 0 else 0

        st.markdown("---")
        st.markdown(f"### Assessment for {drug_name}")

        col_r1, col_r2, col_r3, col_r4 = st.columns(4)
        with col_r1:
            st.metric("WTP Threshold",            f"£{threshold:,}")
        with col_r2:
            st.metric("Your ICER",                f"£{estimated_cost:,}")
        with col_r3:
            st.metric("Similar Appraisals",       total_similar)
        with col_r4:
            st.metric("Recommendation Proportion (retrieved set)", f"{approval_rate:.0f}%")

        col_s1, col_s2, col_s3, col_s4, col_s5 = st.columns(5)
        with col_s1:
            st.metric("Recommended",     recommended_count)
        with col_s2:
            st.metric("Optimised",       optimised_count)
        with col_s3:
            st.metric("Not Recommended", rejected_count)
        with col_s4:
            st.metric("Managed Access",  managed_count)
        with col_s5:
            st.metric("Terminated",     terminated_count)

        patterns = []
        if total_similar > 0 and keyword:
            st.markdown("**Similar appraisals found:**")

            if not query_has_tags:
                st.caption(
                    "Ranked by indication keyword match only — fill in the Advanced profile "
                    "above (line of therapy, mechanism, comparator type, biomarker) to enable "
                    "weighted similarity scoring against tagged appraisals."
                )
            elif not has_tag_coverage(similar):
                st.caption(
                    "Advanced profile fields were entered, but none of the retrieved appraisals "
                    "in this indication currently have structured tags — weighted similarity "
                    "scoring is only available for the tagged lung cancer subset (48 appraisals). "
                    "Showing keyword-matched results instead."
                )
            else:
                st.caption(
                    "Ranked by weighted similarity score where tags are available on both sides. "
                    "Expand any row below the table to see exactly which factors did and didn't "
                    "match, and how much each was worth."
                )

            display_cols = ["drug_name", "decision_simple", "indication", "year", "appraisal_id"]
            rename_map = {
                "drug_name": "Drug", "decision_simple": "Decision",
                "indication": "Indication", "year": "Year", "appraisal_id": "Appraisal ID"
            }
            scoring_active = query_has_tags and has_tag_coverage(similar)

            if scoring_active:
                display_df = similar.head(10).copy()
                display_df["Similarity"] = display_df["_similarity_score"].apply(
                    lambda x: f"{x:.0f}%" if pd.notna(x) else "—"
                )
                display_df["Drug"] = display_df.apply(
                    lambda r: (r["drug_name"] + " 🔁") if r.get("_same_drug") else r["drug_name"], axis=1
                )
                st.dataframe(
                    display_df[["Drug", "decision_simple", "indication", "year", "appraisal_id", "Similarity"]].rename(columns=rename_map),
                    use_container_width=True
                )
                if display_df["_same_drug"].any():
                    st.caption("🔁 = same drug name as your query, appraised previously under a different TA number.")

                # ── Detailed, explainable breakdown per row ──
                st.markdown("**Similarity breakdown by appraisal:**")
                st.caption(
                    "Each factor shown as points earned / points possible for that factor in "
                    "this specific comparison. Factors your query didn't specify are omitted "
                    "entirely — not counted as a miss."
                )
                scored_rows = similar[similar["_similarity_score"].notna()].head(5)
                for _, row in scored_rows.iterrows():
                    breakdown = row["_similarity_breakdown"]
                    score_val = row["_similarity_score"]
                    same_drug_tag = " 🔁 same drug, prior appraisal" if row.get("_same_drug") else ""
                    with st.expander(f"{row['drug_name']} ({row['appraisal_id']}) — {score_val:.0f}% overall similarity{same_drug_tag}"):
                        for factor in breakdown:
                            if factor["status"] == "not_available":
                                st.markdown(f"⬜ **{factor['label']}** — not tagged for this appraisal (excluded)")
                            else:
                                marker = "✅" if factor["points"] >= factor["weight"] * 0.99 else ("🟡" if factor["points"] > 0 else "❌")
                                st.markdown(f"{marker} **{factor['label']}**: {factor['points']:g}/{factor['weight']}")

            else:
                st.dataframe(
                    similar[display_cols].rename(columns=rename_map).head(10),
                    use_container_width=True
                )

            # ── Time trend of the retrieved comparable set ────
            if total_similar >= 3:
                try:
                    era_df = similar.copy()
                    era_df["_year_num"] = pd.to_numeric(era_df["year"].astype(str).str[:4], errors="coerce")
                    era_df = era_df.dropna(subset=["_year_num"])
                    if len(era_df) >= 3:
                        era_df["_era"] = pd.cut(
                            era_df["_year_num"],
                            bins=[1998, 2010, 2016, 2022, 2030],
                            labels=["2005–2010", "2011–2016", "2017–2022", "2023–2026"]
                        )
                        era_stats = era_df.groupby("_era", observed=True).apply(
                            lambda g: pd.Series({
                                "n": len(g),
                                "rate": (g["decision_simple"].isin(["Recommended", "Optimised"]).sum() / len(g) * 100) if len(g) > 0 else None
                            })
                        ).dropna()

                        if len(era_stats) > 0:
                            st.markdown("**Recommendation rate by era (retrieved set):**")
                            fig_era = px.bar(
                                x=era_stats.index.astype(str), y=era_stats["rate"],
                                labels={"x": "Era", "y": "Recommendation rate (%)"},
                                text=[f"{r:.0f}% (n={int(n)})" for r, n in zip(era_stats["rate"], era_stats["n"])],
                                color_discrete_sequence=["#3498db"]
                            )
                            fig_era.update_traces(textposition="outside")
                            fig_era.update_layout(height=300, margin=dict(t=20, b=10), yaxis_range=[0, 105])
                            st.plotly_chart(fig_era, use_container_width=True)
                            st.caption(
                                "Recommendation rate = Recommended + Optimised as a share of retrieved "
                                "appraisals in that era; 'n' is the count behind each bar — treat eras "
                                "with small n as indicative only. Earlier eras may not reflect current "
                                "NICE methods, managed access frameworks, or NHS treatment pathways."
                            )
                except Exception:
                    pass

            similar_with_icer = similar[similar["icer_lower"].notna()]
            if len(similar_with_icer) > 0:
                approved_icer = similar_with_icer[similar_with_icer["decision_simple"].isin(["Recommended", "Optimised"])]

                if len(approved_icer) < 3:
                    st.markdown("**Published ICER data (limited):**")
                    st.warning(
                        f"Only {len(approved_icer)} appraisal(s) in this retrieval set have a "
                        f"publicly reported ICER. A sample this small is not a reliable benchmark "
                        f"range — treat the figure(s) below as a single historical data point, "
                        f"not an average or accepted range."
                    )
                else:
                    st.markdown("**Published ICER benchmarks from similar appraisals:**")

                if len(approved_icer) > 0:
                    col_i1, col_i2, col_i3 = st.columns(3)
                    with col_i1:
                        st.metric("Lowest reported ICER", f"£{approved_icer['icer_lower'].min():,.0f}")
                    with col_i2:
                        st.metric("Highest reported ICER", f"£{approved_icer['icer_lower'].max():,.0f}")
                    with col_i3:
                        avg = similar_with_icer["icer_lower"].mean()
                        st.metric("Your submitted ICER vs this figure", f"{'Above' if estimated_cost > avg else 'Below'} (£{avg:,.0f})")
                    st.caption(
                        "'ICER Lower' as reported in the workbook is not consistently labelled as "
                        "company base-case, EAG-corrected, or committee-preferred — treat as an "
                        "indicative published figure only, not a confirmed accepted value."
                    )
                    st.dataframe(
                        similar_with_icer[["drug_name", "indication", "icer_lower", "icer_upper", "decision_simple"]].rename(columns={
                            "drug_name": "Drug", "indication": "Indication",
                            "icer_lower": "Reported ICER Lower", "icer_upper": "Reported ICER Upper", "decision_simple": "Decision"
                        }),
                        use_container_width=True
                    )
                else:
                    st.info("No approved/optimised appraisals in this set have a reported ICER to compare against.")

            rejected_similar = similar[similar["decision_simple"] == "Not Recommended"]

            if len(rejected_similar) > 0:
                # ── Synthesised cross-appraisal themes ────
                ranked_themes, theme_sample_size, theme_sources = synthesise_themes(rejected_similar)
                patterns = [(label_text, count) for label_text, (emoji, count) in ranked_themes]
                if ranked_themes:
                    st.markdown(f"**Common themes across {theme_sample_size} rejected comparable appraisals:**")
                    theme_table = pd.DataFrame([
                        {"Theme": f"{emoji} {label_text}", "Frequency": f"{count}/{theme_sample_size}"}
                        for label_text, (emoji, count) in ranked_themes
                    ])
                    st.dataframe(theme_table, use_container_width=True, hide_index=True)

                    for label_text, (emoji, count) in ranked_themes:
                        source_ids = theme_sources.get(label_text, [])
                        with st.expander(f"Where '{label_text}' was raised"):
                            for aid in source_ids:
                                quote_row = rejected_similar[rejected_similar["appraisal_id"] == aid]
                                quote_text = None
                                if len(quote_row) > 0 and "original_nice_comment" in quote_row.columns:
                                    raw_quote = quote_row.iloc[0].get("original_nice_comment")
                                    if pd.notna(raw_quote):
                                        quote_text = str(raw_quote).strip()
                                        if len(quote_text) > 220:
                                            quote_text = quote_text[:220].rsplit(" ", 1)[0] + "…"
                                st.markdown(f"**{aid}**")
                                if quote_text:
                                    st.markdown(f"> {quote_text}")
                                else:
                                    st.caption("No verbatim committee text available for this appraisal — see full guidance link below.")
                    st.caption(
                        "Synthesised from committee reasoning text across the rejected appraisals "
                        "shown below. A theme count reflects how many of these specific appraisals "
                        "raised that concern — not a general base rate for the indication."
                    )

            has_detail_col = "primary_reason_category" in rejected_similar.columns
            rejected_with_reasoning = rejected_similar[rejected_similar["rejection_reasoning"].notna()]

            if len(rejected_with_reasoning) > 0:
                # Build the comparison set once, so each card can be checked
                # against what the OTHER rejected appraisals in this set raised.
                all_cards, concern_counter, comparison_size = build_concern_frequency(
                    rejected_with_reasoning, has_detail_col
                )

                st.markdown("**Individual rejected appraisals:**")
                if comparison_size > 1:
                    st.caption(
                        f"Each appraisal's concerns are compared against the other "
                        f"{comparison_size - 1} rejected appraisal(s) in this retrieved set, "
                        f"so you can see what's a common pattern versus specific to that drug."
                    )

                for row, card in all_cards[:5]:
                    label = f"{row['drug_name']} - {row['indication']} ({row['year']})"
                    with st.expander(label):
                        st.markdown(f"**Committee conclusion**")
                        st.write(card["conclusion"])

                        if comparison_size > 1:
                            shared, unique = split_shared_unique(card["concerns"], concern_counter, comparison_size)

                            if shared:
                                st.markdown(f"**Shared concerns** _(also raised in other rejected appraisals here)_")
                                for c, freq in shared:
                                    st.markdown(f"- {c} — shared with {freq}/{comparison_size} appraisals")

                            if unique:
                                st.markdown(f"**Unique to this appraisal**")
                                for c in unique:
                                    st.markdown(f"- {c}")

                            if not shared and not unique:
                                st.write("Specific concerns not itemised in source text — see full guidance below.")
                        else:
                            st.markdown(f"**Key evidence concerns**")
                            for c in card["concerns"]:
                                st.markdown(f"- {c}")

                        st.markdown(f"**Reported ICER**")
                        st.write(card["icer_line"])

                        if not has_detail_col or pd.isna(row.get("detailed_reasoning")):
                            st.caption(
                                "Structured from the standardised NICE recommendation-comment text. "
                                "This names the decision category but may not capture the full "
                                "committee discussion — see the linked guidance below."
                            )

                        with st.expander("Show full source text"):
                            st.write(card["raw"])

                        if pd.notna(row.get("url")):
                            st.markdown(f"[View NICE guidance — full committee discussion]({row['url']})")

        # ── Optimised appraisals in this retrieved set ────
        optimised_similar = similar[similar["decision_simple"] == "Optimised"]
        if len(optimised_similar) > 0:
            st.markdown(f"**Optimised appraisals in this retrieved set ({len(optimised_similar)}):**")
            st.caption(
                "These were recommended only within a restricted population or under "
                "specific conditions (e.g. biomarker subgroup, prior treatment requirement, "
                "commercial arrangement). Structured restriction-type data is not yet "
                "extracted for this dataset — this lists the appraisals so you can review "
                "the specific restrictions directly in NICE guidance."
            )
            opt_display = optimised_similar[["drug_name", "indication", "year", "appraisal_id", "url"]].head(10).rename(columns={
                "drug_name": "Drug", "indication": "Indication", "year": "Year",
                "appraisal_id": "Appraisal ID", "url": "NICE Link"
            })
            st.dataframe(opt_display, use_container_width=True, hide_index=True)

        # ── Evidence gaps suggested by historical precedent ──
        nonroutine_similar = similar[similar["decision_simple"].isin(["Not Recommended", "Terminated", "Managed Access"])]
        if len(nonroutine_similar) >= 3:
            gap_themes, gap_sample_size, _ = synthesise_themes(nonroutine_similar, max_examples=15)
            if gap_themes:
                st.markdown("**Evidence gaps suggested by historical precedent:**")
                st.caption(
                    f"Based on {gap_sample_size} non-routine appraisals (not recommended, "
                    f"terminated, or managed access) in this retrieved set. This does not "
                    f"mean NICE will raise the same issues for this submission — it "
                    f"indicates areas that have historically required careful justification "
                    f"in this space."
                )
                for label_text, (emoji, count) in gap_themes:
                    st.markdown(f"☑ {label_text}")

        generic_drugs = ["omeprazole", "lansoprazole", "metformin", "atorvastatin",
                         "amlodipine", "ramipril", "lisinopril", "simvastatin",
                         "docetaxel", "paclitaxel", "carboplatin", "cisplatin"]

        # ── Evidence confidence indicators ───────────────
        st.markdown("### Evidence Completeness")
        st.caption(
            "How much of this assessment rests on solid data versus a thin or keyword-only "
            "match — read this before the sections below."
        )

        scoring_active_conf = query_has_tags and has_tag_coverage(similar)
        similarity_conf = 9 if scoring_active_conf else (4 if query_has_tags else 2)

        icer_coverage = similar["icer_lower"].notna().sum() if total_similar > 0 else 0
        icer_conf = round(min(icer_coverage / max(total_similar, 1), 1.0) * 10)

        detail_coverage = similar["detailed_reasoning"].notna().sum() if "detailed_reasoning" in similar.columns and total_similar > 0 else 0
        rejected_total = max(rejected_count, 1)
        reasoning_conf = round(min(detail_coverage / rejected_total, 1.0) * 10) if rejected_count > 0 else 5

        sample_conf = round(min(total_similar / 10, 1.0) * 10)

        def _bar(n):
            return "█" * n + "░" * (10 - n)

        avg_conf = (similarity_conf + sample_conf + icer_conf + reasoning_conf) / 4
        if avg_conf >= 7:
            confidence_label = "High"
        elif avg_conf >= 4:
            confidence_label = "Moderate"
        else:
            confidence_label = "Low"

        st.markdown(f"**Similarity match quality** `{_bar(similarity_conf)}` {similarity_conf}/10")
        st.markdown(f"**Clinical precedent (sample size)** `{_bar(sample_conf)}` {sample_conf}/10 — {total_similar} appraisals retrieved")
        st.markdown(f"**Economic precedent (published ICER)** `{_bar(icer_conf)}` {icer_conf}/10 — {icer_coverage}/{total_similar} with a reported ICER")
        if rejected_count > 0:
            st.markdown(f"**Committee reasoning detail** `{_bar(reasoning_conf)}` {reasoning_conf}/10 — {detail_coverage}/{rejected_count} rejections with structured detail")
        else:
            st.markdown(f"**Committee reasoning detail** — no rejected appraisals in this retrieved set to assess")

        st.markdown(f"**Overall confidence: {confidence_label}**")

        st.caption(
            "Low bars mean a section below is descriptive of very little data — treat any "
            "pattern drawn from it as indicative only, not a robust finding."
        )

        st.markdown("### Evidence Summary")
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            st.markdown(f"""
**Retrieved appraisal set (by indication keyword match):**
- {total_similar} appraisals identified
- {recommended_count} recommended
- {optimised_count} optimised
- {rejected_count} not recommended
- {managed_count} managed access
- Recommendation proportion within this set: {approval_rate:.0f}%
            """)
            st.caption(
                "Descriptive only — this is the proportion of retrieved appraisals that were "
                "recommended, not a predicted probability that this drug will be recommended."
            )
        with col_e2:
            st.markdown(f"""
**Your submitted profile (hypothetical):**
- Submitted ICER: £{estimated_cost:,}/QALY
- WTP reference threshold: £{threshold:,}/QALY
- Position vs threshold: {((estimated_cost/threshold)-1)*100:+.0f}%
- QALYs gained: {qalys}
- Comparator: {comparator or 'Not specified'}
            """)
            st.caption(
                "These are the figures you entered, not a historical or verified NICE value."
            )

        st.markdown("**Contextual considerations:**")
        warnings_list = []
        context_facts = []

        try:
            years_series = similar["year"].dropna().astype(str)
            if len(years_series) > 0:
                oldest = sorted(years_series)[0]
                newest = sorted(years_series)[-1]
                oldest_year_num = oldest[:4]
                newest_year_num = newest[:4]
                if oldest_year_num.isdigit() and newest_year_num.isdigit():
                    span = int(newest_year_num) - int(oldest_year_num)
                    if span >= 10:
                        warnings_list.append(
                            f"Retrieved appraisals span {oldest} to {newest} ({span} years). NICE "
                            f"methods, treatment pathways, comparator prices, and clinical practice "
                            f"have likely changed materially over that period."
                        )
                    else:
                        context_facts.append(f"Retrieved appraisals span {oldest} to {newest} ({span} years).")
        except Exception:
            pass

        if comparator and any(g in comparator.lower() for g in generic_drugs):
            warnings_list.append(f"{comparator} is now a low-cost generic. Historical ICERs involving it as a comparator may understate the true incremental cost burden versus current NHS pricing.")

        if total_similar < 5:
            warnings_list.append(f"Small retrieval set - only {total_similar} similar appraisal(s) found. Treat any pattern or rate drawn from this set with caution.")
        else:
            context_facts.append(f"Retrieval set size: {total_similar} appraisals.")

        if not keyword:
            warnings_list.append("No indication keyword entered - benchmarking against the full database rather than a targeted indication match.")

        # ICER data availability within the retrieved set
        if total_similar > 0:
            icer_available = similar["icer_lower"].notna().sum()
            icer_pct = icer_available / total_similar * 100
            if icer_pct < 20:
                warnings_list.append(
                    f"Only {icer_available} of {total_similar} retrieved appraisals ({icer_pct:.0f}%) "
                    f"have a publicly reported ICER — most modern appraisals keep this commercially "
                    f"confidential, so ICER-based benchmarking here is necessarily thin."
                )
            else:
                context_facts.append(f"Published ICER available for {icer_available} of {total_similar} retrieved appraisals ({icer_pct:.0f}%).")

        # Therapeutic area concentration
        if total_similar >= 3 and "therapeutic_area" in similar.columns:
            area_counts = similar["therapeutic_area"].dropna().value_counts()
            if len(area_counts) > 0:
                top_area = area_counts.index[0]
                top_pct = area_counts.iloc[0] / total_similar * 100
                if top_pct >= 60:
                    context_facts.append(f"{top_pct:.0f}% of retrieved appraisals are in {top_area} — precedent is concentrated in this area rather than spread across therapeutic areas.")

        if warnings_list:
            for w in warnings_list:
                st.warning(w)

        if context_facts:
            st.markdown("_Additional context:_")
            for c in context_facts:
                st.markdown(f"- {c}")

        if not warnings_list and not context_facts:
            st.info("No major contextual concerns identified.")

        # compute termination rate safely (single source of truth, calculated once)
        termination_rate = (terminated_count / total_similar * 100) if total_similar > 0 else 0

        if termination_rate > 50:
            warnings_list.append(
                f"High termination rate: {terminated_count} of {total_similar} similar "
                f"appraisals ({termination_rate:.0f}%) were terminated without a submitted "
                f"evidence package. This is often a sign that manufacturers were unable to "
                f"agree a commercially viable price with NICE — a submitted ICER below "
                f"threshold does not on its own overcome that pattern."
            )

        # determine risk signal (deliberately not framed as a prediction)
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
            "low":                  "Likely Recommended",
            "moderate":             "Borderline",
            "high":                 "Unlikely to be Recommended",
        }[signal]

        st.markdown("**Historical precedent review**")
        st.caption(
            "This is a descriptive signal based on your submitted ICER versus the reference "
            "threshold and retrieved precedent — not a prediction of the committee's decision. "
            "A recommendation cannot be inferred without the full evidence package, model "
            "structure, committee-preferred assumptions, and any confidential commercial "
            "arrangement."
        )

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
- Stress-test the clinical evidence base versus {comparator or 'the stated comparator'} (maturity of survival data, trial design)
- {'Confirm end-of-life criteria are met and evidenced explicitly' if end_of_life == 'Yes' else 'Consider whether CDF/managed access is a fallback if evidence is still maturing'}
- Anticipate that a confidential commercial arrangement is often expected even below threshold
- Note: {optimised_count} appraisal(s) in this retrieved set were approved only with conditions — review what those conditions were
            """)
        elif signal == "moderate":
            st.warning(f"""
Position: Submitted ICER exceeds the reference threshold by {((estimated_cost/threshold)-1)*100:.0f}%.

Possible next steps:
- Review the principal drivers of incremental cost and QALY gain in the model
- Test alternative, evidence-supported assumptions (survival extrapolation, utilities, retreatment)
- Model price or commercial-arrangement scenarios that would bring the ICER within range
- Assess whether {comparator or 'the stated comparator'} reflects current NHS practice
- Explore Cancer Drugs Fund / managed access as a contingency route
            """)
        else:
            st.error(f"""
Position: Submitted ICER exceeds the reference threshold by {((estimated_cost/threshold)-1)*100:.0f}%.

Possible next steps:
- Review the principal drivers of incremental cost and QALY gain — is the model biased toward a favourable case?
- Re-examine whether the clinical evidence (survival, response) is mature enough to support the QALY estimate
- Test alternative retreatment, extrapolation, and utility assumptions for sensitivity
- Assess whether the comparator reflects current NHS practice and pricing
- Seek NICE scientific advice before a formal submission
            """)

        st.caption(
            "This assessment is a preliminary, evidence-retrieval-based signal only. It does not "
            "constitute a prediction of a NICE committee decision and should not replace full "
            "economic modelling, evidence review, or professional market access advice."
        )

        st.markdown("---")
        pdf_buffer = generate_assessment_pdf(
            drug_name, indication, estimated_cost, qalys,
            end_of_life, comparator, threshold,
            total_similar, recommended_count, optimised_count,
            rejected_count, managed_count, approval_rate,
            similar, patterns, warnings_list, verdict
        )
        st.download_button(
            label="📥 Download PDF Report",
            data=pdf_buffer,
            file_name=f"{drug_name.replace(' ', '_')}_market_access_report.pdf",
            mime="application/pdf",
            type="primary"
        )

    else:
        st.warning("Please enter a drug name and indication.")

st.divider()
st.caption("Built with Python & Streamlit | 1,439 appraisals sourced from NICE Technology Appraisals | Preliminary intelligence tool — not a substitute for full economic modelling or professional market access advice")
