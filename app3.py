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
    df = pd.read_excel(r"C:\Users\axelz\NICE_final_v6.xlsx")
    return df

df = load_data()

def get_rejection_patterns(similar_df):
    patterns = []
    for text in similar_df["rejection_patterns"].dropna():
        if text:
            patterns.extend(text.split(" | "))
    return Counter(patterns).most_common(5)

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
        "LIKELY RECOMMENDED - cost-effective at the current threshold"
        if "Likely" in verdict
        else "BORDERLINE - may require managed access route"
        if "Borderline" in verdict
        else "UNLIKELY TO BE RECOMMENDED - exceeds threshold significantly"
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
        f"{drug_name} for {indication} has been assessed against {total_similar} similar NICE "
        f"technology appraisals. With an estimated ICER of £{estimated_cost:,}/QALY against a "
        f"willingness-to-pay threshold of £{threshold:,}/QALY (+{((estimated_cost/threshold)-1)*100:.0f}%), "
        f"the drug appears <b>{verdict_text.lower()}</b>. "
        f"The historical approval rate for similar indications is {approval_rate:.0f}% "
        f"({recommended_count} recommended, {optimised_count} optimised, {rejected_count} not recommended).",
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
        ["Approval Rate",   f"{approval_rate:.0f}%", ""],
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
    sim_data = [["TA ID", "Drug", "Indication", "Decision", "Year"]]
    for _, row in similar_top.iterrows():
        sim_data.append([
            str(row["appraisal_id"]),
            str(row["drug_name"])[:18],
            str(row["indication"])[:38],
            str(row["decision_simple"]),
            str(row["year"])
        ])
    sim_table = Table(sim_data, colWidths=[1.8*cm, 3.5*cm, 7*cm, 3.2*cm, 2*cm])
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
st.markdown("*1,439 pharmaceutical appraisals - complete NICE database*")
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
st.subheader("🤖 AI Market Access Advisor")
st.markdown("*Enter a drug profile - assessed against 1,435 NICE decisions*")

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

if st.button("Get Market Access Assessment", type="primary"):
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
            st.metric("Historical Approval Rate", f"{approval_rate:.0f}%")

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
            st.dataframe(
                similar[["appraisal_id", "drug_name", "indication", "decision_simple", "year"]].head(10),
                use_container_width=True
            )

            similar_with_icer = similar[similar["icer_lower"].notna()]
            if len(similar_with_icer) > 0:
                st.markdown("**ICER benchmarks from similar appraisals:**")
                approved_icer = similar_with_icer[similar_with_icer["decision_simple"].isin(["Recommended", "Optimised"])]
                col_i1, col_i2, col_i3 = st.columns(3)
                with col_i1:
                    st.metric("Lowest ICER accepted", f"£{approved_icer['icer_lower'].min():,.0f}" if len(approved_icer) > 0 else "No data")
                with col_i2:
                    st.metric("Highest ICER accepted", f"£{approved_icer['icer_lower'].max():,.0f}" if len(approved_icer) > 0 else "No data")
                with col_i3:
                    avg = similar_with_icer["icer_lower"].mean()
                    st.metric("Your ICER vs benchmark", f"{'Above' if estimated_cost > avg else 'Below'} average (£{avg:,.0f})")
                st.dataframe(
                    similar_with_icer[["drug_name", "indication", "icer_lower", "icer_upper", "decision_simple"]].rename(columns={
                        "drug_name": "Drug", "indication": "Indication",
                        "icer_lower": "ICER Lower", "icer_upper": "ICER Upper", "decision_simple": "Decision"
                    }),
                    use_container_width=True
                )

            rejected_similar = similar[similar["decision_simple"] == "Not Recommended"]
            if len(rejected_similar) > 0:
                patterns = get_rejection_patterns(rejected_similar)
                if patterns:
                    st.markdown("**Common rejection reasons in similar appraisals:**")
                    for pattern, count in patterns:
                        st.markdown(f"- **{pattern}** - seen in {count} similar rejections")

            rejected_with_reasoning = rejected_similar[rejected_similar["rejection_reasoning"].notna()]
            if len(rejected_with_reasoning) > 0:
                st.markdown("**Similar drugs that were rejected:**")
                for _, row in rejected_with_reasoning.head(5).iterrows():
                    with st.expander(f"{row['drug_name']} - {row['indication']} ({row['year']})"):
                        st.write(row["rejection_reasoning"])
                        if pd.notna(row.get("url")):
                            st.markdown(f"[View NICE guidance]({row['url']})")

        generic_drugs = ["omeprazole", "lansoprazole", "metformin", "atorvastatin",
                         "amlodipine", "ramipril", "lisinopril", "simvastatin",
                         "docetaxel", "paclitaxel", "carboplatin", "cisplatin"]

        st.markdown("### Evidence Summary")
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            st.markdown(f"""
**Appraisal landscape:**
- {total_similar} similar NICE appraisals identified
- {recommended_count} recommended
- {optimised_count} optimised
- {rejected_count} not recommended
- {managed_count} managed access
- Historical approval rate: {approval_rate:.0f}%
            """)
        with col_e2:
            st.markdown(f"""
**Economic profile:**
- Your ICER: £{estimated_cost:,}/QALY
- WTP threshold: £{threshold:,}/QALY
- Position vs threshold: {((estimated_cost/threshold)-1)*100:+.0f}%
- QALYs gained: {qalys}
- Comparator: {comparator or 'Not specified'}
            """)

        st.markdown("**Contextual considerations:**")
        warnings_list = []
        try:
            oldest = str(similar["year"].min())[:7]
            if oldest < "2015":
                warnings_list.append(f"Similar appraisals date from {oldest} - NICE methodology has changed significantly since then.")
        except:
            pass
        if comparator and any(g in comparator.lower() for g in generic_drugs):
            warnings_list.append(f"{comparator} is now a low-cost generic. Historical ICERs may understate the true incremental cost burden.")
        if total_similar < 5:
            warnings_list.append(f"Limited precedent - only {total_similar} similar appraisals found.")
        if not keyword:
            warnings_list.append("No indication keyword entered - benchmarking against full database.")
            

        if warnings_list:
            for w in warnings_list:
                st.warning(w)
        else:
            st.info("No major contextual concerns identified.")

        # compute termination rate safely
        termination_rate = (terminated_count / total_similar * 100) if total_similar > 0 else 0

        # determine verdict
        if termination_rate == 100 and total_similar >= 2:
            verdict = "High Commercial Risk"
        elif termination_rate > 75 and total_similar >= 3:
            verdict = "High Commercial Risk"
        elif estimated_cost <= threshold:
            verdict = "Likely Recommended"
        elif estimated_cost <= threshold * 1.5:
            verdict = "Borderline"
        else:
            verdict = "Unlikely to be Recommended"
        if verdict == "High Commercial Risk":
            st.error(f"""
High Commercial Risk - ICER alone is insufficient

Despite an ICER of £{estimated_cost:,}/QALY appearing cost-effective,
{termination_rate:.0f}% of similar appraisals were terminated without
a NICE recommendation. This strongly suggests:

- Manufacturers cannot achieve a commercially viable price with NICE
- The indication may have structural pricing challenges
- Standard technology appraisal may not be the right route

Recommended actions:
- Investigate Highly Specialised Technologies pathway eligibility
- Conduct early NICE scientific advice before formal submission
- Model multiple price scenarios - list price vs net price
- Consider patient access scheme or managed access agreement
- Review whether UK launch is commercially viable at any price
            """)

        elif estimated_cost <= threshold:
            st.success(f"""
Likely cost-effective at the {'end-of-life' if end_of_life == 'Yes' else 'standard'} threshold of £{threshold:,}/QALY.
- Ensure robust clinical evidence vs {comparator or 'comparator'}
- {'End-of-life criteria must be clearly evidenced' if end_of_life == 'Yes' else 'Consider CDF route if evidence is immature'}
- Commercial negotiation likely required
- {optimised_count} similar drugs approved with conditions - prepare for optimisation
            """)

        elif estimated_cost <= threshold * 1.5:
            st.warning(f"""
Borderline - exceeds threshold by {((estimated_cost/threshold)-1)*100:.0f}%.
- Explore Cancer Drugs Fund / managed access route
- Strengthen evidence package vs {comparator or 'comparator'}
- Consider price reduction to bring ICER below £{threshold:,}
- Conduct PSA to quantify uncertainty
            """)

        else:
            st.error(f"""
Unlikely to be recommended - exceeds threshold by {((estimated_cost/threshold)-1)*100:.0f}%.
- Substantial price reduction required
- Re-examine QALY estimates - are they robust?
- Consider alternative indication with stronger evidence
- Engage NICE scientific advice before submission
            """)

        st.caption("Further economic modelling is strongly recommended before drawing conclusions.")

        # 4. PDF download
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
st.caption("Built with Python & Streamlit | 1,439 appraisals sourced from NICE Technology Appraisals")
