"""
Both chat implementations, their prompts and the context they are given.

Moved out of app3.py by split_app3.py. These functions reach only
constants, never the dataframe or session state, which is what made
the move mechanical.
"""

import re
import requests
from collections import Counter
from datetime import date
import pandas as pd
import streamlit as st
from hta_data import *
from hta_core import *  # noqa: F401,F403
from hta_nice import *  # noqa: F401,F403
from hta_ui import *  # noqa: F401,F403

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
                "max_tokens": 2000,
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
GBA_OUTCOME_ORDER = [
    "Major added benefit",
    "Considerable added benefit",
    "Minor added benefit",
    "Non-quantifiable added benefit",
    GBA_DEEMED,
    "Mixed by subgroup",
    "No added benefit proven",
    "Lesser benefit",
    "In progress",
    "Discontinued",
    "Exempted from assessment",
    GBA_NOT_EXTRACTED,
    "Unclassified",
]
GBA_MD_COLORS = {
    "Major added benefit": "green", "Considerable added benefit": "green",
    "Minor added benefit": "green", "Non-quantifiable added benefit": "blue",
    GBA_DEEMED: "green", "Mixed by subgroup": "orange", "No added benefit proven": "red",
    "Lesser benefit": "red",
}
HAS_OUTCOME_ORDER = [
    "Early access granted",
    "Early access renewed",
    "Early access refused",
    "Early access — other",
    "ASMR I", "ASMR II", "ASMR III", "ASMR IV", "ASMR V",
    "SMR sufficient, no ASMR recorded",
    "SMR mixed by population",
    "SMR insufficient",
    HAS_NO_RATING,
]
HAS_MD_COLORS = {
    "Early access granted": "green", "Early access renewed": "blue",
    "Early access refused": "red", "ASMR I": "green", "ASMR II": "green",
    "ASMR III": "green", "ASMR IV": "green", "ASMR V": "gray",
    "SMR sufficient, no ASMR recorded": "violet", "SMR mixed by population": "orange",
    "SMR insufficient": "red",
}
HAS_PROCEDURE_ORDER = ["First listing", "Extension of indication", "Early access",
                       "Re-evaluation", "Renewal", "Label or listing change", "Delisting",
                       "Other", "Not recorded"]
GBA_PROCEDURE_ORDER = ["Benefit assessment", "New indication", "Re-assessment"]
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
always shown alongside it. Procedures without an extent rating are named for why:
*In progress*, *Discontinued*, *Exempted from assessment*, *Added benefit deemed proven*
(reserve antibiotics under §35a(1c) SGB V, where no extent is rated) or *Concluded —
outcome not extracted* (a finished procedure whose resolution isn't in this extract).
Headline percentages count only assessments with a rating.

**Therapy areas and search.** Therapy areas use one English vocabulary for NICE, the G-BA
and HAS, mapped from the G-BA's own therapy-area tag. The resolutions describe conditions
in German, so English searches are translated (breast cancer → Mammakarzinom, myeloma →
Myelom); German terms work as typed.

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
inferred. Opinions with no SMR or ASMR in the source metadata — mostly renewals, label
changes and delistings — show as *No SMR/ASMR recorded*.

**Therapy areas and search.** Therapy areas use one English vocabulary for NICE, the G-BA
and HAS, mapped from HAS's own categories. HAS files cancers under their body system (breast
cancer under *Sein*), so any category naming a cancer counts as Oncology; where HAS gives no
category (mostly opinions before 2011) the area comes from the ATC code. The opinions are
in French, so English searches are translated (myeloma → myélome, breast cancer → cancer du
sein); French terms work as typed.

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
AGENCY_CONFIG = {
    "G-BA": {
        "sheet": "G-BA_Recent",
        "slug": "gba",
        "title": "🇩🇪 G-BA Benefit Assessment Intelligence",
        "noun": "assessments",
        "noun_title": "Assessments",
        "detail_title": "📋 Assessment Detail",
        "search_examples": "'breast cancer', 'NSCLC', 'Mammakarzinom'",
        "language": "German",
        "condition_placeholder": "e.g. breast cancer, NSCLC, myeloma, Mammakarzinom",
        "procedure_order": GBA_PROCEDURE_ORDER,
        "source_label": "official G-BA resolutions and IQWiG assessments",
        "log_keyword": "G-BA",
        "extra_columns": ["comparator_therapy", "added_benefit_assessment", "ICER_EUR_per_QALY",
                          "procedure_status", "orphan_status"],
        "comparator_col": "comparator_therapy",
        "icer_lower_col": "ICER_EUR_per_QALY",
        "icer_upper_col": None,
        "outcome_fn": gba_outcome_with_status,
        "favourable_fn": gba_favourable_with_status,
        "derived": {"certainty": gba_certainty},
        "outcome_order": GBA_OUTCOME_ORDER,
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
        "search_examples": "'myeloma', 'lymphoma', 'cancer du sein'",
        "language": "French",
        "condition_placeholder": "e.g. myeloma, lung cancer, psoriasis, cancer du sein",
        "procedure_order": HAS_PROCEDURE_ORDER,
        "source_label": "official HAS decisions and Transparency Committee documents",
        "log_keyword": "HAS",
        "extra_columns": ["comparator", "SMR_rating", "ASMR_rating", "ICER_lower_EUR_per_QALY",
                          "ICER_upper_EUR_per_QALY", "current_status", "source_title",
                          "ATC_codes"],
        "comparator_col": "comparator",
        "icer_lower_col": "ICER_lower_EUR_per_QALY",
        "icer_upper_col": "ICER_upper_EUR_per_QALY",
        "outcome_fn": has_outcome_labelled,
        "favourable_fn": has_favourable,
        "derived": {
            "_early_access": has_is_early_access,
            "_smr": lambda r: smr_level(r.get("SMR_rating")),
            "_asmr": lambda r: asmr_level(r.get("ASMR_rating")),
            "_ratings": has_ratings_text,
        },
        "outcome_order": HAS_OUTCOME_ORDER,
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
    st.markdown(f"**Ask AI about the {len(frame):,} {cfg['noun']} your search found**")

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

    shown = min(len(frame), CHAT_MAX_CONTEXT_ROWS)
    st.caption(
        "e.g. \"which comparators came up most?\" or \"how did NICE decide on the same "
        "molecules?\". Answers use only these results — the "
        + (f"{shown} most relevant of {len(frame):,}" if shown < len(frame) else f"{shown}")
        + " — and say so if something isn't covered, rather than guessing. Narrow the search "
          "to put a specific set in front of it.")
    turns_used = len(history) // 2
    counter = st.empty()   # filled in again below once a new answer has landed
    counter.caption(f"{turns_used}/{CHAT_MAX_TURNS} questions used for this selection.")

    for msg in history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if turns_used >= CHAT_MAX_TURNS:
        st.warning("Question limit reached for these results. Run a new search to start a "
                   "new conversation.")
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
                counter.caption(f"{len(history) // 2}/{CHAT_MAX_TURNS} questions used for "
                                f"this selection.")
