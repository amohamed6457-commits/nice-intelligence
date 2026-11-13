"""
Reading an HTA decision: text handling, G-BA and HAS outcomes, classification, and query expansion.

Moved out of app3.py by split_app3.py. These functions reach only
constants, never the dataframe or session state, which is what made
the move mechanical.
"""

import re
import warnings
from datetime import date
from datetime import datetime
import pandas as pd
import streamlit as st
from hta_data import *

__all__ = [
    "CONDITION_TERMS",
    "FAMILY_NO_RATING",
    "GBA_DEEMED",
    "GBA_NOT_EXTRACTED",
    "GBA_NO_RATING",
    "GBA_POSITIVE",
    "GBA_STATUS_OUTCOMES",
    "HAS_NO_RATING",
    "OUTCOME_FAMILIES",
    "_ALL_ENTRY",
    "_ASMR_ORDER",
    "_ATC_AREA_RULES",
    "_ATC_CODE",
    "_DESCRIPTIVE_TAIL",
    "_DESCRIPTOR_WORDS",
    "_ENGLISH_AREA_EN",
    "_FOLD_TABLE",
    "_GBA_AREA_EN",
    "_GBA_CERTAINTY_PATTERNS",
    "_GBA_DEEMED_PATTERN",
    "_GBA_EXTENT_PATTERNS",
    "_HAS_EA_NEGATIVE",
    "_HAS_SYSTEM_EN",
    "_INDICATION_STOPWORDS",
    "_MOLECULE_SEPARATORS",
    "_NON_MOLECULES",
    "_NOTE_PREFIXES",
    "_PLACEHOLDER_VALUES",
    "_PRODUCTION_NOTE",
    "_QUERY_STOPWORDS",
    "_SALT_WORDS",
    "_SMR_PATTERNS",
    "_any_term_regex",
    "_clauses",
    "_clean_inn",
    "_date_span",
    "_decision_text",
    "_english_area",
    "_extents_in",
    "_fold",
    "_fold_series",
    "_gba_extents",
    "_indication_tokens",
    "_inn_stem",
    "_is_stopword",
    "_md_link",
    "_md_safe",
    "_parse_date",
    "_shorten",
    "_substantive",
    "_term_body",
    "_term_pattern",
    "_text",
    "_theme_mode",
    "_to_datetime",
    "asmr_level",
    "condition_label",
    "expand_condition_query",
    "family_order",
    "format_eur_icer",
    "gba_certainty",
    "gba_favourable",
    "gba_favourable_with_status",
    "gba_headlines",
    "gba_outcome",
    "gba_outcome_with_status",
    "gba_procedure_group",
    "gba_subtitle",
    "has_favourable",
    "has_headlines",
    "has_is_early_access",
    "has_outcome",
    "has_outcome_labelled",
    "has_procedure_group",
    "has_ratings_text",
    "has_subtitle",
    "indication_overlap",
    "molecule_set",
    "orphan_flag",
    "outcome_family",
    "smr_level",
    "therapy_areas",
]

_PLACEHOLDER_VALUES = {"", "nan", "none", "n/a", "na", "-", "—", "not specified", "not reported",
                       "sans objet", "non applicable", "non évalué", "non evalue"}
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
GBA_DEEMED = "Added benefit deemed proven"
GBA_NOT_EXTRACTED = "Concluded — outcome not extracted"
GBA_POSITIVE = {
    "Major added benefit", "Considerable added benefit",
    "Minor added benefit", "Non-quantifiable added benefit",
}
GBA_STATUS_OUTCOMES = {
    "verfahren begonnen": "In progress",
    "beschlussfassung wird vorbereitet": "In progress",
    "stellungnahmeverfahren eröffnet": "In progress",
    "verfahrensstart verschoben": "In progress",
    "verfahren ausgesetzt": "In progress",
    "verfahren eingestellt": "Discontinued",
    "wurde freigestellt": "Exempted from assessment",
}
GBA_NO_RATING = set(GBA_STATUS_OUTCOMES.values()) | {GBA_NOT_EXTRACTED, "Unclassified"}
_GBA_DEEMED_PATTERN = r"gilt als belegt|deemed (?:established|proven)"
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
    if row.get("outcome") in {"No added benefit proven", GBA_DEEMED} | GBA_NO_RATING:
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
def gba_outcome_with_status(row):
    """
    gba_outcome, plus a name for every procedure that has no extent rating.

    On the full catalogue about one assessment in five had no rating, and
    reading them all as 'Unclassified' made them look like a data problem.
    Most are simply not rated (yet): the procedure is still running, was
    discontinued or exempted, or it's a reserve antibiotic whose added
    benefit is deemed proven by law. What's left is a concluded procedure
    whose resolution isn't in the extract, and it's labelled as that.
    """
    outcome = gba_outcome(row)
    if outcome != "Unclassified":
        return outcome
    text = f"{_decision_text(row)} {_text(row.get('added_benefit_assessment')) or ''}"
    if re.search(_GBA_DEEMED_PATTERN, text, re.IGNORECASE):
        return GBA_DEEMED
    status = (_text(row.get("procedure_status")) or "").lower()
    if status in GBA_STATUS_OUTCOMES:
        return GBA_STATUS_OUTCOMES[status]
    if status == "verfahren abgeschlossen":
        return GBA_NOT_EXTRACTED
    return "Unclassified"
def gba_favourable_with_status(row):
    if row.get("outcome") == GBA_DEEMED:
        return True
    if row.get("outcome") in GBA_NO_RATING:
        return None
    return gba_favourable(row)
def gba_subtitle(frame):
    return f"{len(frame):,} AMNOG early benefit assessments (§35a SGB V)"
def gba_headlines(frame):
    rated = frame[~frame["outcome"].isin(GBA_NO_RATING)]
    n = len(rated)
    positive = int(rated["favourable"].eq(True).sum())
    none_proven = int(rated["favourable"].eq(False).sum())
    pct = (lambda k: f" ({k / n * 100:.0f}%)" if n else "")
    return [
        ("Added benefit in ≥1 subgroup", f"{positive:,}/{n:,}{pct(positive)}",
         "Of the assessments with a rating: major, considerable, minor or non-quantifiable "
         "added benefit, benefit deemed proven, plus mixed resolutions where at least one "
         "subgroup received a positive rating. Procedures still in progress, discontinued, "
         "exempted or without an extracted outcome are left out of the count. Not equivalent "
         "to NICE's recommendation proportion: a medicine without added benefit is still "
         "reimbursed in Germany; the rating shapes its negotiated price."),
        ("No added benefit proven", f"{none_proven:,}/{n:,}{pct(none_proven)}",
         "Of the assessments with a rating: no added benefit proven for any subgroup, or a "
         "lesser benefit — including mixed resolutions where every subgroup's rating was "
         "negative. Mixed resolutions naming only a certainty level, not an extent, are in "
         "neither count."),
    ]
HAS_NO_RATING = "No SMR/ASMR recorded"
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
def has_outcome_labelled(row):
    """has_outcome, with 'Unclassified' named for what it is on the real
    catalogue: mostly renewals, label changes and delistings that carry no
    SMR/ASMR rating in the source metadata."""
    outcome = has_outcome(row)
    return HAS_NO_RATING if outcome == "Unclassified" else outcome
def has_favourable(row):
    """ASMR V (no improvement) is neutral, not a refusal: the medicine can still
    be reimbursed, it just gains no price premium."""
    outcome = row.get("outcome")
    if outcome in ("Early access granted", "Early access renewed",
                   "ASMR I", "ASMR II", "ASMR III", "ASMR IV"):
        return True
    if outcome in ("Early access refused", "SMR insufficient"):
        return False
    # No SMR/ASMR in the metadata, but the opinion itself was against listing.
    if outcome == HAS_NO_RATING and re.search(
            r"\bunfavou?rable\b|\bd[ée]favorable\b", _decision_text(row)):
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
        out.append(("Early access granted", f"{ok:,}/{len(early):,}",
                    "Early-access decisions in this selection that granted access, including "
                    "renewals of an earlier authorisation."))
    asmr = frame[frame["_asmr"].notna()]
    if len(asmr):
        ok = int(asmr["_asmr"].isin(["I", "II", "III", "IV"]).sum())
        out.append(("ASMR I–IV", f"{ok:,}/{len(asmr):,}",
                    "Reimbursement opinions recognising some improvement in clinical benefit "
                    "(ASMR V = none)."))
    smr = frame[frame["_smr"].notna()]
    if len(smr):
        ok = int((smr["_smr"] != "Insufficient").sum())
        out.append(("Sufficient SMR", f"{ok:,}/{len(smr):,}",
                    "Opinions rating the actual clinical benefit sufficient for reimbursement "
                    "in at least one population."))
    return out
OUTCOME_FAMILIES = {
    "G-BA": [
        ("Added benefit", GBA_POSITIVE | {GBA_DEEMED}),
        ("Mixed by subgroup", {"Mixed by subgroup"}),
        ("No added benefit", {"No added benefit proven", "Lesser benefit"}),
    ],
    "HAS": [
        ("Improvement (ASMR I–IV)", {"ASMR I", "ASMR II", "ASMR III", "ASMR IV"}),
        ("No improvement (ASMR V)", {"ASMR V"}),
        ("SMR only, no ASMR", {"SMR sufficient, no ASMR recorded", "SMR mixed by population"}),
        ("SMR insufficient", {"SMR insufficient"}),
        ("Early access", {"Early access granted", "Early access renewed", "Early access refused",
                          "Early access — other"}),
    ],
}
FAMILY_NO_RATING = {"G-BA": "No rating (yet)", "HAS": "No rating recorded"}
def family_order(agency):
    return [name for name, _ in OUTCOME_FAMILIES[agency]] + [FAMILY_NO_RATING[agency]]
def outcome_family(agency, outcome):
    for name, members in OUTCOME_FAMILIES[agency]:
        if outcome in members:
            return name
    return FAMILY_NO_RATING[agency]
def _theme_mode():
    """'dark' or 'light' — the viewer's current Streamlit theme."""
    try:
        return "dark" if st.context.theme.type == "dark" else "light"
    except Exception:
        return "light"
_FOLD_TABLE = str.maketrans({
    "à": "a", "á": "a", "â": "a", "ä": "a", "ã": "a", "å": "a", "ç": "c", "è": "e", "é": "e",
    "ê": "e", "ë": "e", "ì": "i", "í": "i", "î": "i", "ï": "i", "ñ": "n", "ò": "o", "ó": "o",
    "ô": "o", "ö": "o", "õ": "o", "ù": "u", "ú": "u", "û": "u", "ü": "u", "ý": "y", "ÿ": "y",
    "œ": "oe", "æ": "ae", "ß": "ss", "’": "'", "‘": "'", "´": "'", "β": "beta", "–": "-",
    "—": "-", " ": " ",
})
def _fold(value):
    """Lower-case, accent-free text, so 'Myélome', 'myelome' and 'MYÉLOME' match."""
    s = _text(value)
    return s.lower().translate(_FOLD_TABLE) if s else ""
def _fold_series(series, limit=None):
    s = series.map(_text).fillna("")
    if limit:
        s = s.str.slice(0, limit)
    return s.str.lower().str.translate(_FOLD_TABLE)
_GBA_AREA_EN = {
    "onkologische erkrankungen": "Oncology",
    "stoffwechselkrankheiten": "Endocrine / Metabolic",
    "infektionskrankheiten": "Infectious disease",
    "krankheiten des nervensystems": "Neurology",
    "krankheiten des blutes und der blutbildenden organe": "Haematology",
    "hauterkrankungen": "Dermatology",
    "krankheiten des muskel-skelett-systems": "Musculoskeletal",
    "krankheiten des atmungssystems": "Respiratory",
    "krankheiten des verdauungssystems": "Gastroenterology / Hepatology",
    "herz-kreislauf-erkrankungen": "Cardiovascular",
    "augenerkrankungen": "Ophthalmology",
    "krankheiten des urogenitalsystems": "Renal / Urology",
    "psychische erkrankungen": "Mental health",
    "koloniestimulierende faktoren": "Haematology",
    "sonstiges": "Other",
}
_HAS_SYSTEM_EN = {
    "douleur - inflammation": "Pain & inflammation",
    "peau et tissu conjonctif": "Dermatology",
    "appareil digestif": "Gastroenterology / Hepatology",
    "maladies genetiques, rares, congenitales": "Rare / Genetic diseases",
    "infections": "Infectious disease",
    "appareil respiratoire": "Respiratory",
    "sang et lymphe": "Haematology",
    "metabolisme et systeme endocrinien": "Endocrine / Metabolic",
    "appareil genital": "Reproductive & sexual health",
    "grossesse et reproduction": "Reproductive & sexual health",
    "sein": "Reproductive & sexual health",
    "appareil locomoteur": "Musculoskeletal",
    "systeme immunitaire": "Immunology / Inflammatory",
    "coeur": "Cardiovascular",
    "vaisseaux sanguins": "Cardiovascular",
    "oeil": "Ophthalmology",
    "troubles mentaux et comportementaux": "Mental health",
    "systeme nerveux": "Neurology",
    "systeme urinaire": "Renal / Urology",
    "cancers - divers": "Oncology",
    "bouche et dents": "Other",
    "iatrogenie - intoxications": "Other",
    "orl": "Other",
    "handicap": "Other",
}
_ENGLISH_AREA_EN = {
    "oncology": "Oncology", "cancer": "Oncology", "haematology": "Haematology",
    "hematology": "Haematology", "immunology": "Immunology / Inflammatory",
    "infectious": "Infectious disease", "endocrine": "Endocrine / Metabolic",
    "metabolic": "Endocrine / Metabolic", "neurology": "Neurology", "mental": "Mental health",
    "cardiovascular": "Cardiovascular", "respiratory": "Respiratory",
    "gastroenterology": "Gastroenterology / Hepatology", "renal": "Renal / Urology",
    "musculoskeletal": "Musculoskeletal", "dermatology": "Dermatology",
    "ophthalmology": "Ophthalmology", "rare": "Rare / Genetic diseases",
}
_ATC_AREA_RULES = [
    ("L01", "Oncology"), ("L02", "Oncology"), ("L03", "Immunology / Inflammatory"),
    ("L04", "Immunology / Inflammatory"), ("A10", "Endocrine / Metabolic"),
    ("A11", "Endocrine / Metabolic"), ("A12", "Endocrine / Metabolic"),
    ("A16", "Rare / Genetic diseases"), ("A", "Gastroenterology / Hepatology"),
    ("B05", "Other"), ("B", "Haematology"), ("C", "Cardiovascular"), ("D", "Dermatology"),
    ("G04", "Renal / Urology"), ("G", "Reproductive & sexual health"),
    ("H", "Endocrine / Metabolic"), ("J", "Infectious disease"), ("P", "Infectious disease"),
    ("M", "Musculoskeletal"), ("N01", "Pain & inflammation"), ("N02", "Pain & inflammation"),
    ("N05", "Mental health"), ("N06", "Mental health"), ("N07B", "Mental health"),
    ("N", "Neurology"), ("R", "Respiratory"), ("S01", "Ophthalmology"), ("S", "Other"),
    ("V", "Other"),
]
_ATC_CODE = re.compile(r"\b([A-Z]\d{2}[A-Z]{0,2}\d{0,2})\b")
def _english_area(text):
    first = re.split(r"[\s/,(]+", _fold(text), maxsplit=1)[0]
    return _ENGLISH_AREA_EN.get(first)
def therapy_areas(agency, raw_area, atc=None):
    """
    (primary area, every area) in the shared English vocabulary.

    G-BA: the trailing parenthetical Therapiegebiet — 'Melanom (onkologische
    Erkrankungen)' → Oncology. HAS: each 'System/Specific' pair, with any pair
    that names a cancer read as Oncology; Oncology is the primary area whenever
    it's among them. No category at all → the ATC code, if any.
    """
    text = _text(raw_area)
    areas = []
    if text and agency == "G-BA":
        m = re.search(r"\(([^()]+)\)\s*$", text)
        group = _fold(m.group(1)) if m else _fold(text)
        area = _GBA_AREA_EN.get(group) or _english_area(text)
        areas.append(area or "Other")
    elif text and agency == "HAS":
        for pair in text.split(";"):
            folded = _fold(pair).strip()
            if not folded:
                continue
            system = folded.split("/")[0].strip()
            if "cancer" in folded:
                area = "Oncology"
            else:
                area = _HAS_SYSTEM_EN.get(system) or _english_area(pair) or "Other"
            if area not in areas:
                areas.append(area)
    elif text:
        areas.append(_english_area(text) or "Other")
    if not areas and atc:
        m = _ATC_CODE.search(str(atc).upper())
        if m:
            code = m.group(1)
            areas.append(next((a for p, a in _ATC_AREA_RULES if code.startswith(p)), "Other"))
    if not areas:
        return "Not classified", ("Not classified",)
    primary = "Oncology" if "Oncology" in areas else areas[0]
    return primary, tuple(areas)
def has_procedure_group(value):
    """HAS records ~720 different free-text procedure names; these are the eight
    kinds of procedure behind them."""
    t = _fold(value)
    if not t:
        return "Not recorded"
    if "acces precoce" in t or "early access" in t:
        return "Early access"
    if re.search(r"extensions? d.?indications?", t):
        return "Extension of indication"
    if "renouvellement" in t[:60]:
        return "Renewal"
    if "reevaluation" in t or "nouvel examen" in t or "reexamen" in t:
        return "Re-evaluation"
    if "radiation" in t:
        return "Delisting"
    if "modification" in t:
        return "Label or listing change"
    if "inscription" in t:
        return "First listing"
    return "Other"
def gba_procedure_group(row):
    t = _fold(row.get("indication"))
    if re.search(r"neubewertung|erneute (?:nutzen)?bewertung|fristablauf"
                 r"|neuer wissenschaftlicher erkenntnisse|uberschreitung", t):
        return "Re-assessment"
    if "neues anwendungsgebiet" in t:
        return "New indication"
    return "Benefit assessment"
def condition_label(agency, row):
    """A short name for the condition, for tables. G-BA: the disease in its
    therapy-area tag ('Mammakarzinom'); HAS: the condition in the opinion title
    ('COLUMVI (glofitamab) - Lymphome diffus à grande cellule B'), else the
    specific part of its category."""
    if agency == "G-BA":
        area = _text(row.get("therapeutic_area")) or ""
        if re.search(r"\([^()]+\)\s*$", area):
            return re.sub(r"\s*\([^()]*\)\s*$", "", area).strip() or area
        # No German tag (early English rows): the indication, trimmed.
        return _shorten(_text(row.get("indication")) or area, 80)
    title = _text(row.get("source_title")) or ""
    m = re.search(r"\)\s+-\s+(.+)$", title)
    if m:
        return m.group(1).strip()
    area = _text(row.get("therapeutic_area"))
    if area:
        first = area.split(";")[0]
        return first.split("/")[-1].strip()
    return ""
def orphan_flag(value):
    tokens = set(re.findall(r"[a-z0-9]+", _fold(value)))
    if tokens & {"ja", "yes", "1", "true"}:
        return "Yes"
    if tokens & {"nein", "no", "0", "false"}:
        return "No"
    return None
CONDITION_TERMS = [
    # ── Oncology: specific entries before the family they belong to
    {"label": "Triple-negative breast cancer",
     "triggers": ["tnbc", "triple negative breast", "triple-negative breast"],
     "en": ["triple-negative", "triple negative", "tnbc"],
     "de": ["triple-negativ", "tnbc"], "fr": ["triple negatif", "triple-negatif", "tnbc"]},
    {"label": "Breast cancer",
     "triggers": ["breast cancer", "breast carcinoma", "breast", "mbc", "her2-positive breast",
                  "her2 positive breast", "hr-positive breast"],
     "en": ["breast cancer", "breast carcinoma"], "de": ["mammakarzinom", "brustkrebs"],
     "fr": ["cancer du sein", "cancers du sein", "carcinome mammaire", "cancer mammaire"]},
    {"label": "Non-small cell lung cancer",
     "triggers": ["nsclc", "non-small cell lung cancer", "non small cell lung cancer",
                  "non-small-cell lung cancer", "non-small cell", "non small cell"],
     "en": ["nsclc", "non-small cell lung", "non-small-cell lung", "non small cell lung"],
     "de": ["nsclc", "nicht-kleinzellig", "nicht kleinzellig"],
     "fr": ["cbnpc", "non a petites cellules"]},
    {"label": "Small cell lung cancer",
     "triggers": ["sclc", "small cell lung cancer", "small-cell lung cancer", "small cell lung"],
     "en": ["sclc", r"re:(?<!non-)(?<!non )small[- ]cell lung"],
     "de": ["sclc", r"re:(?<!nicht-)(?<!nicht )kleinzellig"],
     "fr": ["cbpc", r"re:(?<!non )a petites cellules"]},
    {"label": "Lung cancer",
     "triggers": ["lung cancer", "lung carcinoma", "lung tumour", "lung tumor"],
     "en": ["lung cancer", "lung carcinoma", "nsclc", "sclc"],
     "de": ["lungenkarzinom", "bronchialkarzinom", "lungenkrebs", "nsclc", "sclc"],
     "fr": ["cancer bronchique", "cancer du poumon", "cancer pulmonaire", "cbnpc", "cbpc",
            "cancer de l'appareil respiratoire"]},
    {"label": "Multiple myeloma", "triggers": ["myeloma", "multiple myeloma", "mm"],
     "en": ["myeloma"], "de": ["myelom"], "fr": ["myelome"]},
    {"label": "Melanoma", "triggers": ["melanoma"], "en": ["melanoma"], "de": ["melanom"],
     "fr": ["melanome"]},
    {"label": "Prostate cancer",
     "triggers": ["prostate cancer", "prostate", "mcrpc", "crpc", "mhspc", "nmcrpc"],
     "en": ["prostate cancer", "prostate carcinoma", "mcrpc", "crpc", "mhspc"],
     "de": ["prostatakarzinom", "prostatakrebs"],
     "fr": ["cancer de la prostate", "cancers de la prostate", "adenocarcinome de la prostate",
            "carcinome de la prostate"]},
    {"label": "Colorectal cancer",
     "triggers": ["colorectal", "colorectal cancer", "bowel cancer", "colon cancer",
                  "rectal cancer", "crc", "mcrc"],
     "en": ["colorectal", "colon cancer", "rectal cancer", "mcrc"],
     "de": ["kolorektal", "kolonkarzinom", "rektumkarzinom", "darmkrebs"],
     "fr": ["colorectal", "cancer du colon", "cancer du rectum"]},
    {"label": "Renal cell carcinoma",
     "triggers": ["renal cell", "renal cell carcinoma", "kidney cancer", "rcc"],
     "en": ["renal cell", "rcc"], "de": ["nierenzellkarzinom"],
     "fr": ["cancer du rein", "carcinome renal", "carcinome a cellules renales"]},
    {"label": "Urothelial / bladder cancer",
     "triggers": ["urothelial", "urothelial carcinoma", "bladder cancer"],
     "en": ["urothelial", "bladder cancer"],
     "de": ["urothelkarzinom", "harnblasenkarzinom", "blasenkarzinom"],
     "fr": ["urothelial", "cancer de la vessie"]},
    {"label": "Ovarian cancer", "triggers": ["ovarian", "ovarian cancer"],
     "en": ["ovarian cancer", "ovarian carcinoma"], "de": ["ovarialkarzinom", "eierstockkrebs"],
     "fr": ["cancer de l'ovaire", "cancers de l'ovaire", "carcinome de l'ovaire",
            "cancer ovarien", "carcinome ovarien"]},
    {"label": "Endometrial cancer", "triggers": ["endometrial", "endometrial cancer"],
     "en": ["endometrial cancer", "endometrial carcinoma"], "de": ["endometriumkarzinom"],
     "fr": ["cancer de l'endometre", "carcinome de l'endometre"]},
    {"label": "Cervical cancer", "triggers": ["cervical cancer"], "en": ["cervical cancer"],
     "de": ["zervixkarzinom"], "fr": ["cancer du col de l'uterus", "col de l'uterus"]},
    {"label": "Gastric / oesophageal cancer",
     "triggers": ["gastric", "gastric cancer", "stomach cancer", "gastro-oesophageal",
                  "gastroesophageal", "gej", "oesophageal", "esophageal", "oesophageal cancer",
                  "esophageal cancer"],
     "en": ["gastric cancer", "gastric adenocarcinoma", "stomach cancer", "oesophageal cancer",
            "esophageal cancer", "gastroesophageal junction", "gastro-oesophageal junction"],
     "de": ["magenkarzinom", "adenokarzinom des magens", "karzinom des magens",
            "osophaguskarzinom", "karzinom des osophagus", "plattenepithelkarzinom des osophagus",
            "gastroosophagealen ubergangs"],
     "fr": ["cancer de l'estomac", "cancer gastrique", "adenocarcinome gastrique",
            "cancer de l'oesophage", "carcinome epidermoide de l'oesophage",
            "jonction oesogastrique", "jonction oeso-gastrique"]},
    {"label": "Liver cancer (HCC)",
     "triggers": ["hepatocellular", "hepatocellular carcinoma", "hcc", "liver cancer"],
     "en": ["hepatocellular", "hcc", "liver cancer"],
     "de": ["hepatozellular", "leberzellkarzinom"], "fr": ["hepatocellulaire", "chc"]},
    {"label": "Primary biliary cholangitis",
     "triggers": ["primary biliary cholangitis", "primary biliary cirrhosis", "pbc"],
     "en": ["primary biliary cholangitis", "primary biliary cirrhosis", "pbc"],
     "de": ["primar biliare cholangitis", "primare biliare cholangitis", "pbc"],
     "fr": ["cholangite biliaire primitive", "cirrhose biliaire primitive", "cbp"]},
    {"label": "Biliary tract cancer",
     "triggers": ["biliary", "biliary tract cancer", "biliary cancer", "cholangiocarcinoma",
                  "bile duct cancer"],
     "en": ["biliary tract cancer", "biliary cancer", "cholangiocarcinoma", "bile duct cancer"],
     "de": ["cholangiokarzinom", "gallengangskarzinom", "biliares karzinom",
            "karzinom der gallenwege", "gallenwegskarzinom"],
     "fr": ["cancer des voies biliaires", "cancers des voies biliaires", "cholangiocarcinome",
            "carcinome des voies biliaires"]},
    {"label": "Pancreatic cancer", "triggers": ["pancreatic", "pancreatic cancer"],
     "en": ["pancreatic cancer", "pancreatic adenocarcinoma", "pancreatic ductal"],
     "de": ["pankreaskarzinom", "adenokarzinom des pankreas", "karzinom des pankreas"],
     "fr": ["cancer du pancreas", "adenocarcinome du pancreas", "adenocarcinome pancreatique",
            "adenocarcinome canalaire pancreatique"]},
    {"label": "Head and neck cancer",
     "triggers": ["head and neck", "head and neck cancer", "hnscc", "scchn"],
     "en": ["head and neck", "hnscc", "scchn"], "de": ["kopf-hals", "kopf- und hals"],
     "fr": ["tete et du cou", "voies aerodigestives superieures"]},
    {"label": "Thyroid cancer", "triggers": ["thyroid cancer"],
     "en": ["thyroid cancer", "thyroid carcinoma"],
     "de": ["schilddrusenkarzinom", "schilddrusenneoplasie"],
     "fr": ["cancer de la thyroide", "carcinome thyroidien", "cancer thyroidien"]},
    {"label": "Brain tumours",
     "triggers": ["glioma", "glioblastoma", "brain tumour", "brain tumor"],
     "en": ["glioma", "glioblastoma", "brain tumour", "brain tumor"],
     "de": ["gliom", "glioblastom"], "fr": ["gliome", "glioblastome"]},
    {"label": "Chronic lymphocytic leukaemia",
     "triggers": ["cll", "chronic lymphocytic", "chronic lymphocytic leukaemia",
                  "chronic lymphocytic leukemia"],
     "en": ["chronic lymphocytic", "cll"],
     "de": ["chronische lymphatische leukamie", "chronisch lymphatische leukamie", "cll"],
     "fr": ["leucemie lymphoide chronique", "llc"]},
    {"label": "Acute myeloid leukaemia",
     "triggers": ["aml", "acute myeloid", "acute myeloid leukaemia", "acute myeloid leukemia"],
     "en": ["acute myeloid", "aml"],
     "de": ["akute myeloische leukamie", "leukamie, myeloische, akute", "aml"],
     "fr": ["leucemie aigue myeloide", "lam"]},
    {"label": "Chronic myeloid leukaemia",
     "triggers": ["cml", "chronic myeloid", "chronic myeloid leukaemia",
                  "chronic myeloid leukemia"],
     "en": ["chronic myeloid", "cml"], "de": ["chronische myeloische leukamie", "cml"],
     "fr": ["leucemie myeloide chronique", "lmc"]},
    {"label": "Acute lymphoblastic leukaemia",
     "triggers": ["acute lymphoblastic", "acute lymphocytic", "acute lymphoblastic leukaemia",
                  "acute lymphoblastic leukemia", "b-all", "t-all"],
     "en": ["acute lymphoblastic", "acute lymphocytic"],
     "de": ["akute lymphatische leukamie", "akute lymphoblastische leukamie",
            "lymphatische b-zell-leukamie", "b-zell-vorlaufer"],
     "fr": ["leucemie aigue lymphoblastique", "lal"]},
    {"label": "Leukaemia", "triggers": ["leukaemia", "leukemia"],
     "en": ["leukaemia", "leukemia"], "de": ["leukamie"], "fr": ["leucemie"]},
    {"label": "Diffuse large B-cell lymphoma",
     "triggers": ["dlbcl", "diffuse large b", "diffuse large b-cell lymphoma",
                  "diffuse large b cell lymphoma"],
     "en": ["diffuse large b", "dlbcl"], "de": ["grosszellig", "dlbcl"],
     "fr": ["grandes cellules b", "grande cellule b", "ldgcb"]},
    {"label": "Follicular lymphoma", "triggers": ["follicular", "follicular lymphoma", "fl"],
     "en": ["follicular lymphoma"], "de": ["follikulares lymphom", "follikulares b-zell"],
     "fr": ["lymphome folliculaire"]},
    {"label": "Mantle cell lymphoma",
     "triggers": ["mantle cell", "mantle cell lymphoma", "mcl"],
     "en": ["mantle cell"], "de": ["mantelzell"], "fr": ["cellules du manteau"]},
    {"label": "Non-Hodgkin lymphoma", "triggers": ["non-hodgkin", "non hodgkin", "nhl"],
     "en": ["non-hodgkin", "non hodgkin"], "de": ["non-hodgkin"],
     "fr": ["non hodgkinien", "non-hodgkinien", "lnh"]},
    {"label": "Hodgkin lymphoma",
     "triggers": ["hodgkin", "hodgkin lymphoma", "hodgkin's lymphoma", "classical hodgkin"],
     "en": [r"re:(?<!non-)(?<!non )hodgkin"], "de": [r"re:(?<!non-)(?<!non )hodgkin"],
     "fr": [r"re:(?<!non-)(?<!non )hodgkin"]},
    {"label": "Lymphoma", "triggers": ["lymphoma"], "en": ["lymphoma"], "de": ["lymphom"],
     "fr": ["lymphome"]},
    {"label": "Sarcoma / GIST", "triggers": ["sarcoma", "gist"], "en": ["sarcoma", "gist"],
     "de": ["sarkom", "stromatumor"], "fr": ["sarcome", "tumeurs stromales"]},
    {"label": "Solid tumours (tumour-agnostic)",
     "triggers": ["solid tumour", "solid tumor", "solid tumours", "solid tumors", "ntrk"],
     "en": ["solid tumour", "solid tumor", "ntrk"], "de": ["solide tumor", "ntrk"],
     "fr": ["tumeurs solides", "ntrk"]},
    {"label": "Bone metastases",
     "triggers": ["bone metastases", "bone metastasis", "skeletal-related"],
     "en": ["bone metastases", "skeletal-related"],
     "de": ["knochenmetastasen", "skelettbezogen"], "fr": ["metastases osseuses"]},
    # ── Metabolic, infectious, neurological, immune and other
    {"label": "Type 1 diabetes",
     "triggers": ["type 1 diabetes", "type 1 diabetes mellitus", "t1d", "t1dm"],
     "en": ["type 1 diabetes"],
     "de": ["diabetes mellitus typ 1", "typ-1-diabetes", "typ 1 diabetes"],
     "fr": ["diabete de type 1"]},
    {"label": "Type 2 diabetes",
     "triggers": ["type 2 diabetes", "type 2 diabetes mellitus", "t2dm", "t2d"],
     "en": ["type 2 diabetes"],
     "de": ["diabetes mellitus typ 2", "typ-2-diabetes", "typ 2 diabetes"],
     "fr": ["diabete de type 2"]},
    {"label": "Diabetes", "triggers": ["diabetes"], "en": ["diabetes"], "de": ["diabetes"],
     "fr": ["diabete"]},
    {"label": "Obesity", "triggers": ["obesity", "weight management", "overweight"],
     "en": ["obesity", "weight management"], "de": ["adipositas", "ubergewicht"],
     "fr": ["obesite", "surpoids"]},
    {"label": "Cystic fibrosis", "triggers": ["cystic fibrosis", "cf"],
     "en": ["cystic fibrosis"], "de": ["zystische fibrose", "mukoviszidose"],
     "fr": ["mucoviscidose"]},
    {"label": "HIV", "triggers": ["hiv"], "en": ["hiv"], "de": ["hiv"], "fr": ["vih", "hiv"]},
    {"label": "Hepatitis C", "triggers": ["hepatitis c", "hcv"], "en": ["hepatitis c", "hcv"],
     "de": ["hepatitis c", "hcv"], "fr": ["hepatite c", "vhc"]},
    {"label": "Hepatitis B / D",
     "triggers": ["hepatitis b", "hepatitis d", "hepatitis delta", "hbv", "hdv"],
     "en": ["hepatitis b", "hepatitis d", "hepatitis delta", "hbv", "hdv"],
     "de": ["hepatitis b", "hepatitis d", "hepatitis delta"],
     "fr": ["hepatite b", "hepatite d", "hepatite delta", "vhb"]},
    {"label": "COVID-19", "triggers": ["covid", "covid-19", "sars-cov-2"],
     "en": ["covid", "sars-cov-2"], "de": ["covid", "coronavirus"], "fr": ["covid"]},
    {"label": "RSV", "triggers": ["rsv", "respiratory syncytial"],
     "en": ["rsv", "respiratory syncytial"], "de": ["rsv", "synzytial"],
     "fr": ["vrs", "virus respiratoire syncytial"]},
    {"label": "Bacterial infections",
     "triggers": ["bacterial infection", "bacterial infections", "antibiotic", "antibiotics",
                  "pneumonia"],
     "en": ["bacterial infection", "antibiotic", "pneumonia"],
     "de": ["bakterielle infektion", "pneumonie", "antibiot"],
     "fr": ["infection bacterienne", "infections bacteriennes", "antibiot", "pneumonie"]},
    {"label": "Vaccines",
     "triggers": ["vaccine", "vaccines", "vaccination", "immunisation", "immunization"],
     "en": ["vaccine", "vaccination"], "de": ["impfstoff", "impfung", "immunisierung"],
     "fr": ["vaccin"]},
    {"label": "Multiple sclerosis",
     "triggers": ["multiple sclerosis", "relapsing multiple sclerosis", "ms", "rrms"],
     "en": ["multiple sclerosis", "rrms"], "de": ["multiple sklerose", "multipler sklerose"],
     "fr": ["sclerose en plaques"]},
    {"label": "Epilepsy", "triggers": ["epilepsy", "seizure", "seizures", "dravet", "lennox"],
     "en": ["epilepsy", "seizure", "dravet", "lennox"],
     "de": ["epilepsie", "anfall", "anfalle", "dravet", "lennox"],
     "fr": ["epilepsie", "dravet", "lennox"]},
    {"label": "Migraine", "triggers": ["migraine"], "en": ["migraine"], "de": ["migrane"],
     "fr": ["migraine"]},
    {"label": "Alzheimer's / dementia", "triggers": ["alzheimer", "alzheimer's", "dementia"],
     "en": ["alzheimer", "dementia"], "de": ["alzheimer", "demenz"],
     "fr": ["alzheimer", "demence"]},
    {"label": "Parkinson's disease", "triggers": ["parkinson", "parkinson's"],
     "en": ["parkinson"], "de": ["parkinson"], "fr": ["parkinson"]},
    {"label": "Spinal muscular atrophy", "triggers": ["spinal muscular atrophy", "sma"],
     "en": ["spinal muscular atrophy"],
     "de": ["spinale muskelatrophie", "spinaler muskelatrophie"],
     "fr": ["amyotrophie spinale"]},
    {"label": "Duchenne muscular dystrophy", "triggers": ["duchenne", "dmd"],
     "en": ["duchenne"], "de": ["duchenne"], "fr": ["duchenne"]},
    {"label": "Myasthenia gravis", "triggers": ["myasthenia", "myasthenia gravis"],
     "en": ["myasthenia"], "de": ["myasthenia", "myasthenie"], "fr": ["myasthenie"]},
    {"label": "Depression",
     "triggers": ["depression", "depressive", "major depressive disorder", "mdd"],
     "en": ["depression", "depressive"], "de": ["depression", "depressive"],
     "fr": ["depression", "depressif"]},
    {"label": "Schizophrenia", "triggers": ["schizophrenia"], "en": ["schizophrenia"],
     "de": ["schizophrenie"], "fr": ["schizophrenie"]},
    {"label": "Psoriasis / psoriatic arthritis",
     "triggers": ["psoriasis", "psoriatic", "psoriatic arthritis", "plaque psoriasis"],
     "en": ["psoriasis", "psoriatic"], "de": ["psoriasis"],
     "fr": ["psoriasis", "psoriasique"]},
    {"label": "Atopic dermatitis", "triggers": ["atopic dermatitis", "eczema"],
     "en": ["atopic dermatitis", "eczema"],
     "de": ["atopische dermatitis", "neurodermitis", "ekzem"],
     "fr": ["dermatite atopique", "eczema"]},
    {"label": "Rheumatoid arthritis",
     "triggers": ["rheumatoid", "rheumatoid arthritis", "ra"],
     "en": ["rheumatoid"], "de": ["rheumatoide arthritis", "rheumatoider arthritis"],
     "fr": ["polyarthrite rhumatoide"]},
    {"label": "Axial spondyloarthritis",
     "triggers": ["ankylosing spondylitis", "axial spondyloarthritis", "axspa"],
     "en": ["ankylosing spondylitis", "spondyloarthritis", "axspa"],
     "de": ["spondylitis", "spondyloarthritis"], "fr": ["spondylarthrite", "spondyloarthrite"]},
    {"label": "Ulcerative colitis", "triggers": ["ulcerative colitis", "uc"],
     "en": ["ulcerative colitis"], "de": ["colitis ulcerosa"],
     "fr": ["rectocolite hemorragique", "rch"]},
    {"label": "Crohn's disease", "triggers": ["crohn", "crohn's", "crohn's disease"],
     "en": ["crohn"], "de": ["crohn"], "fr": ["crohn"]},
    {"label": "Inflammatory bowel disease", "triggers": ["ibd", "inflammatory bowel"],
     "en": ["inflammatory bowel", "ulcerative colitis", "crohn"],
     "de": ["colitis ulcerosa", "crohn"], "fr": ["rectocolite hemorragique", "crohn"]},
    {"label": "Asthma", "triggers": ["asthma"], "en": ["asthma"], "de": ["asthma"],
     "fr": ["asthme"]},
    {"label": "COPD", "triggers": ["copd"], "en": ["copd"],
     "de": ["copd", "chronisch obstruktive"],
     "fr": ["bpco", "bronchopneumopathie chronique", "broncho-pneumopathie chronique"]},
    {"label": "Idiopathic pulmonary fibrosis",
     "triggers": ["pulmonary fibrosis", "idiopathic pulmonary fibrosis", "ipf"],
     "en": ["pulmonary fibrosis"], "de": ["lungenfibrose"], "fr": ["fibrose pulmonaire"]},
    {"label": "Pulmonary hypertension",
     "triggers": ["pulmonary hypertension", "pulmonary arterial hypertension", "pah"],
     "en": ["pulmonary hypertension", "pulmonary arterial hypertension", "pah"],
     "de": ["pulmonale hypertonie", "pulmonalen hypertonie", "pulmonal arterielle hypertonie",
            "pulmonal arteriellen hypertonie"],
     "fr": ["hypertension arterielle pulmonaire", "hypertension pulmonaire", "htap"]},
    {"label": "Heart failure", "triggers": ["heart failure", "hf", "hfref", "hfpef"],
     "en": ["heart failure"], "de": ["herzinsuffizienz"], "fr": ["insuffisance cardiaque"]},
    {"label": "Hypercholesterolaemia",
     "triggers": ["hypercholesterolaemia", "hypercholesterolemia", "cholesterol",
                  "dyslipidaemia", "dyslipidemia", "ldl"],
     "en": ["hypercholesterolaemia", "hypercholesterolemia", "dyslipidaemia", "dyslipidemia"],
     "de": ["hypercholesterinamie", "dyslipidamie", "cholesterin"],
     "fr": ["hypercholesterolemie", "dyslipidemie", "cholesterol"]},
    {"label": "Hypertension",
     "triggers": ["hypertension", "high blood pressure", "arterial hypertension",
                  "essential hypertension"],
     # English puts the qualifier first ('pulmonary hypertension'); French, whose
     # 'hypertension' this also matches, puts it after ('hypertension oculaire').
     "en": [r"re:(?<!pulmonary arterial )(?<!pulmonary )(?<!ocular )(?<!portal )"
            r"(?<!intracranial )hypertension(?! arterielle pulmonaire)(?! pulmonaire)"
            r"(?! oculaire)(?! intraoculaire)(?! intra-oculaire)(?! portale)"
            r"(?! intracranienne)(?! arterielle pulmonaire)"],
     "de": [r"re:(?<!pulmonal )(?<!pulmonale )arterielle hypertonie",
            "essentielle hypertonie", "bluthochdruck"],
     "fr": [r"re:hypertension arterielle(?! pulmonaire)", "hta"]},
    {"label": "Thrombosis / anticoagulation",
     "triggers": ["thrombosis", "vte", "dvt", "pulmonary embolism", "anticoagulant",
                  "anticoagulation"],
     "en": ["thrombosis", "vte", "dvt", "pulmonary embolism", "anticoagula"],
     "de": ["thrombose", "embolie"], "fr": ["thrombose", "embolie", "mtev"]},
    {"label": "Atrial fibrillation / stroke", "triggers": ["atrial fibrillation", "af", "stroke"],
     "en": ["atrial fibrillation", "stroke"], "de": ["vorhofflimmern", "schlaganfall"],
     "fr": ["fibrillation atriale", "fibrillation auriculaire",
            "accident vasculaire cerebral", "avc"]},
    {"label": "Haemophilia", "triggers": ["haemophilia", "hemophilia"],
     "en": ["haemophilia", "hemophilia"], "de": ["hamophilie"], "fr": ["hemophilie"]},
    {"label": "Sickle cell / thalassaemia",
     "triggers": ["sickle cell", "sickle cell disease", "thalassaemia", "thalassemia"],
     "en": ["sickle cell", "thalassaemia", "thalassemia"],
     "de": ["sichelzell", "thalassamie"], "fr": ["drepanocytose", "thalassemie"]},
    {"label": "PNH", "triggers": ["pnh", "paroxysmal nocturnal"],
     "en": ["pnh", "paroxysmal nocturnal"], "de": ["hamoglobinurie", "pnh"],
     "fr": ["hemoglobinurie", "hpn"]},
    {"label": "Amyloidosis", "triggers": ["amyloidosis", "attr"], "en": ["amyloidosis", "attr"],
     "de": ["amyloidose"], "fr": ["amylose"]},
    {"label": "Macular degeneration / oedema", "triggers": ["macular", "amd", "dme"],
     "en": ["macular"], "de": ["makuladegeneration", "makulaodem", "makula"],
     "fr": ["degenerescence maculaire", "dmla", "oedeme maculaire", "maculaire"]},
    {"label": "Hereditary angioedema", "triggers": ["angioedema", "hae"],
     "en": ["angioedema"], "de": ["angioodem"], "fr": ["angioedeme", "angio-oedeme"]},
    {"label": "Lysosomal storage disorders",
     "triggers": ["fabry", "pompe", "gaucher", "mucopolysaccharidosis", "mps"],
     "en": ["fabry", "pompe", "gaucher", "mucopolysaccharidosis"],
     "de": ["fabry", "pompe", "gaucher", "mukopolysaccharidose", "mucopolysaccharidose"],
     "fr": ["fabry", "pompe", "gaucher", "mucopolysaccharidose"]},
    {"label": "Chronic kidney disease",
     "triggers": ["chronic kidney disease", "ckd", "kidney disease"],
     "en": ["chronic kidney disease", "ckd"],
     "de": ["chronische nierenerkrankung", "niereninsuffizienz"],
     "fr": ["maladie renale chronique", "insuffisance renale"]},
    {"label": "Osteoporosis", "triggers": ["osteoporosis"], "en": ["osteoporosis"],
     "de": ["osteoporose"], "fr": ["osteoporose"]},
    {"label": "Urticaria", "triggers": ["urticaria"], "en": ["urticaria"], "de": ["urtikaria"],
     "fr": ["urticaire"]},
    {"label": "Lupus", "triggers": ["lupus", "sle"], "en": ["lupus"], "de": ["lupus"],
     "fr": ["lupus"]},
    {"label": "Graft-versus-host disease",
     "triggers": ["graft-versus-host", "graft versus host", "gvhd"],
     "en": ["graft-versus-host", "graft versus host", "gvhd"],
     "de": ["graft-versus-host", "gvhd"],
     "fr": ["greffon contre l'hote", "maladie du greffon", "gvh"]},
    {"label": "Transplantation", "triggers": ["transplant", "transplantation"],
     "en": ["transplant"], "de": ["transplantation"], "fr": ["greffe", "transplantation"]},
    {"label": "Pain", "triggers": ["pain", "analgesic"], "en": ["pain", "analgesic"],
     "de": ["schmerz"], "fr": ["douleur", "antalgique"]},
    {"label": "Neurofibromatosis", "triggers": ["neurofibromatosis", "nf1"],
     "en": ["neurofibromatosis", "nf1"], "de": ["neurofibromatose"],
     "fr": ["neurofibromatose"]},
    {"label": "Glaucoma", "triggers": ["glaucoma"], "en": ["glaucoma"], "de": ["glaukom"],
     "fr": ["glaucome"]},
    {"label": "Insomnia / sleep", "triggers": ["insomnia", "sleep disorder", "narcolepsy"],
     "en": ["insomnia", "sleep disorder", "narcolepsy"],
     "de": ["schlafstorung", "insomnie", "narkolepsie"],
     "fr": ["insomnie", "troubles du sommeil", "narcolepsie"]},
    {"label": "Contraception", "triggers": ["contraception", "contraceptive"],
     "en": ["contracepti"], "de": ["kontrazeption", "verhutung"],
     "fr": ["contraception", "contraceptif"]},
    {"label": "Growth hormone deficiency", "triggers": ["growth hormone"],
     "en": ["growth hormone"], "de": ["wachstumshormon"], "fr": ["hormone de croissance"]},
]
_QUERY_STOPWORDS = {
    "and", "or", "the", "for", "with", "without", "in", "of", "to", "a", "an", "adult", "adults",
    "patient", "patients", "treatment", "therapy", "advanced", "metastatic", "metastasis",
    "metastases", "locally", "unresectable", "relapsed", "refractory", "first", "second",
    "third", "line", "chronic", "acute", "severe", "moderate", "disease", "disorder",
    "syndrome", "cancer", "carcinoma", "tumour", "tumor", "positive", "negative", "mutation",
    "mutated", "type", "stage", "early", "previously", "untreated", "treated", "children",
    "child", "paediatric", "pediatric", "infection", "deficiency", "condition", "primary",
    "secondary", "maintenance", "adjuvant", "neoadjuvant", "combination", "monotherapy",
}
def _is_stopword(word):
    return (word in _QUERY_STOPWORDS
            or (word.endswith("s") and word[:-1] in _QUERY_STOPWORDS)
            or (word.endswith("es") and word[:-2] in _QUERY_STOPWORDS))
def _term_body(term):
    """A term that must start a word; very short ones (abbreviations) must also
    end it. 're:' terms are regular expressions, used as written."""
    if term.startswith("re:"):
        return term[3:]
    return re.escape(term) + (r"(?![a-z0-9])" if len(term) <= 4 else "")
def _term_pattern(term):
    return rf"(?<![a-z0-9]){_term_body(term)}"
def _any_term_regex(terms):
    """One pattern for all terms, with the shared word-start check hoisted out of
    the alternation — several times faster over 12k HAS records than repeating
    it per term."""
    if not terms:
        return None
    ordered = sorted(set(terms), key=len, reverse=True)
    return r"(?<![a-z0-9])(?:" + "|".join(_term_body(t) for t in ordered) + ")"
_ALL_ENTRY = next(e for e in CONDITION_TERMS if e["label"] == "Acute lymphoblastic leukaemia")
def expand_condition_query(agency, query):
    """
    ([term groups], [condition labels], [ranking words]) for a condition search.

    Each recognised condition is one group of terms (any one matches), and a
    record must match every group — 'HIV hepatitis' means both. Qualifiers and
    words the dictionary doesn't know only rank, since English words can't be
    expected in German or French text. When nothing is recognised, each
    distinctive word must match, cut to a stem so English Latin roots still
    meet their German and French cousins: 'myelodysplastic' → 'myelodyspla'
    finds 'myelodysplastische' and 'myélodysplasiques'. Abbreviations that are
    ambiguous in running text (UC, RA, MS, MM, …) are listed as triggers only,
    never as search terms.
    """
    raw = _text(query) or ""
    q = re.sub(r"\s+", " ", _fold(raw)).strip()
    if not q:
        return [], [], []
    lang = "de" if agency == "G-BA" else "fr"
    remaining = f" {q} "
    groups, labels = [], []

    def take(entry):
        if entry["label"] not in labels:
            labels.append(entry["label"])
            groups.append([t for t in dict.fromkeys(entry["en"] + entry[lang])])

    # 'ALL' in capitals is acute lymphoblastic leukaemia; 'all' is just a word.
    if re.search(r"(?<![A-Za-z])ALL(?![A-Za-z])", raw):
        take(_ALL_ENTRY)
        remaining = re.sub(r"(?<![a-z0-9])all(?![a-z0-9])", " ", remaining)
    triggers = sorted(((trigger, entry) for entry in CONDITION_TERMS
                       for trigger in entry["triggers"] + entry["en"] + entry["de"] + entry["fr"]
                       if not trigger.startswith("re:")),
                      key=lambda pair: -len(pair[0]))
    for trigger, entry in triggers:
        pattern = rf"(?<![a-z0-9]){re.escape(trigger)}(?![a-z0-9])"
        if re.search(pattern, remaining):
            remaining = re.sub(pattern, " ", remaining)
            take(entry)
    words = [w for w in re.findall(r"[a-z0-9][a-z0-9'-]*", remaining) if len(w) >= 2]
    distinct = [w if len(w) <= 5 else w[:max(6, len(w) - 4)] for w in words
                if not _is_stopword(w)]
    if groups:
        return groups, labels, distinct
    return [[w] for w in distinct], labels, []
_MOLECULE_SEPARATORS = re.compile(
    r"\s*(?:\bin combination with\b|\bcombined with\b|\bfollowed by\b|\bwith or without\b"
    r"|\bwithout\b|\bwith\b|\bplus\b|\band\b|\bor\b|\bthen\b|\+|,|;|/|&|–|—|\s-\s)\s*",
    re.IGNORECASE)
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
        # Long dashes join regimens ('Nivolumab–relatlimab'); read them as
        # separators before folding turns them into plain hyphens.
        s = _fold((_text(name) or "").replace("–", ";").replace("—", ";"))
        if not s:
            continue
        # Accents folded first ('lénalidomide' was read as 'l' + 'nalidomide'),
        # and a line break separates substances like any other separator.
        s = re.sub(r"\([^)]*\)|\([^)]*$", " ", s).replace("\n", ";").replace("\r", ";")
        for part in _MOLECULE_SEPARATORS.split(s):
            part = _DESCRIPTIVE_TAIL.split(part, maxsplit=1)[0]
            words = [w for w in re.findall(r"[a-z]+", part)
                     if w not in _SALT_WORDS and w not in _DESCRIPTOR_WORDS]
            key = "".join(words)
            if len(key) >= 4 and " ".join(words) not in _NON_MOLECULES:
                keys.add(_inn_stem(key))
    return frozenset(keys)
def _inn_stem(key):
    """German and French INNs differ from English by a final 'e' — Enzalutamid /
    enzalutamide, céfazoline / cefazolin — so keys drop it for matching."""
    return key[:-1] if len(key) > 5 and key.endswith("e") else key
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
_PRODUCTION_NOTE = re.compile(
    r"\s*\(\([^()]*\)\)|\s*\((?:mammif|humain|bact[eé]r|levure|v[eé]g[eé]tal|origine)[^()]*\)",
    re.IGNORECASE)
def _clean_inn(name):
    """HAS INNs can carry the production system in brackets — 'isatuximab
    ((MAMMIFERE/HAMSTER/CELLULES CHO))' — which is noise in a table."""
    if not name:
        return name
    return _PRODUCTION_NOTE.sub("", name).strip() or name
def format_eur_icer(row):
    lo, hi = row.get("icer_eur_lower"), row.get("icer_eur_upper")
    if lo is None or pd.isna(lo):
        return None
    if hi is None or pd.isna(hi) or float(hi) == float(lo):
        return f"€{float(lo):,.0f}/QALY"
    return f"€{float(lo):,.0f}–€{float(hi):,.0f}/QALY"
