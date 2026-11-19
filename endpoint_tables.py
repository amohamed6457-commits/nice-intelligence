"""
Sort recovered G-BA tables into the four endpoint categories.

Why
---
A record says "No usable data submitted" against quality of life and stops
there. A reader wants to see what WAS submitted. The recovered tables hold
it; nothing has ever grouped them by category.

How the tables are shaped
-------------------------
A G-BA endpoint table marks its sections with a spanning row -- first cell
filled, the rest empty:

    ['Mortality', '', '', '', '', '']
    ['Overall survival', '506', 'n.r. 113 (22.3)', '504', ..., '0.79 [0.62; 1.01]']

Sections nest, and the OUTER heading decides the category:

    ['Symptomatology', '', ...]        <- morbidity
    ['EORTC QLQ-C30', '', ...]         <- maps to nothing on its own
    ['Fatigue', '472', ...]            <- morbidity, inherited

That nesting is the whole reason the instrument cannot decide the category.
The QLQ-C30's symptom scales are filed under Symptomatology (morbidity)
while its functional scales are filed under quality of life -- the same
questionnaire in two categories. A keyword match on "QLQ-C30" would put
them in the same bucket and be wrong about half the time, so a sub-heading
that maps to nothing inherits the category above it instead.

What this does NOT do
---------------------
It does not decide whether the data was any good. A table appearing under
"No usable data submitted" is not a contradiction: G-BA often receives data
and rules it nicht bewertbar, usually because the study was single-arm or
the comparator arm was not the one it specified. The panel shows the
committee's wording beside the tables so the two read as one statement.

    python endpoint_tables.py --audit          # coverage across the file
    python endpoint_tables.py --appraisal D-1058
"""

import argparse
import json
import os
import re
import sys

TABLES_FILE = "gba_tables.json"

# Tables that are not endpoint data at all. A G-BA resolution carries cost
# and comparator tables in the same document, and their rows are drug names
# (Cisplatin, Carboplatin, Pemetrexed) under headings like "Zu bewertendes
# Arzneimittel". Classifying those into an endpoint category would put drug
# prices under "Adverse events". They are excluded outright: showing the
# wrong table under a verdict is worse than showing none.
NOT_ENDPOINT = re.compile(
    r"zu\s+bewertendes?\s+arzneimittel|medicinal\s+product\s+to\s+be\s+assessed|"
    r"zweckm[aä][sß]ige\s+vergleichstherapie|appropriate\s+comparator|"
    r"jahrestherapiekosten|behandlungsdauer|kosten\s+f[uü]r|"
    r"annual\s+treatment\s+costs|treatment\s+duration|"
    r"verbrauch|consumption|dosierung|dosage|"
    # Treatment-regimen tables: arms, phases and doses, whose rows are drug
    # names (Cisplatin, Carboplatin, Vinorelbine). Same problem as the cost
    # tables -- real content, but not endpoint results.
    r"\binduction\b|maintenance\s+(?:treatment|therapy)|"
    r"induktionstherapie|erhaltungstherapie|therapieschema|"
    r"\d+\s*mg\s*/\s*(?:kg|m)|\bmg/kg/d\b|regimen",
    re.IGNORECASE)

# Structural rows that label nothing: a repeated header, a totals line.
# A trailing single letter is a footnote marker ("Totald", "Endpointa"), so
# it is allowed for, or rows like "Totald" slip through as unmatched.
STRUCTURAL = re.compile(r"^(endpoint|endpunkt|total|gesamt|insgesamt|n|"
                        r"study|studie|meta-?analysis|metaanalyse|"
                        r"subgroup|subgruppe)[a-z]?[\s:]*$", re.IGNORECASE)

# International Nonproprietary Name suffixes. Listing individual drugs never
# converges -- Cisplatin, Carboplatin, Docetaxel, Pemetrexed, Vinorelbine,
# Bendamustine kept reappearing in the audit -- but INN stems are a closed
# vocabulary, so this generalises to drugs the data has not shown yet.
DRUG_SUFFIX = re.compile(
    r"(?:mab|nib|platin|taxel|rubicin|mustine|tabine|arabine|phosphamide|"
    r"trexed|trexate|elbine|cycline|mycin|parib|ciclib|lisib|degib|"
    r"sertib|tinib|zumab|ximab|umab|vastatin|prazole|sartan|olol|"
    r"ustekinumab|cept|tide|sone|olone)\s*[a-z]?$", re.IGNORECASE)


# Positive evidence that a table reports OUTCOMES. The blacklist above only
# catches cost tables whose heading was anticipated; six of the eleven errors
# in the 50-row benchmark were cost, comparator or regimen tables phrased in
# a way it did not cover, landing under a clinical category. Requiring proof
# that outcome data is present inverts the default from "include unless it
# looks like a cost table" to "exclude unless it looks like results".
OUTCOME_EVIDENCE = [
    r"\[\s*-?[\d.,]+\s*[;:]\s*-?[\d.,]+\s*\]",       # [0,62; 1,01]
    r"\b(?:HR|RR|OR|RD|SMD|MD)\b\s*[:=]",            # HR = 0,79
    r"\b\d+\s*\(\s*\d+[.,]?\d*\s*%?\s*\)",           # 113 (22.3)
    r"\bp\s*[-\s]?(?:Wert|value)\b|\bp\s*[<=>]\s*0",  # p-Wert / p < 0.05
    r"\b(?:MV|MW)\s*\(\s*S[DE]\s*\)",                # MV (SD)
    r"\bn\.?\s?r\.?\b|\bk\.\s?A\.\b|\bn\.\s?d\.\b",  # not reached / k. A.
]
OUTCOME_EVIDENCE = [re.compile(p, re.IGNORECASE) for p in OUTCOME_EVIDENCE]

# Cells that mean money or dosing rather than results.
PRICE_OR_DOSE = re.compile(
    r"€|\bEUR\b|\bmg\b|\bml\b|\b[µu]g\b|Packung|Durchstechflasche|"
    r"Tablette|Filmtablette|Fertigspritze|pack\s+size|\bIE\b|\bI\.E\.\b",
    re.IGNORECASE)


def _has_outcome_evidence(table):
    """True when the cells actually carry results, not prices or doses."""
    cells = [str(c or "") for row in (table.get("rows") or []) for c in row]
    cells += [str(c or "") for c in (table.get("header") or [])]
    filled = [c for c in cells if c.strip()]
    if not filled:
        return False
    blob = " ".join(filled)
    if not any(p.search(blob) for p in OUTCOME_EVIDENCE):
        return False
    # Even with a CI somewhere, a table dominated by prices or dose units is
    # a cost table that happens to contain a number in brackets.
    money = sum(1 for c in filled if PRICE_OR_DOSE.search(c))
    return money / len(filled) <= 0.25


def _looks_like_drug(label):
    """True for a bare drug name: a short label, no digits, INN suffix."""
    text = re.sub(r"^nab-|\s*\(.*?\)\s*$", "", str(label or "").strip())
    if not text or len(text.split()) > 3 or any(c.isdigit() for c in text):
        return False
    return bool(DRUG_SUFFIX.search(text))

CATEGORY_RULES = [
    ("Mortality", r"mortalit|gesamt[uü]berleben|overall survival|\bdeath\b|"
                  r"todesf[aä]lle|\bsurvival\b|[uü]berlebenszeit"),
    # The QLQ-C30's FUNCTIONAL scales are quality of life, while its SYMPTOM
    # scales (below, under morbidity) are not. Same questionnaire, two
    # categories -- this pair of rules is where that split is enforced.
    ("Quality of life", r"quality of life|lebensqualit[aä]t|\bEQ-?5D\b|"
                        r"\bSF-?36\b|\bKCCQ\b|\bFACT-|health status|"
                        r"funktionsskalen|functional scale|"
                        r"(?:cognitive|social|emotional|role|physical)\s+"
                        r"functioning|"
                        r"(?:kognitive|soziale|emotionale|k[oö]rperliche)\s+"
                        r"funktion|rollenfunktion"),
    ("Adverse events", r"adverse event|side effect|nebenwirkung|"
                       r"unerw[uü]nschte|\bAEs?\b|\bSAEs?\b|\bSUE\b|\bUE\b|"
                       r"therapieabbruch|therapy discontinuation|toxicit|"
                       # MedDRA System Organ Classes, how AE tables are filed
                       r"infections and infestations|infektionen und|"
                       r"\b\w+\s+disorders\b|erkrankungen\s+des|"
                       r"investigations|untersuchungen"),
    ("Morbidity", r"morbidit|symptomatolog|symptomatik|recurrence|rezidiv|"
                  r"remission|response|ansprechen|visual|disease activity|"
                  r"krankheitsaktivit|exacerbation|hospitalis|progression|"
                  r"\bpain\b|schmerz|functional status|"
                  # QLQ-C30 symptom scales, which appear as bare rows in
                  # tables with no heading. In a table that DOES have an
                  # adverse-events heading these inherit it instead, because
                  # a heading always beats a row label.
                  r"insomnia|schlaflosigkeit|dyspno[ea]|atemnot|"
                  r"constipation|verstopfung|diarrho?ea|diarrh[oö]|durchfall|"
                  r"(?:appetite loss|loss of appetite)|appetitverlust|"
                  r"appetitlosigkeit|nausea|[uü]belkeit|"
                  r"vomiting|erbrechen|fatigue|m[uü]digkeit|lymphoedema|"
                  r"lymph[oö]dem"),
]

# Order matters: a heading naming both survival and adverse events is
# vanishingly rare, but mortality is the least ambiguous so it is tested
# first, and morbidity last because its vocabulary is the broadest.


def category_of(label):
    """The endpoint category a section heading names, or None."""
    text = str(label or "")
    if not text.strip():
        return None
    for name, pattern in CATEGORY_RULES:
        if re.search(pattern, text, re.IGNORECASE):
            return name
    return None


def _is_heading(row):
    """A spanning row: the first cell carries text and the rest are blank."""
    cells = [str(c or "").strip() for c in (row or [])]
    if not cells or not cells[0]:
        return False
    return not any(cells[1:])


def is_endpoint_table(table):
    """False for the cost, comparator and regimen tables sharing the file."""
    text = " ".join(str(c or "") for c in (table.get("header") or []))
    for row in (table.get("rows") or [])[:4]:
        if _is_heading(row):
            text += " " + str(row[0])
    if NOT_ENDPOINT.search(text):
        return False

    # A table whose row labels are mostly drug names lists a regimen or its
    # cost, whatever its heading says. Endpoint tables are labelled by
    # outcome ("Overall survival", "Fatigue"), not by active substance.
    labels = [str((row or [""])[0] or "").strip()
              for row in (table.get("rows") or []) if not _is_heading(row)]
    labels = [lab for lab in labels if lab]
    if len(labels) >= 3:
        drugs = sum(1 for lab in labels if _looks_like_drug(lab))
        if drugs / len(labels) > 0.5:
            return False

    # Last and strictest: the table must show evidence of reporting results.
    # This costs recall -- a purely descriptive endpoint table with no
    # effect estimate is now excluded -- and that is the right trade here,
    # because a cost table rendered under "Quality of life" looks
    # authoritative and is wrong, while a missing table is merely missing.
    return _has_outcome_evidence(table)


def sections(table):
    """[(heading, category, rows)] for one recovered table.

    Rows before any heading are returned under heading None, so a table that
    is all data and no heading is not silently dropped -- its category is
    then decided by the table's own header text.
    """
    if not is_endpoint_table(table):
        return []
    rows = table.get("rows") or []
    header_category = category_of(" ".join(
        str(c or "") for c in (table.get("header") or [])))

    out, current, heading, category = [], [], None, header_category
    inherited = header_category
    for row in rows:
        if _is_heading(row):
            if current:
                out.append((heading, category, current))
                current = []
            heading = str(row[0]).strip()
            found = category_of(heading)
            if found:
                inherited = found
            # A sub-heading that names no category keeps the one above it.
            category = found or inherited
        else:
            current.append(row)
    if current:
        out.append((heading, category, current))

    # Many tables carry no section heading at all -- they are a bare list of
    # endpoints, and the category lives in each ROW's own label ("Overall
    # mortality", "Pain (HAQ-DI VAS)"). Classify those row by row rather
    # than discarding the table, and keep rows that still match nothing
    # under None so the audit can see them.
    resolved = []
    for heading, found, block in out:
        if found:
            resolved.append((heading, found, block))
            continue
        buckets = {}
        for row in block:
            label = str((row or [""])[0] or "")
            if STRUCTURAL.match(label.strip()):
                continue
            buckets.setdefault(category_of(label), []).append(row)
        for name, rows_in in buckets.items():
            resolved.append((heading, name, rows_in))
    return resolved


def load(path=TABLES_FILE):
    if not os.path.exists(path):
        sys.exit(f"{path} not found. Run this beside it.")
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def for_category(tables, appraisal_id, category):
    """Every table section for one appraisal and one category."""
    out = []
    for table in (tables.get(str(appraisal_id).strip()) or []):
        for heading, found, rows in sections(table):
            if found == category and rows:
                out.append({"page": table.get("page"),
                            "heading": heading,
                            "header": table.get("header") or [],
                            "rows": rows})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tables", default=TABLES_FILE)
    ap.add_argument("--appraisal")
    ap.add_argument("--audit", action="store_true")
    ap.add_argument("--unmatched", type=int, default=25)
    args = ap.parse_args()
    tables = load(args.tables)
    print(f"{len(tables):,} appraisals, "
          f"{sum(len(v) for v in tables.values()):,} tables\n")

    if args.appraisal:
        got = tables.get(args.appraisal.strip())
        if not got:
            sys.exit(f"No tables recovered at all for {args.appraisal} — "
                     "the extraction found none in this resolution.")
        skipped = [t for t in got if not is_endpoint_table(t)]
        print(f"{len(got)} table(s) recovered · {len(skipped)} skipped as "
              f"cost/regimen tables · {len(got) - len(skipped)} endpoint "
              f"table(s)\n")
        if len(skipped) == len(got):
            print("  Every recovered table is a cost or regimen table, so "
                  "there is no\n  endpoint data to show. That is a real "
                  "finding, not a gap in the tool.\n")
        for name in ("Mortality", "Morbidity", "Quality of life",
                     "Adverse events"):
            found = for_category(tables, args.appraisal, name)
            print(f"=== {name}: {len(found)} section(s)")
            for block in found[:2]:
                print(f"    page {block['page']}  heading: {block['heading']}")
                for row in block["rows"][:3]:
                    cells = [str(c or "")[:22] for c in row]
                    print("      | " + " | ".join(cells))
            print()
        return

    from collections import Counter
    rows_total = rows_placed = 0
    per_category = Counter()
    unmatched = Counter()
    appraisals_with = Counter()
    skipped = 0
    for aid, group in tables.items():
        seen = set()
        for table in group:
            if not is_endpoint_table(table):
                skipped += 1
                continue
            for heading, category, rows in sections(table):
                rows_total += len(rows)
                if category:
                    rows_placed += len(rows)
                    per_category[category] += len(rows)
                    seen.add(category)
                elif heading:
                    unmatched[f"[heading] {heading[:50]}"] += 1
                else:
                    for row in rows[:3]:
                        label = str((row or [""])[0] or "").strip()
                        if label:
                            unmatched[f"[row] {label[:50]}"] += 1
        for category in seen:
            appraisals_with[category] += 1

    print(f"cost / comparator tables skipped: {skipped:,}")
    print(f"data rows placed in a category  : {rows_placed:,} of "
          f"{rows_total:,}  ({rows_placed / max(rows_total, 1):.1%})\n")
    print("rows by category:")
    for name, n in per_category.most_common():
        print(f"  {name:<18} {n:>7,} rows   "
              f"{appraisals_with[name]:>4} appraisals")
    if unmatched:
        print(f"\ntop unmatched section headings "
              f"({sum(unmatched.values()):,} sections):")
        for heading, n in unmatched.most_common(args.unmatched):
            print(f"  {n:>5}  {heading}")
        print("\nPaste these back and the keyword rules can be extended.")


if __name__ == "__main__":
    main()
