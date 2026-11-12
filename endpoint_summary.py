"""
Turn G-BA subgroup assessments into a grid you can read in two seconds.

The problem this solves
-----------------------
The record view currently shows passages. A reader who has to read all of it
could have opened the PDF instead, so the passage view adds nothing. What a
reader actually wants first is the verdict per endpoint category, and only
then the numbers behind it.

G-BA already writes that verdict. Every subgroup row carries four one-line
fields -- mortality, morbidity, quality of life, adverse events -- filled on
1,713 of 1,752 rows. They are short and highly formulaic: 18 phrases cover
80% of all 6,852 values. So they can be reduced to a signal.

    verdict            share of all 6,852 values
    NO DATA                   65%
    NO DIFFERENCE             17%
    ADVANTAGE                 12%
    DISADVANTAGE               3%
    MIXED / NO EVENTS          3%

and by category:

    category          no data   no diff   advantage   disadvantage
    Mortality            61%      22%        11%           0%
    Morbidity            63%      13%        20%           2%
    Quality of life      77%      14%         8%           1%
    Adverse events       62%      18%         7%           9%

That quality-of-life row is the headline finding: across 1,713 subgroup
assessments, G-BA judged the quality-of-life evidence unusable in 77% of
them. It is the evidence gap, quantified, and it is the strongest argument
for generating real-world evidence that this dataset contains.

Provenance
----------
These fields come from G-BA's AIS database, NOT from parsing the resolution
PDF, so they have no page number and never will -- the text was never on a
page. That is a real limitation and the UI must say so. What they do have is
better than an extraction: they are the agency's own summary, citable to the
appraisal's G-BA URL. Label them "G-BA AIS record" and link the appraisal.

The quantified endpoints that sit UNDER each verdict are a separate matter:
those come from the recovered tables and do carry a page.

Translation
-----------
About 80% of values are formulaic ("Es liegen keine bewertbaren Daten vor")
and are rendered into English from a template here -- no API call, no cost,
no translation error. The remaining ADVANTAGE / DISADVANTAGE lines carry
real clinical content and vary, so those are collected by --needs-translation
for translate_passages.py to handle.

    python endpoint_summary.py                     # the cross-tab above
    python endpoint_summary.py --appraisal D-792   # one appraisal's grid
    python endpoint_summary.py --needs-translation # phrases worth translating
"""

import argparse
import json
import os
import re
import sys

WORKBOOKS = ["NICE_HAS_GBA_updated_2026-11-07.xlsx",
             "NICE_v14_updated_2026-09-29.xlsx"]

CATEGORIES = [("mortality", "Mortality"),
              ("morbidity", "Morbidity"),
              ("quality_of_life", "Quality of life"),
              ("adverse_events", "Adverse events")]

# Order matters: the first pattern that matches wins, and ADVANTAGE /
# DISADVANTAGE are tested before the "no ..." patterns so that a sentence
# reporting a benefit is never swallowed by a stray negation.
_ADV = r"\bVorteil|Verbesserung|Reduktion\s+von|Verringerung\s+von"
_DIS = r"\bNachteil|Verschlechterung|Erhöhung\s+von\s+.{0,30}Ereignis"

RULES = [
    ("NO DATA",
     r"keine\s+(?:bewertbaren\s+|geeigneten\s+|verwertbaren\s+|relevanten\s+|"
     r"vergleichenden\s+|für die Nutzenbewertung\s+\w+\s+)?Daten|"
     r"nicht\s+bewertbar|keine\s+Daten\s+vorgelegt|keine\s+Angaben|"
     r"nicht\s+vorgelegt|keine\s+Auswertung|nicht\s+vollständig|"
     r"nicht\s+für\s+die\s+Nutzenbewertung\s+herangezogen"),
    ("NO EVENTS",
     r"traten\s+keine|keine\s+Ereignisse|keine\s+Todesfälle"),
    ("NO DIFFERENCE",
     r"kein[e]?\s+(?:für die Nutzenbewertung\s+)?relevante[rn]?\s+Unterschied|"
     r"kein\s+Unterschied|keine\s+Unterschiede|weder\s+Vorteile\s+noch\s+Nachteile|"
     r"kein\s+statistisch\s+signifikanter"),
    ("DISADVANTAGE", _DIS),
    ("ADVANTAGE", _ADV),
]

# What the reader sees instead of the German, for the formulaic verdicts.
# These are renderings of a fixed phrase, not translations of free text.
PLAIN = {
    "NO DATA":       "No usable data submitted",
    "NO DIFFERENCE": "No difference relevant to the assessment",
    "NO EVENTS":     "No events occurred",
    "ADVANTAGE":     "Advantage",
    "DISADVANTAGE":  "Disadvantage",
    "MIXED":         "Advantages and disadvantages",
}

# For the UI: how each verdict should read at a glance.
TONE = {
    "NO DATA": "gap", "NO DIFFERENCE": "neutral", "NO EVENTS": "neutral",
    "ADVANTAGE": "positive", "DISADVANTAGE": "negative", "MIXED": "mixed",
}

# Verdicts whose German text carries clinical content worth translating.
INFORMATIVE = {"ADVANTAGE", "DISADVANTAGE", "MIXED"}


def verdict(text):
    """One of the six labels, or None if the phrase is not recognised."""
    t = str(text or "").strip()
    if len(t) < 3:
        return None
    if re.search(_ADV, t, re.I) and re.search(_DIS, t, re.I):
        return "MIXED"
    for label, pattern in RULES:
        if re.search(pattern, t, re.I):
            return label
    return None


def find_workbook(explicit=None):
    if explicit:
        if os.path.exists(explicit):
            return explicit
        sys.exit(f"Workbook not found: {explicit}")
    for w in WORKBOOKS:
        if os.path.exists(w):
            return w
    sys.exit("No workbook found. Pass --workbook.")


def _translations():
    """Cached English for the informative phrases, if a run has produced any."""
    for path in ("translations.json", "endpoint_translations.json"):
        if os.path.exists(path):
            try:
                with open(path, encoding="utf-8") as fh:
                    return json.load(fh)
            except (ValueError, OSError):
                pass
    return {}


def _english(text, label, cache):
    """What to show the reader. Formulaic verdicts render from PLAIN; the
    informative ones use a cached translation when one exists, and otherwise
    fall back to the label rather than inventing an English sentence."""
    if label in INFORMATIVE:
        import hashlib
        key = hashlib.sha1(str(text).strip().encode("utf-8")).hexdigest()[:16]
        got = cache.get(key)
        if got and not str(got).startswith("__ERROR__"):
            return got
        return PLAIN.get(label, label)
    return PLAIN.get(label, label)


def grid(workbook, appraisal_id=None):
    """Per subgroup, the four category verdicts with their source."""
    import pandas as pd
    subs = pd.read_excel(workbook, sheet_name="GBA_Subgroups")
    subs["appraisal_id"] = subs["appraisal_id"].astype(str).str.strip()
    if appraisal_id:
        subs = subs[subs["appraisal_id"] == str(appraisal_id).strip()]
        if not len(subs):
            sys.exit(f"No subgroup rows for {appraisal_id}.")
    cache = _translations()
    out = []
    for r in subs.itertuples():
        cats = []
        for field, name in CATEGORIES:
            raw = getattr(r, field, None)
            label = verdict(raw)
            cats.append({
                "category": name,
                "verdict": label,
                "tone": TONE.get(label, "unknown"),
                "english": _english(raw, label, cache) if label else None,
                "original": (str(raw).strip() if raw is not None else None),
            })
        out.append({
            "appraisal_id": r.appraisal_id,
            "subgroup_id": getattr(r, "subgroup_id", None),
            "drug": getattr(r, "generic_name", None),
            "subgroup": getattr(r, "patient_subgroup", None),
            "comparator": getattr(r, "comparator_therapy", None),
            "benefit_extent": getattr(r, "benefit_extent", None),
            "categories": cats,
            # Honest provenance: an agency database record, not a parsed page.
            "source_kind": "G-BA AIS record",
            "source_url": getattr(r, "official_source_url", None),
            "source_page": None,
        })
    return out


def crosstab(workbook):
    import pandas as pd
    subs = pd.read_excel(workbook, sheet_name="GBA_Subgroups")
    order = ["NO DATA", "NO DIFFERENCE", "ADVANTAGE", "DISADVANTAGE",
             "MIXED", "NO EVENTS"]
    rows, total, known = [], 0, 0
    for field, name in CATEGORIES:
        vals = subs[field].dropna().astype(str)
        vals = vals[vals.str.strip().str.len() > 2]
        v = vals.map(verdict)
        total += len(v)
        known += v.notna().sum()
        share = v.value_counts(normalize=True)
        rows.append([name, len(v)] + [f"{share.get(k, 0):.0%}" for k in order])
    width = [17, 6] + [13 for _ in order]
    head = ["category", "n"] + [k.lower() for k in order]
    print("  ".join(h.ljust(w) for h, w in zip(head, width)))
    print("  ".join("-" * w for w in width))
    for row in rows:
        print("  ".join(str(c).ljust(w) for c, w in zip(row, width)))
    print(f"\nclassified {known:,}/{total:,} = {known / total:.1%}")
    print("\nSource: G-BA AIS subgroup records. These are the agency's own "
          "summary\nfields, not extracted from the resolution PDF, so they "
          "carry no page number.")


def needs_translation(workbook):
    """The phrases whose German actually carries clinical content."""
    import pandas as pd
    from collections import Counter
    subs = pd.read_excel(workbook, sheet_name="GBA_Subgroups")
    counts = Counter()
    for field, _ in CATEGORIES:
        for raw in subs[field].dropna().astype(str):
            text = raw.strip()
            if verdict(text) in INFORMATIVE:
                counts[text] += 1
    chars = sum(len(t) * 1 for t in counts)
    print(f"{len(counts):,} distinct phrases carry clinical content "
          f"({sum(counts.values()):,} values, {chars:,} characters).")
    print("The other ~80% are formulaic and render from a template for free.\n")
    with open("endpoint_phrases.json", "w", encoding="utf-8") as fh:
        json.dump(sorted(counts), fh, ensure_ascii=False, indent=1)
    print("-> endpoint_phrases.json")
    for text, n in counts.most_common(12):
        print(f"  {n:>4}  {text[:90]}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook")
    ap.add_argument("--appraisal", help="print one appraisal's grid")
    ap.add_argument("--needs-translation", action="store_true")
    args = ap.parse_args()
    wb = find_workbook(args.workbook)
    print(f"workbook: {wb}\n")

    if args.needs_translation:
        return needs_translation(wb)
    if args.appraisal:
        for row in grid(wb, args.appraisal):
            print(f"=== {row['appraisal_id']}  {row['drug']} "
                  f"(subgroup {row['subgroup_id']}) ===")
            print(f"  indication : {str(row['subgroup'])[:100]}")
            print(f"  comparator : {str(row['comparator'])[:100]}")
            print(f"  decision   : {row['benefit_extent']}")
            for c in row["categories"]:
                print(f"    {c['category']:<17} {str(c['verdict']):<14} "
                      f"{str(c['english'])[:60]}")
            print(f"  source     : {row['source_kind']} - no page anchor")
            print(f"               {str(row['source_url'])[:92]}\n")
        return
    crosstab(wb)


if __name__ == "__main__":
    main()
