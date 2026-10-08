"""
Which indication has the deepest precedent set?

For picking a demo case for the counterfactual work: you want a disease with
many appraisals, because a counterfactual that moves a count from 2 to 3 shows
nothing. This ranks the diseases inside your coded set by how many appraisals
they have, how many share a mechanism, and how many have committee sections
deep enough to quote.

Restricted to the appraisals actually coded by extract_attributes.py, so the
numbers describe the data the app can reason over rather than the whole
workbook.

Usage:
  python indication_depth.py
  python indication_depth.py --workbook NICE_v14_updated_2026-09-28.xlsx
  python indication_depth.py --all          # ignore the coded subset
  python indication_depth.py --disease nsclc   # list the appraisals
"""

import argparse
import json
import os
import re

# Free-text indications fragment: "relapsed or refractory multiple myeloma"
# and "previously treated multiple myeloma" are one disease and two strings.
# So group on the disease. Order matters - the first match wins, which is why
# psoriatic arthritis sits above plaque psoriasis.
GROUPS = [
    ("NSCLC", r"non[- ]small[- ]cell|nsclc"),
    ("SCLC", r"small[- ]cell lung|sclc"),
    ("Breast cancer", r"breast cancer"),
    ("Multiple myeloma", r"myeloma"),
    ("Renal cell carcinoma", r"renal cell"),
    ("Prostate cancer", r"prostate"),
    ("Colorectal cancer", r"colorectal|bowel cancer"),
    ("Melanoma", r"melanoma"),
    ("Ovarian cancer", r"ovarian"),
    ("CLL", r"chronic lymphocytic"),
    ("AML", r"acute myeloid"),
    ("ALL", r"acute lymphoblastic"),
    ("DLBCL / follicular lymphoma", r"diffuse large b|follicular lymphoma"),
    ("Hodgkin lymphoma", r"hodgkin"),
    ("Bladder / urothelial cancer", r"urothelial|bladder cancer"),
    ("Hepatocellular carcinoma", r"hepatocellular"),
    ("Gastric / oesophageal cancer", r"gastric|oesophag|esophag"),
    ("Head and neck cancer", r"head and neck"),
    ("Psoriatic arthritis", r"psoriatic arthritis"),
    ("Plaque psoriasis", r"psoriasis"),
    ("Rheumatoid arthritis", r"rheumatoid arthritis"),
    ("Axial spondyloarthritis", r"spondyloarthritis|ankylosing spondylitis"),
    ("Ulcerative colitis", r"ulcerative colitis"),
    ("Crohn's disease", r"crohn"),
    ("Atopic dermatitis", r"atopic dermatitis|atopic eczema"),
    ("Multiple sclerosis", r"multiple sclerosis"),
    ("Migraine", r"migraine"),
    ("Epilepsy", r"epilep|seizure"),
    ("Asthma", r"asthma"),
    ("COPD", r"copd|chronic obstructive"),
    ("Cystic fibrosis", r"cystic fibrosis"),
    ("Type 2 diabetes", r"type 2 diabetes"),
    ("Heart failure", r"heart failure"),
    ("Chronic kidney disease", r"chronic kidney"),
    ("Transplant rejection", r"transplant"),
    ("COVID-19", r"covid"),
    ("Hepatitis", r"hepatitis"),
    ("HIV", r"\bhiv\b"),
    ("Sickle cell disease", r"sickle cell"),
    ("Haemophilia", r"haemophilia|hemophilia"),
    ("Spinal muscular atrophy", r"spinal muscular"),
    ("Duchenne muscular dystrophy", r"duchenne"),
    ("Amyloidosis", r"amyloidosis"),
    ("Macular degeneration / DMO", r"macular|diabetic macular"),
]


def disease_of(text):
    t = (text or "").lower()
    for name, pat in GROUPS:
        if re.search(pat, t):
            return name
    return None


def coded_tas(path):
    """TA numbers present in evidence_attributes.json, as a set of strings."""
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        recs = json.load(f)
    if isinstance(recs, dict):
        recs = list(recs.values())
    out = set()
    for r in recs:
        m = re.search(r"(\d+)", str(r.get("ta", "")))
        if m:
            out.add(m.group(1).lstrip("0") or "0")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", default="NICE_v14_updated_2026-09-28.xlsx")
    ap.add_argument("--sheet", default="Sheet1")
    ap.add_argument("--attributes", default="evidence_attributes.json")
    ap.add_argument("--sections", default="chapter_sections.json",
                    help="optional, adds a section count per disease")
    ap.add_argument("--all", action="store_true",
                    help="use the whole workbook, not just the coded subset")
    ap.add_argument("--disease", help="list the appraisals in one disease "
                                      "(substring match on the group name)")
    ap.add_argument("--top", type=int, default=20)
    args = ap.parse_args()

    try:
        import pandas as pd
    except ImportError:
        raise SystemExit("needs pandas")

    if not os.path.exists(args.workbook):
        raise SystemExit(f"Not found: {args.workbook}")
    df = pd.read_excel(args.workbook, args.sheet)

    df["_ta"] = (df["appraisal_id"].astype(str)
                 .str.extract(r"TA[^0-9]*(\d+)", expand=False)
                 .str.lstrip("0"))

    scope = "whole workbook"
    if not args.all:
        tas = coded_tas(args.attributes)
        if tas is None:
            print(f"note: {args.attributes} not found — falling back to the "
                  f"whole workbook. Pass --all to silence this.\n")
        else:
            df = df[df["_ta"].isin(tas)]
            scope = f"{len(tas)} coded appraisals"

    df["_disease"] = df["indication"].map(disease_of)

    # Sections per appraisal, if the parse output is to hand.
    sec_count = {}
    if os.path.exists(args.sections):
        try:
            with open(args.sections, encoding="utf-8") as f:
                for s in json.load(f):
                    k = str(s.get("ta", "")).lstrip("0")
                    sec_count[k] = sec_count.get(k, 0) + 1
        except (json.JSONDecodeError, OSError):
            pass

    if args.disease:
        want = args.disease.lower()
        names = [n for n, _ in GROUPS if want in n.lower()]
        if not names:
            raise SystemExit(f"No disease group matches {args.disease!r}. "
                             f"Try one of: {', '.join(n for n, _ in GROUPS[:8])} ...")
        sub = df[df["_disease"].isin(names)].copy()
        sub["_sec"] = sub["_ta"].map(lambda t: sec_count.get(t, 0))
        sub = sub.sort_values("_ta", key=lambda s: s.astype(float))
        print(f"{', '.join(names)} — {len(sub)} appraisals in {scope}\n")
        def _txt(v):
            # pandas gives NaN for blanks, and str(NaN) is "nan", which then
            # prints as if it were data.
            s = "" if v is None else str(v).strip()
            return "" if s.lower() in ("nan", "none", "") else s

        for _, r in sub.iterrows():
            moa = _txt(r.get("mechanism_of_action")) or "—"
            bio = _txt(r.get("biomarker"))
            print(f"  TA{r['_ta']:<6} {str(r['drug_name'])[:26]:<27} "
                  f"{moa[:28]:<29} {r.get('decision_simple')}")
            print(f"           {str(r['indication'])[:100]}")
            if bio:
                print(f"           biomarker: {bio[:80]}")
            if r["_sec"]:
                print(f"           {r['_sec']} committee sections indexed")
        return

    print(f"Scope: {scope}\n")
    rows = []
    for name, _ in GROUPS:
        sub = df[df["_disease"] == name]
        if sub.empty:
            continue
        moas = sub["mechanism_of_action"].astype(str).str.strip()
        moas = moas[~moas.str.lower().isin(["", "nan", "none"])]
        secs = sum(sec_count.get(t, 0) for t in sub["_ta"] if t)
        rows.append((name, len(sub), moas.nunique(), secs))

    rows.sort(key=lambda r: -r[1])
    print(f"{'disease':<30}{'appraisals':>11}{'distinct MOAs':>15}"
          f"{'sections':>10}")
    print("-" * 66)
    for name, n, nm, secs in rows[: args.top]:
        print(f"{name:<30}{n:>11}{nm:>15}{secs:>10}")

    unmatched = int(df["_disease"].isna().sum())
    print(f"\n{unmatched} of {len(df)} appraisals did not match a disease "
          f"group — mostly rare conditions with one appraisal each, which are")
    print("the opposite of what you want for a precedent demo.")
    print("\nA good demo case needs: many appraisals, several mechanisms (so "
          "the comparator question is live), and plenty of indexed sections")
    print("(so answers can quote). Sort by whichever of those matters most.")
    print("\nCaveat: the biomarker and mechanism_of_action columns are only "
          "partly populated in the workbook, so a 0 means 'not recorded',")
    print("not 'no targeted therapy exists'. Melanoma is the obvious example.")


if __name__ == "__main__":
    main()
