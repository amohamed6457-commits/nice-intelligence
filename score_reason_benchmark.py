import argparse
import hashlib
import json
import math
import os
import sys
from collections import Counter, defaultdict

from classify_reasons import CATEGORIES

SHEET = "reason_benchmark.xlsx"
LABELS = "reason_categories.json"


def wilson(hits, n, z=1.96):
    if not n:
        return 0.0, 0.0
    p = hits / n
    d = 1 + z * z / n
    centre = (p + z * z / (2 * n)) / d
    half = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return max(0.0, centre - half), min(1.0, centre + half)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sheet", default=SHEET)
    ap.add_argument("--labels", default=LABELS)
    ap.add_argument("--show", action="store_true")
    args = ap.parse_args()

    try:
        import openpyxl
    except ImportError:
        sys.exit("pip install openpyxl")
    for path in (args.sheet, args.labels):
        if not os.path.exists(path):
            sys.exit(f"{path} not found.")

    with open(args.labels, encoding="utf-8") as fh:
        labels = json.load(fh)

    wb = openpyxl.load_workbook(args.sheet, data_only=True)
    if "Coding" not in wb.sheetnames:
        sys.exit("No 'Coding' sheet in that workbook.")
    sheet = wb["Coding"]
    header = [str(c.value or "").strip() for c in sheet[1]]
    try:
        ci = {name: header.index(name) for name in
              ("Item", "Ref", "Appraisal", "Page", "Sentence", "Your_Category")}
    except ValueError as exc:
        sys.exit(f"Column missing from the Coding sheet: {exc}")

    rows, blank, unknown_ref, unsure = [], 0, 0, 0
    for row in sheet.iter_rows(min_row=2, values_only=True):
        ref = str(row[ci["Ref"]] or "").strip()
        if not ref:
            continue
        mine = str(row[ci["Your_Category"]] or "").strip().upper()
        if not mine:
            blank += 1
            continue
        if mine == "UNSURE":
            unsure += 1
            continue
        if mine not in CATEGORIES:
            sys.exit(f"Item {row[ci['Item']]}: '{mine}' is not a category "
                     f"code. Fix it in the sheet and re-run.")
        got = labels.get(ref)
        if not got or str(got).startswith("__"):
            unknown_ref += 1
            continue
        rows.append((row[ci["Item"]], row[ci["Appraisal"]], row[ci["Page"]],
                     str(row[ci["Sentence"]] or ""), mine, got))

    print(f"coded and scorable : {len(rows)}")
    if blank:
        print(f"left blank         : {blank}")
    if unsure:
        print(f"marked UNSURE      : {unsure}  (excluded from the "
              f"denominator)")
    if unknown_ref:
        print(f"no prediction      : {unknown_ref}  (Ref not in "
              f"{args.labels})")
    if not rows:
        sys.exit("\nNothing to score.")

    agree = sum(1 for *_x, mine, got in rows if mine == got)
    n = len(rows)
    lo, hi = wilson(agree, n)
    print(f"\nAGREEMENT  {agree}/{n} = {agree / n:.1%}"
          f"   95% CI {lo:.1%} to {hi:.1%}")

    per = defaultdict(lambda: [0, 0])
    for *_x, mine, got in rows:
        per[got][1] += 1
        if mine == got:
            per[got][0] += 1
    print("\nWhere the model said X, how often you agreed:")
    print(f"  {'category':<32} {'agree':>9}   {'rate':>6}")
    for code in sorted(per, key=lambda c: -per[c][1]):
        hits, total = per[code]
        print(f"  {code:<32} {hits:>4}/{total:<4} {hits / total:>6.0%}")

    recall = defaultdict(lambda: [0, 0])
    for *_x, mine, got in rows:
        recall[mine][1] += 1
        if mine == got:
            recall[mine][0] += 1
    missed = {c: v for c, v in recall.items() if v[0] < v[1]}
    if missed:
        print("\nWhere you said X, how often the model found it:")
        for code in sorted(missed, key=lambda c: -missed[c][1]):
            hits, total = missed[code]
            print(f"  {code:<32} {hits:>4}/{total:<4} {hits / total:>6.0%}")

    pairs = Counter((got, mine) for *_x, mine, got in rows if mine != got)
    if pairs:
        print(f"\n{sum(pairs.values())} disagreement(s), most common first:")
        for (got, mine), count in pairs.most_common():
            print(f"  {count:>3}x  model {got}  ->  you {mine}")

    if args.show and pairs:
        print("\n" + "=" * 70)
        for item, aid, page, text, mine, got in rows:
            if mine == got:
                continue
            print(f"\nitem {item}  {aid} p{page}")
            print(f"  {' '.join(text[:300].split())}")
            print(f"  model: {got}")
            print(f"  you  : {mine}")

    print("\n" + "=" * 70)
    if n < 25:
        print(f"{n} items is too few to quote. The interval above is "
              f"{hi - lo:.0%} wide.")
    elif lo >= 0.80:
        print(f"The lower bound is {lo:.0%}, so 80% holds even at the "
              f"pessimistic end.")
    elif agree / n >= 0.80:
        print(f"Point estimate clears 80% but the lower bound is {lo:.0%}. "
              f"Quote the\ninterval, not the {agree / n:.0%}.")
    else:
        print(f"Below 80%. Re-run --show, read the disagreements, and decide "
              f"whether they\nare the model being wrong or the categories "
              f"being badly drawn. Do not put\nthis on screen at "
              f"{agree / n:.0%}.")
    if unsure:
        print(f"\n{unsure} UNSURE: sentences a human could not label either. "
              f"Those are the\nextractor's fault, not the classifier's.")


if __name__ == "__main__":
    main()
