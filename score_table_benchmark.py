"""
Score the table classifier against the hand-coded benchmark.

Reads table_benchmark.xlsx, recomputes what endpoint_tables.py predicts for
each sampled section, and compares. No answer key file exists: predictions
are derived from gba_tables.json at scoring time, so the coding sheet could
never have leaked them.

What is reported
----------------
PRECISION per category -- when the tool says "Quality of life", how often is
it right. That is the question the record view asks, because a wrongly filed
table appears under a verdict looking authoritative.

The confusion matrix shows WHERE the errors go. "Quality of life" rows
landing in Morbidity is the EORTC QLQ-C30 split failing; anything landing in
"Not an endpoint table" means cost or regimen tables are getting through,
which is the more embarrassing failure.

Wilson intervals, because 10 per category is a small sample and a point
estimate alone would overstate what it tells you.

    python score_table_benchmark.py
    python score_table_benchmark.py --errors
"""

import argparse
import json
import os
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import endpoint_tables
    from make_table_benchmark import UNPLACED, collect
except ImportError as exc:
    sys.exit(f"endpoint_tables.py and make_table_benchmark.py must sit "
             f"beside this script ({exc}).")

NOT_ENDPOINT_LABEL = "Not an endpoint table"
CANNOT_TELL = "Cannot tell"


def wilson(successes, n, z=1.96):
    if not n:
        return 0.0, 0.0
    p = successes / n
    d = 1 + z * z / n
    centre = (p + z * z / (2 * n)) / d
    half = z * ((p * (1 - p) / n + z * z / (4 * n * n)) ** 0.5) / d
    return max(0.0, centre - half), min(1.0, centre + half)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--benchmark", default="table_benchmark.xlsx")
    ap.add_argument("--tables", default=endpoint_tables.TABLES_FILE)
    ap.add_argument("--errors", action="store_true",
                    help="print every disagreement in full")
    args = ap.parse_args()

    for path in (args.benchmark, args.tables):
        if not os.path.exists(path):
            sys.exit(f"{path} not found.")

    import openpyxl
    wb = openpyxl.load_workbook(args.benchmark, data_only=True)
    ws = wb["Coding"] if "Coding" in wb.sheetnames else wb[wb.sheetnames[0]]
    header = [str(c.value or "").strip().lower() for c in ws[2]]
    try:
        col = {name: header.index(name) + 1 for name in
               ("appraisal_id", "page", "section heading", "true category")}
    except ValueError:
        sys.exit("Unexpected columns. Was this made by "
                 "make_table_benchmark.py?")

    with open(args.tables, encoding="utf-8") as fh:
        tables = json.load(fh)
    # Rebuild predictions and key them the way the sheet identifies a row.
    predictions = defaultdict(list)
    for section in collect(tables):
        key = (str(section["appraisal_id"]), section["page"],
               section["heading"])
        predictions[key].append(section["predicted"])

    coded = matched = 0
    pairs = []
    unknown = 0
    for r in range(3, ws.max_row + 1):
        truth = str(ws.cell(r, col["true category"]).value or "").strip()
        if not truth:
            continue
        coded += 1
        key = (str(ws.cell(r, col["appraisal_id"]).value or "").strip(),
               ws.cell(r, col["page"]).value,
               str(ws.cell(r, col["section heading"]).value or "").strip())
        got = predictions.get(key)
        if not got:
            unknown += 1
            continue
        matched += 1
        pairs.append((key, got[0], truth))

    print(f"benchmark : {args.benchmark}")
    print(f"coded rows: {coded}   matched to a prediction: {matched}")
    if unknown:
        print(f"  ({unknown} row(s) could not be matched back — "
              "gba_tables.json may have changed since the sheet was made)")
    if not pairs:
        sys.exit("\nNothing to score yet. Fill in the 'true category' column.")

    # The tool's UNPLACED and the coder's "not an endpoint table" agree.
    def same(predicted, truth):
        if truth == CANNOT_TELL:
            return None
        if truth == NOT_ENDPOINT_LABEL:
            return predicted == UNPLACED
        return predicted == truth

    per_category = defaultdict(lambda: [0, 0])     # predicted -> [right, n]
    confusion = Counter()
    undecidable = 0
    for _key, predicted, truth in pairs:
        verdict = same(predicted, truth)
        if verdict is None:
            undecidable += 1
            continue
        per_category[predicted][1] += 1
        if verdict:
            per_category[predicted][0] += 1
        else:
            confusion[(predicted, truth)] += 1

    right = sum(v[0] for v in per_category.values())
    total = sum(v[1] for v in per_category.values())
    print("\n" + "=" * 62)
    print("PRECISION — when the tool says X, how often is it right\n")
    for name in ["Mortality", "Morbidity", "Quality of life",
                 "Adverse events", UNPLACED]:
        ok, n = per_category.get(name, [0, 0])
        if not n:
            continue
        lo, hi = wilson(ok, n)
        print(f"  {name:<18} {ok:>3}/{n:<3} {ok / n:>6.0%}   "
              f"95% CI {lo:.0%}–{hi:.0%}")
    if total:
        lo, hi = wilson(right, total)
        print(f"\n  OVERALL            {right:>3}/{total:<3} {right / total:>6.0%}   "
              f"95% CI {lo:.0%}–{hi:.0%}")
    if undecidable:
        print(f"\n  {undecidable} row(s) coded 'Cannot tell', excluded from "
              "the rate above.")

    if confusion:
        print("\nWHERE THE ERRORS GO")
        for (predicted, truth), n in confusion.most_common():
            print(f"  {n:>3}  tool said {predicted:<18} "
                  f"actually {truth}")
        print("\n  'Quality of life' confused with 'Morbidity' is the "
              "EORTC QLQ-C30 split\n  failing. Anything confused with 'Not an "
              "endpoint table' means cost or\n  regimen tables are still "
              "getting through, which is the worse failure.")

    if args.errors:
        print("\nEVERY DISAGREEMENT")
        for key, predicted, truth in pairs:
            if same(predicted, truth) is False:
                print(f"  {key[0]} page {key[1]}: {key[2][:50]}")
                print(f"      tool: {predicted}   coded: {truth}")


if __name__ == "__main__":
    main()
