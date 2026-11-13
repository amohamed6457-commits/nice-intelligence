"""
Score recovered tables against the hand-coded reference standard.

This is the measurement the platform has never had. Agreement between
extraction runs is known (89% unanimous across 6,305 cells); accuracy against
a human has never been measured, and those are different claims.

What is compared
----------------
The reference is 181 endpoint rows hand-coded from 20 G-BA resolutions. The
candidate is gba_tables.json, the tables recovered from the same resolutions'
PDFs by extract_gba_tables.py.

Matching by endpoint NAME would be unreliable — the reference names are the
coder's English renderings of German headings, and a fuzzy string match
between "Spleen volume reduction >=35%" and "Reduktion des Milzvolumens um
>= 35 %" would measure my matcher, not the extraction. So the comparison is
on the NUMBERS, which are language-independent and are what a reader would
check:

  FOUND TOGETHER  the effect value and both CI bounds appear in ONE recovered
                  table row — the extraction captured this endpoint intact
  SPLIT           all three appear somewhere in the appraisal's tables, but
                  not in a single row — the values survived, the structure
                  did not
  PARTIAL         some of the three appear, others do not
  MISSED          the effect value does not appear anywhere

A reference row the coder marked "None reported" is scored separately: there
is no number to find, so the question is whether extraction wrongly invented
one. Those are counted but kept out of the headline rate, and reported on
their own.

Numbers are compared after normalising German decimal commas, so 0,79 and
0.79 are the same value, and with a relative tolerance for rounding.

    python score_extraction.py
    python score_extraction.py --reference REFERENCE_TABLE_EDITED_2.xlsx
    python score_extraction.py --detail D-161
"""

import argparse
import json
import os
import re
import sys
from collections import Counter

REFERENCE = "REFERENCE_TABLE_EDITED_2.xlsx"
TABLES = "gba_tables.json"
TOLERANCE = 0.011          # relative; covers 2-dp rounding both ways


def num(value):
    """Float from a cell or a table cell, handling German decimal commas."""
    if value is None:
        return None
    s = str(value).strip()
    if not s or s.lower() in {"nr", "n.d.", "k.a.", "none", "na", "n/a", "-"}:
        return None
    s = s.replace("<", "").replace(">", "").replace("≤", "").replace("≥", "")
    s = s.replace("−", "-").strip()
    # 1.234,56 -> 1234.56 ; 0,79 -> 0.79 ; 1,234.56 -> 1234.56
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") \
            else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except ValueError:
        return None


def close(a, b):
    if a is None or b is None:
        return False
    if a == b:
        return True
    scale = max(abs(a), abs(b), 1e-9)
    return abs(a - b) / scale <= TOLERANCE


def cell_numbers(text):
    """Every number in one recovered cell — a cell may hold '0,79 [0,62; 1,01]'."""
    out = []
    for token in re.findall(r"-?\d[\d.,]*", str(text or "")):
        v = num(token)
        if v is not None:
            out.append(v)
    return out


def load_reference(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Coding"] if "Coding" in wb.sheetnames else wb[wb.sheetnames[0]]
    hdr = [c.value for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(hdr) if h}
    need = ["appraisal_id", "endpoint", "effect_value", "ci_lower", "ci_upper",
            "endpoint_category", "effect_measure"]
    missing = [c for c in need if c not in idx]
    if missing:
        sys.exit(f"{path} is missing column(s): {', '.join(missing)}")
    rows, current = [], None
    for r in range(2, ws.max_row + 1):
        v = {h: ws.cell(r, i).value for h, i in idx.items()}
        if v.get("appraisal_id"):
            current = str(v["appraisal_id"]).strip()
        if not any(v.get(h) not in (None, "") for h in list(idx)[3:]):
            continue
        v["appraisal_id"] = current
        v["_row"] = r
        rows.append(v)
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--reference", default=REFERENCE)
    ap.add_argument("--tables", default=TABLES)
    ap.add_argument("--detail", help="print every row for one appraisal")
    args = ap.parse_args()

    if not os.path.exists(args.reference):
        xl = [f for f in os.listdir(".") if f.lower().endswith((".xlsx", ".xlsm"))]
        sys.exit(f"Reference not found: {args.reference}\nWorkbooks here: {', '.join(xl)}")
    if not os.path.exists(args.tables):
        sys.exit(f"{args.tables} not found — run extract_gba_tables.py for the "
                 "20 reference appraisals first.")

    reference = load_reference(args.reference)
    with open(args.tables, encoding="utf-8") as fh:
        tables = json.load(fh)

    print(f"reference : {args.reference}  ({len(reference)} rows, "
          f"{len({r['appraisal_id'] for r in reference})} appraisals)")
    print(f"candidate : {args.tables}  ({len(tables)} appraisals, "
          f"{sum(len(v) for v in tables.values())} tables)\n")

    verdicts, per_appraisal, detail = Counter(), {}, []
    for ref in reference:
        aid = ref["appraisal_id"]
        got = tables.get(aid) or []
        rows = [r for t in got for r in (t.get("rows") or [])]
        row_numbers = [sum((cell_numbers(c) for c in row), []) for row in rows]
        every = [n for nums in row_numbers for n in nums]

        e, lo, hi = (num(ref["effect_value"]), num(ref["ci_lower"]),
                     num(ref["ci_upper"]))
        measure = str(ref.get("effect_measure") or "").lower()
        if e is None and measure.startswith("none"):
            invented = "none-reported: extraction also reports nothing"
            verdict = "NONE-REPORTED"
        elif e is None:
            verdict = "SKIPPED (no value coded)"
        else:
            wanted = [v for v in (e, lo, hi) if v is not None]
            together = any(all(any(close(w, n) for n in nums) for w in wanted)
                           for nums in row_numbers)
            found = [w for w in wanted if any(close(w, n) for n in every)]
            if together:
                verdict = "FOUND TOGETHER"
            elif len(found) == len(wanted):
                verdict = "SPLIT"
            elif found:
                verdict = "PARTIAL"
            else:
                verdict = "MISSED"
        verdicts[verdict] += 1
        per_appraisal.setdefault(aid, Counter())[verdict] += 1
        detail.append((aid, str(ref["endpoint"])[:44], verdict, e, lo, hi))

    scored = sum(v for k, v in verdicts.items()
                 if k in {"FOUND TOGETHER", "SPLIT", "PARTIAL", "MISSED"})
    print("=" * 66)
    print("VERDICTS")
    for k in ("FOUND TOGETHER", "SPLIT", "PARTIAL", "MISSED",
              "NONE-REPORTED", "SKIPPED (no value coded)"):
        if verdicts[k]:
            share = f"{verdicts[k] / scored:6.1%}" if scored and k in {
                "FOUND TOGETHER", "SPLIT", "PARTIAL", "MISSED"} else "      "
            print(f"  {verdicts[k]:4d}  {share}  {k}")
    if scored:
        intact = verdicts["FOUND TOGETHER"] / scored
        anywhere = (verdicts["FOUND TOGETHER"] + verdicts["SPLIT"]) / scored
        lo_ci, hi_ci = wilson(verdicts["FOUND TOGETHER"], scored)
        print(f"\n  endpoints recovered INTACT : {intact:.1%}  "
              f"(95% CI {lo_ci:.1%} to {hi_ci:.1%}, n={scored})")
        print(f"  values present ANYWHERE    : {anywhere:.1%}")

    print("\nBY APPRAISAL")
    for aid in sorted(per_appraisal, key=lambda a: -per_appraisal[a].get("MISSED", 0)):
        c = per_appraisal[aid]
        tot = sum(c[k] for k in ("FOUND TOGETHER", "SPLIT", "PARTIAL", "MISSED"))
        rate = f"{c['FOUND TOGETHER'] / tot:5.0%}" if tot else "    –"
        print(f"  {aid:<8} {rate} intact   found {c['FOUND TOGETHER']:>3}  "
              f"split {c['SPLIT']:>3}  partial {c['PARTIAL']:>3}  "
              f"missed {c['MISSED']:>3}   ({len(tables.get(aid) or [])} tables recovered)")

    if args.detail:
        print(f"\nDETAIL — {args.detail}")
        for aid, endpoint, verdict, e, lo, hi in detail:
            if aid == args.detail:
                print(f"  {verdict:<16} {endpoint:<46} {e} [{lo}; {hi}]")

    print("\nNote: 'intact' means the effect value and both CI bounds were found "
          "in ONE\nrecovered row. 'split' means the numbers survived but the row "
          "structure did not.")


def wilson(successes, n, z=1.96):
    if not n:
        return 0.0, 0.0
    p = successes / n
    d = 1 + z * z / n
    centre = (p + z * z / (2 * n)) / d
    half = z * ((p * (1 - p) / n + z * z / (4 * n * n)) ** 0.5) / d
    return max(0.0, centre - half), min(1.0, centre + half)


if __name__ == "__main__":
    main()
