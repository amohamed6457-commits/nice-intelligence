import argparse
import hashlib
import json
import os
import random
import sys
from collections import Counter, defaultdict

from classify_reasons import CATEGORIES, SYSTEM

REASONS = "reasons.json"
LABELS = "reason_categories.json"
OUT = "reason_benchmark.xlsx"


def key_of(text):
    return hashlib.sha1(text.strip().encode("utf-8")).hexdigest()[:16]


def load(path):
    if not os.path.exists(path):
        sys.exit(f"{path} not found.")
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def candidates(reasons, labels):
    seen = defaultdict(list)
    for aid, entries in reasons.items():
        for entry in entries:
            text = str(entry.get("sentence") or "").strip()
            if len(text) < 40:
                continue
            key = key_of(text)
            got = labels.get(key)
            if not got or str(got).startswith("__"):
                continue
            seen[key].append((aid, entry.get("page"), text, got))
    return seen


def allocate(by_label, n, floor):
    picked = {}
    for label, keys in by_label.items():
        picked[label] = min(floor, len(keys))
    spare = n - sum(picked.values())
    if spare > 0:
        weights = {l: len(k) for l, k in by_label.items()}
        total = sum(weights.values())
        order = sorted(weights, key=lambda l: -weights[l])
        for label in order:
            if spare <= 0:
                break
            room = len(by_label[label]) - picked[label]
            want = min(room, spare, max(1, round(spare * weights[label] / total)))
            picked[label] += want
            spare -= want
        while spare > 0:
            moved = False
            for label in order:
                if spare <= 0:
                    break
                if picked[label] < len(by_label[label]):
                    picked[label] += 1
                    spare -= 1
                    moved = True
            if not moved:
                break
    return {l: c for l, c in picked.items() if c}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--reasons", default=REASONS)
    ap.add_argument("--labels", default=LABELS)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--n", type=int, default=50)
    ap.add_argument("--floor", type=int, default=2)
    ap.add_argument("--seed", type=int, default=20261113)
    args = ap.parse_args()

    try:
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.worksheet.datavalidation import DataValidation
    except ImportError:
        sys.exit("pip install openpyxl")

    reasons = load(args.reasons)
    labels = load(args.labels)
    pool = candidates(reasons, labels)
    if not pool:
        sys.exit("No classified sentences found. Run classify_reasons.py first.")

    by_label = defaultdict(list)
    for key, occurrences in pool.items():
        by_label[occurrences[0][3]].append(key)

    quota = allocate(by_label, args.n, args.floor)
    random.seed(args.seed)

    rows = []
    for label in sorted(quota):
        keys = sorted(by_label[label])
        for key in random.sample(keys, quota[label]):
            aid, page, text, _got = random.choice(sorted(pool[key]))
            rows.append((key, aid, page, text))
    random.shuffle(rows)

    codes = sorted(CATEGORIES) + ["UNSURE"]
    wb = openpyxl.Workbook()

    ref = wb.active
    ref.title = "Categories"
    ref.append(["Code", "What it means", "RWE could address"])
    for code in sorted(CATEGORIES):
        text, rwe = CATEGORIES[code]
        ref.append([code, text, rwe])
    ref.append(["UNSURE", "You cannot tell from this sentence alone. "
                          "Use it freely -- it is excluded from scoring, "
                          "not counted against the model.", "-"])
    ref.column_dimensions["A"].width = 34
    ref.column_dimensions["B"].width = 86
    ref.column_dimensions["C"].width = 18
    for cell in ref[1]:
        cell.font = Font(bold=True)
    for row in ref.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")

    rules = wb.create_sheet("Rules")
    rules.append(["The model was given exactly these rules. Code by them, "
                  "not by your own reading of the categories."])
    rules.append([])
    for line in SYSTEM.split("CATEGORIES")[0].strip().splitlines():
        rules.append([line])
    rules.column_dimensions["A"].width = 100
    rules["A1"].font = Font(bold=True)

    sheet = wb.create_sheet("Coding", 0)
    sheet.append(["Item", "Ref", "Appraisal", "Page", "Sentence",
                  "Your_Category", "Notes"])
    for i, (key, aid, page, text) in enumerate(rows, 1):
        sheet.append([i, key, aid, page, text, "", ""])

    widths = {"A": 6, "B": 18, "C": 11, "D": 7, "E": 96, "F": 34, "G": 40}
    for col, width in widths.items():
        sheet.column_dimensions[col].width = width
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
    for row in sheet.iter_rows(min_row=2):
        row[4].alignment = Alignment(wrap_text=True, vertical="top")
        row[6].alignment = Alignment(wrap_text=True, vertical="top")
    sheet.freeze_panes = "A2"

    last = len(codes) + 1
    dv = DataValidation(type="list",
                        formula1=f"=Categories!$A$2:$A${last}",
                        allow_blank=True, showDropDown=False)
    dv.error = "Pick a code from the list, or UNSURE."
    dv.errorTitle = "Not a category"
    sheet.add_data_validation(dv)
    dv.add(f"F2:F{len(rows) + 1}")

    wb.save(args.out)

    spread = Counter(
        labels[key] for key, _a, _p, _t in rows)
    print(f"{len(pool):,} distinct classified sentences available")
    print(f"sampled {len(rows)} of them, seed {args.seed}, "
          f"floor {args.floor} per category\n")
    for label in sorted(spread, key=lambda l: -spread[l]):
        have = len(by_label[label])
        print(f"  {spread[label]:>3} of {have:>4}   {label}")
    missing = sorted(set(CATEGORIES) - set(spread))
    if missing:
        print(f"\nNot represented (nothing classified as these): "
              f"{', '.join(missing)}")
    print(f"\n-> {args.out}")
    print("\nThe model's answers are NOT in this file. Column F is yours. "
          "Code all\nof them, then run score_reason_benchmark.py, which "
          "looks the predictions\nup by Ref at scoring time.")
    print("\nRead the Rules sheet first. The table benchmark lost points to "
          "definitional\ndisagreement rather than error, and that is "
          "avoidable here.")


if __name__ == "__main__":
    main()
