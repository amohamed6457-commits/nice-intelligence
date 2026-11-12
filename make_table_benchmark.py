"""
Build a blind benchmark for the table-to-category classifier.

What is being measured
----------------------
endpoint_tables.py sorts recovered tables into Mortality / Morbidity /
Quality of life / Adverse events. The audit says 73.9% of rows land in SOME
category. Nobody has checked whether they land in the RIGHT one.

That matters more than coverage now, because those tables render inside the
expanders on the record view, directly under a verdict. A table filed under
the wrong category there is worse than a table not shown at all: it looks
authoritative and it is wrong.

How the sample is drawn
-----------------------
Stratified random, fixed seed, from the PREDICTED categories. Stratifying by
prediction measures precision -- "when the tool says Quality of life, is it
right?" -- which is the question that matters for the expanders. Measuring
recall would need every section of every table coded, which is days of work.

The workbook deliberately does NOT show what the tool predicted. If it did,
the coding would anchor on it and the result would be worthless. The scorer
recomputes predictions from gba_tables.json, so there is no answer key to
avoid opening.

    python make_table_benchmark.py
    python make_table_benchmark.py --n 15 --seed 20261112

Writes table_benchmark.xlsx. Code the 'true category' column from the
resolution, then run score_table_benchmark.py.
"""

import argparse
import json
import os
import random
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import endpoint_tables
except ImportError as exc:
    sys.exit(f"endpoint_tables.py must sit beside this script ({exc}).")

CATEGORIES = ["Mortality", "Morbidity", "Quality of life", "Adverse events"]
UNPLACED = "(not placed)"

COLUMNS = [
    ("appraisal_id", 14), ("page", 7), ("section heading", 30),
    ("first rows of the table", 62),
    ("true category", 20), ("is this an endpoint table at all?", 22),
    ("notes", 26),
]


def collect(tables):
    """Every section, with what the classifier decided (kept for scoring)."""
    out = []
    for aid, group in tables.items():
        for index, table in enumerate(group):
            if not endpoint_tables.is_endpoint_table(table):
                out.append({"appraisal_id": aid, "table": index,
                            "page": table.get("page"),
                            "heading": "(excluded as cost/regimen table)",
                            "rows": (table.get("rows") or [])[:3],
                            "predicted": UNPLACED, "excluded": True})
                continue
            for heading, category, rows in endpoint_tables.sections(table):
                if not rows:
                    continue
                out.append({"appraisal_id": aid, "table": index,
                            "page": table.get("page"),
                            "heading": heading or "(no heading)",
                            "rows": rows[:3],
                            "predicted": category or UNPLACED,
                            "excluded": False})
    return out


def rows_as_text(rows):
    lines = []
    for row in rows:
        cells = [str(c or "").strip() for c in row]
        cells = [c for c in cells if c]
        lines.append(" | ".join(cells)[:150])
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tables", default=endpoint_tables.TABLES_FILE)
    ap.add_argument("--n", type=int, default=10,
                    help="sections per predicted category (default 10)")
    ap.add_argument("--seed", type=int, default=20261112)
    ap.add_argument("--out", default="table_benchmark.xlsx")
    args = ap.parse_args()

    if not os.path.exists(args.tables):
        sys.exit(f"{args.tables} not found. Run this beside it.")
    with open(args.tables, encoding="utf-8") as fh:
        tables = json.load(fh)

    sections = collect(tables)
    print(f"{len(tables):,} appraisals, {len(sections):,} sections\n")
    buckets = {}
    for section in sections:
        buckets.setdefault(section["predicted"], []).append(section)
    for name in CATEGORIES + [UNPLACED]:
        print(f"  {name:<18} {len(buckets.get(name, [])):>6,} sections")

    rng = random.Random(args.seed)
    picked = []
    for name in CATEGORIES + [UNPLACED]:
        pool = buckets.get(name, [])
        picked += rng.sample(pool, min(args.n, len(pool)))
    rng.shuffle(picked)          # so the sheet is not grouped by prediction

    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    HEAD = PatternFill("solid", fgColor="1F3864")
    INPUT = PatternFill("solid", fgColor="FFF2CC")
    THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)
    A = "Arial"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Coding"
    ws.sheet_view.showGridLines = False
    title = ws.cell(1, 1, "Which endpoint category does each table belong to?")
    title.font = Font(name=A, size=12, bold=True, color="1F3864")
    ws.merge_cells(start_row=1, start_column=1, end_row=1,
                   end_column=len(COLUMNS))
    for i, (name, width) in enumerate(COLUMNS, 1):
        c = ws.cell(2, i, name)
        c.font = Font(name=A, size=10, bold=True, color="FFFFFF")
        c.fill = HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center",
                                horizontal="center")
        c.border = THIN
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[2].height = 40
    ws.freeze_panes = "A3"

    for r, section in enumerate(picked, 3):
        values = [section["appraisal_id"], section["page"],
                  section["heading"], rows_as_text(section["rows"]), "", "", ""]
        for i, value in enumerate(values, 1):
            c = ws.cell(r, i, value)
            c.font = Font(name=A, size=9)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = THIN
            if i >= 5:
                c.fill = INPUT
        ws.row_dimensions[r].height = 46

    last = len(picked) + 2
    dv = DataValidation(
        type="list",
        formula1='"Mortality,Morbidity,Quality of life,Adverse events,'
                 'Not an endpoint table,Cannot tell"',
        allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"E3:E{last}")
    dv2 = DataValidation(type="list", formula1='"yes,no"', allow_blank=True)
    ws.add_data_validation(dv2)
    dv2.add(f"F3:F{last}")

    notes = wb.create_sheet("How to code")
    notes.sheet_view.showGridLines = False
    notes.column_dimensions["A"].width = 3
    notes.column_dimensions["B"].width = 100
    guide = [
        ("HOW TO CODE THIS", 14, True, "1F3864"),
        ("", 11, False, "000000"),
        ("For each row, decide which endpoint category the table belongs to, "
         "using the heading\nand the first rows shown. Open the resolution if "
         "you need to - appraisal id and page\nare given.", 11, False, "000000"),
        ("", 11, False, "000000"),
        ("THE RULE THAT MATTERS", 12, True, "1F3864"),
        ("The heading above the table decides the category, not the "
         "instrument. EORTC QLQ-C30\nSYMPTOM scales (fatigue, nausea, pain, "
         "insomnia) are filed under Morbidity. Its\nFUNCTIONAL scales "
         "(physical, role, social, emotional, cognitive functioning) are "
         "filed\nunder Quality of life. Same questionnaire, two categories.",
         11, False, "000000"),
        ("", 11, False, "000000"),
        ("'Not an endpoint table' means a treatment cost table, a comparator "
         "table, or a\nregimen table - rows that are drug names and doses "
         "rather than outcomes.", 11, False, "000000"),
        ("", 11, False, "000000"),
        ("'Cannot tell' is a legitimate answer. Use it rather than guessing - "
         "a guess scores\nthe coder, not the tool.", 11, False, "000000"),
        ("", 11, False, "000000"),
        ("The sheet does NOT show what the tool predicted, on purpose. "
         "Seeing the prediction\nwould anchor the coding and the result would "
         "be worthless.", 11, True, "C00000"),
    ]
    r = 2
    for text, size, bold, colour in guide:
        c = notes.cell(r, 2, text)
        c.font = Font(name=A, size=size, bold=bold, color=colour)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        notes.row_dimensions[r].height = 15 * (1 + text.count("\n"))
        r += 1

    wb.save(args.out)
    print(f"\n{len(picked)} sections drawn (seed {args.seed})")
    print(f"-> {args.out}")
    print("\nCode the 'true category' column, then run "
          "score_table_benchmark.py.")


if __name__ == "__main__":
    main()
