import argparse
import json
import math
import os
import random
import re
import sys
from collections import Counter

from make_reason_benchmark import englished

SECTIONS = "has_clinical_sections.json"
OUT = "has_benchmark.xlsx"

VERDICTS = [
    ("CORRECT", "The section holds the clinical evidence for this "
                "appraisal -- the studies, the endpoints, the results."),
    ("PARTLY", "Some of it is here but clinical content sits outside this "
               "page range too. Say what is missing in the Missing column."),
    ("WRONG_SECTION", "This is not the clinical evidence. It is some other "
                      "part of the avis."),
    ("NOTHING_CLINICAL", "This avis has no clinical evidence to find -- a "
                         "renewal, a generic, a purely administrative "
                         "opinion."),
    ("UNSURE", "You cannot tell. Excluded from scoring."),
]


def wilson(hits, n, z=1.96):
    if not n:
        return 0.0, 0.0
    p = hits / n
    d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return max(0.0, c - h), min(1.0, c + h)


def next_free(path):
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    n = 2
    while os.path.exists(f"{stem}_{n}{ext}"):
        n += 1
    return f"{stem}_{n}{ext}"


def newest(path):
    import glob
    stem, ext = os.path.splitext(path)
    have = [c for c in [path] + glob.glob(f"{stem}_*{ext}")
            if os.path.exists(c)]
    return max(have, key=os.path.getmtime) if have else path


def load(path):
    for candidate in (path, path + ".sample"):
        if os.path.exists(candidate):
            with open(candidate, encoding="utf-8") as fh:
                return json.load(fh), candidate
    sys.exit(f"{path} not found (nor {path}.sample). Run "
             f"extract_has_sections.py first.")


def clinical_of(doc):
    best = None
    for s in doc.get("sections", []):
        if s["role"] != "clinical_evidence":
            continue
        if best is None or s["chars"] > best["chars"]:
            best = s
    return best


def build(args):
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.worksheet.datavalidation import DataValidation
    except ImportError:
        sys.exit("pip install openpyxl")

    args.out = next_free(args.out)
    data, source = load(args.sections)
    print(f"read {len(data):,} appraisals from {source}")

    eligible = {a: d for a, d in data.items()
                if clinical_of(d) and not d.get("synthese")}
    print(f"with a clinical_evidence section: {len(eligible):,}")
    if args.min_year or args.max_year:
        lo = args.min_year or 0
        hi = args.max_year or 9999
        eligible = {a: d for a, d in eligible.items()
                    if lo <= int(d.get("year") or 0) <= hi}
        print(f"within {lo}-{hi}: {len(eligible):,}")
    if not eligible:
        sys.exit("No appraisal has a clinical_evidence section.")

    random.seed(args.seed)
    picked = sorted(random.sample(sorted(eligible),
                                 min(args.n, len(eligible))))

    previews = []
    for aid in picked:
        text = clinical_of(data[aid])["text"]
        previews.append((aid, aid, 0, text[:args.preview]))
    english = ({} if args.no_english
               else englished(previews, args.translate_model, args.workers))

    wb = openpyxl.Workbook()

    ref = wb.active
    ref.title = "Answers"
    ref.append(["Answer", "What it means"])
    for code, text in VERDICTS:
        ref.append([code, text])
    ref.column_dimensions["A"].width = 20
    ref.column_dimensions["B"].width = 92
    for cell in ref[1]:
        cell.font = Font(bold=True)
    for row in ref.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")

    howto = wb.create_sheet("How to code this")
    for line in [
        "One row per appraisal. Two questions:",
        "",
        "1. Does the extracted section hold the clinical evidence?",
        "2. Is the page range right?",
        "",
        "OPEN THE DOCUMENT. The Link column goes to the first page of the",
        "extracted range. Preview_EN is the opening of that section in",
        "English, just to orient you -- it is truncated, so judge from the",
        "PDF, not the preview.",
        "",
        "Page_Range: say yes if the range covers the clinical evidence and",
        "little else. If it starts or stops in the wrong place, say no and",
        "give the range you would use (for example 'should be 17-30').",
        "",
        "Endpoints: list what is actually measured in that section --",
        "survival, exacerbations, quality of life, whatever it reports. This",
        "is the column that tells us whether endpoint extraction is possible",
        "from these sections at all.",
        "",
        "Missing: clinical content that is in the avis but outside this page",
        "range. That measures whether the boundaries are drawn right.",
        "",
        "NOTHING_CLINICAL is a real answer. Some avis genuinely have no",
        "clinical assessment, and that is a fact about the document.",
        "",
        "The All_Sections sheet shows every section the parser found for",
        "each appraisal, so you can see whether it picked the right one.",
    ]:
        howto.append([line])
    howto.column_dimensions["A"].width = 76
    howto["A1"].font = Font(bold=True)

    allsec = wb.create_sheet("All_Sections")
    allsec.append(["Appraisal", "Number", "Role", "Title", "Page",
                   "Page_End", "Chars"])
    for aid in picked:
        for s in data[aid].get("sections", []):
            allsec.append([aid, s["number"], s["role"], s["title"],
                           s["page"], s["page_end"], s["chars"]])
    for col, width in {"A": 12, "B": 9, "C": 20, "D": 48, "E": 7,
                       "F": 9, "G": 9}.items():
        allsec.column_dimensions[col].width = width
    for cell in allsec[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
    allsec.freeze_panes = "A2"

    sheet = wb.create_sheet("Coding", 0)
    sheet.append(["Appraisal", "Year", "Section", "Title", "Pages", "Chars",
                  "Link", "Preview_EN", "Section_Correct",
                  "Page_Range_OK", "Endpoints", "Missing", "Notes"])
    for i, aid in enumerate(picked, 1):
        doc = data[aid]
        s = clinical_of(doc)
        sheet.append([aid, doc.get("year", ""), s["number"], s["title"],
                      f"{s['page']}-{s['page_end']}", s["chars"], "",
                      english.get(s["text"][:args.preview], ""),
                      "", "", "", "", ""])
        cell = sheet.cell(row=i + 1, column=7)
        url = doc.get("url", "")
        if url:
            cell.value = f"p{s['page']}"
            cell.hyperlink = f"{url}#page={s['page']}"
            cell.style = "Hyperlink"
    for col, width in {"A": 12, "B": 7, "C": 9, "D": 34, "E": 9, "F": 8,
                       "G": 8, "H": 86, "I": 19, "J": 15, "K": 40,
                       "L": 40, "M": 34}.items():
        sheet.column_dimensions[col].width = width
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
    for row in sheet.iter_rows(min_row=2):
        for idx in (3, 7, 10, 11, 12):
            row[idx].alignment = Alignment(wrap_text=True, vertical="top")
    sheet.freeze_panes = "H2"

    dv = DataValidation(type="list",
                        formula1=f"=Answers!$A$2:$A${len(VERDICTS) + 1}",
                        allow_blank=True, showDropDown=False)
    dv.error = "Pick one of the five answers on the Answers sheet."
    sheet.add_data_validation(dv)
    dv.add(f"I2:I{len(picked) + 1}")

    dy = DataValidation(type="list", formula1='"yes,no"',
                        allow_blank=True, showDropDown=False)
    sheet.add_data_validation(dy)
    dy.add(f"J2:J{len(picked) + 1}")

    wb.save(args.out)
    spans = [clinical_of(data[a]) for a in picked]
    years = Counter(int(data[a].get("year") or 0) for a in picked)
    print(f"\nsampled {len(picked)}, seed {args.seed}")
    print(f"years: {', '.join(f'{y}x{n}' for y, n in sorted(years.items()))}")
    print(f"section sizes: median {sorted(s['chars'] for s in spans)[len(spans) // 2]:,} chars, "
          f"max {max(s['chars'] for s in spans):,}")
    print(f"page spans   : median "
          f"{sorted(s['page_end'] - s['page'] for s in spans)[len(spans) // 2]} pages")
    print(f"\n-> {args.out}")
    print("\nColumns I to M are yours. Read 'How to code this' first.")


def score(args):
    try:
        import openpyxl
    except ImportError:
        sys.exit("pip install openpyxl")
    if args.out == OUT or not os.path.exists(args.out):
        args.out = newest(args.out)
    if not os.path.exists(args.out):
        sys.exit(f"{args.out} not found.")
    print(f"scoring {args.out}\n")
    sheet = openpyxl.load_workbook(args.out, data_only=True)["Coding"]
    hdr = [str(c.value or "").strip() for c in sheet[1]]
    ci = {n: hdr.index(n) for n in hdr if n}

    rows, unsure, blank = [], 0, 0
    for row in sheet.iter_rows(min_row=2, values_only=True):
        aid = str(row[ci["Appraisal"]] or "").strip()
        if not aid:
            continue
        verdict = str(row[ci["Section_Correct"]] or "").strip().upper()
        if not verdict:
            blank += 1
            continue
        if verdict == "UNSURE":
            unsure += 1
            continue
        known = {c for c, _t in VERDICTS}
        if verdict not in known:
            sys.exit(f"{aid}: '{verdict}' is not one of {sorted(known)}")
        rows.append((aid, verdict,
                     str(row[ci["Page_Range_OK"]] or "").strip().lower(),
                     str(row[ci["Endpoints"]] or "").strip(),
                     str(row[ci["Missing"]] or "").strip(),
                     str(row[ci["Notes"]] or "").strip()))

    print(f"coded   : {len(rows)}")
    if unsure:
        print(f"UNSURE  : {unsure}  (excluded)")
    if blank:
        print(f"blank   : {blank}")
    if not rows:
        sys.exit("\nNothing to score.")

    counts = Counter(v for _a, v, _p, _e, _m, _n in rows)
    clinical = [r for r in rows if r[1] != "NOTHING_CLINICAL"]
    good = sum(1 for r in clinical if r[1] == "CORRECT")
    usable = sum(1 for r in clinical if r[1] in ("CORRECT", "PARTLY"))

    print("\nverdicts:")
    for v, n in counts.most_common():
        print(f"  {n:>3}  {v}")

    if clinical:
        lo, hi = wilson(good, len(clinical))
        ulo, uhi = wilson(usable, len(clinical))
        print(f"\nAmong the {len(clinical)} avis that do have clinical "
              f"content:")
        print(f"  section fully correct  {good}/{len(clinical)} = "
              f"{good / len(clinical):.0%}   95% CI {lo:.0%} to {hi:.0%}")
        print(f"  correct or partial     {usable}/{len(clinical)} = "
              f"{usable / len(clinical):.0%}   95% CI {ulo:.0%} to {uhi:.0%}")

    pages = [r for r in rows if r[2] in ("yes", "no")]
    if pages:
        ok = sum(1 for r in pages if r[2] == "yes")
        plo, phi = wilson(ok, len(pages))
        print(f"  page range right       {ok}/{len(pages)} = "
              f"{ok / len(pages):.0%}   95% CI {plo:.0%} to {phi:.0%}")

    named = [r for r in rows if r[3]]
    print(f"\nrows where you could list endpoints from the section: "
          f"{len(named)} of {len(rows)}")
    for aid, _v, _p, endpoints, _m, _n in named[:12]:
        print(f"  {aid:<11} {' '.join(endpoints[:88].split())}")

    missing = [r for r in rows if r[4]]
    if missing:
        print(f"\nclinical content outside the extracted range "
              f"({len(missing)} rows):")
        for aid, _v, _p, _e, gap, _n in missing:
            print(f"  {aid:<11} {' '.join(gap[:88].split())}")

    print("\n" + "=" * 70)
    if clinical and usable / len(clinical) >= 0.8:
        print("The sections are usable. Endpoint extraction from them is")
        print("worth building, and quotations can cite a page.")
    elif clinical:
        print("Below 80% usable. Read the Missing column before scaling --")
        print("the boundaries are probably drawn in the wrong place.")
    if len(rows) < 8:
        print(f"\n{len(rows)} appraisals is a look, not a measurement.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sections", default=SECTIONS)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--n", type=int, default=10)
    ap.add_argument("--preview", type=int, default=1200)
    ap.add_argument("--seed", type=int, default=20261116)
    ap.add_argument("--min-year", type=int)
    ap.add_argument("--max-year", type=int)
    ap.add_argument("--no-english", action="store_true")
    ap.add_argument("--translate-model", default="claude-haiku-4-5-20251001")
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--score", action="store_true")
    args = ap.parse_args()
    score(args) if args.score else build(args)


if __name__ == "__main__":
    main()
