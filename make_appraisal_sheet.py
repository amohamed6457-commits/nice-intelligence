import argparse
import hashlib
import json
import math
import os
import random
import sys
from collections import Counter, defaultdict

from classify_reasons import CATEGORIES, FAMILY
from make_reason_benchmark import englished, provenance

REASONS = "reasons.json"
LABELS = "reason_categories_v7.json"
PROGRESS = "fetch_iqwig_progress.json"
OUT = "appraisal_sheet.xlsx"

NOT_REASONS = {"NOT_A_REASON", "REASON_NOT_STATED"}

SHOWN = {"comparator": "COMPARATOR_NOT_IMPLEMENTED (any of the three)"}


def show(code):
    return SHOWN.get(code, code)


def key_of(text):
    return hashlib.sha1(text.strip().encode("utf-8")).hexdigest()[:16]


def family(code):
    return FAMILY.get(code, code)


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
    if not os.path.exists(path):
        sys.exit(f"{path} not found.")
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def by_appraisal(reasons, labels):
    out = {}
    for aid, entries in reasons.items():
        rows, codes = [], set()
        for entry in entries:
            text = str(entry.get("sentence") or "").strip()
            if len(text) < 40:
                continue
            code = labels.get(key_of(text))
            if not code or str(code).startswith("__"):
                continue
            rows.append((entry.get("page"), text, code))
            if code not in NOT_REASONS:
                codes.add(family(code))
        if rows:
            out[aid] = (rows, codes)
    return out


def parse_set(raw):
    import re
    found, unknown = set(), []
    for segment in re.split(r"[,;/\n]+|\band\b|\+", str(raw or ""),
                            flags=re.IGNORECASE):
        token = re.sub(r"[^A-Z]+", "_", segment.upper()).strip("_")
        if not token:
            continue
        if token in CATEGORIES:
            found.add(family(token))
            continue
        hits = [c for c in CATEGORIES if c in token or token in c]
        if hits:
            found.add(family(max(hits, key=len)))
        else:
            unknown.append(segment.strip())
    return found, unknown


def build(args):
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.worksheet.datavalidation import DataValidation
    except ImportError:
        sys.exit("pip install openpyxl")

    args.out = next_free(args.out)
    reasons = load(args.reasons)
    labels = load(args.labels)
    grouped = by_appraisal(reasons, labels)
    if not grouped:
        sys.exit(f"No classified appraisals in {args.labels}.")

    eligible = [a for a, (rows, _c) in grouped.items()
                if args.min_sentences <= len(rows) <= args.max_sentences]
    if not eligible:
        sys.exit(f"No appraisal has between {args.min_sentences} and "
                 f"{args.max_sentences} sentences.")
    random.seed(args.seed)
    picked = sorted(random.sample(sorted(eligible),
                                 min(args.n, len(eligible))))

    where = provenance(args.progress)
    to_do = []
    for aid in picked:
        for page, text, _code in grouped[aid][0]:
            to_do.append((key_of(text), aid, page, text))
    english = ({} if args.no_english
               else englished(to_do, args.translate_model, args.workers))

    wb = openpyxl.Workbook()

    ref = wb.active
    ref.title = "Categories"
    ref.append(["Code", "What it means", "RWE could address"])
    for code in sorted(CATEGORIES):
        if code in NOT_REASONS:
            continue
        text, rwe = CATEGORIES[code]
        ref.append([code, text, rwe])
    ref.column_dimensions["A"].width = 32
    ref.column_dimensions["B"].width = 88
    ref.column_dimensions["C"].width = 18
    for cell in ref[1]:
        cell.font = Font(bold=True)
    for row in ref.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")

    howto = wb.create_sheet("How to code this")
    for line in [
        "One row per appraisal. The question is:",
        "",
        "   Why did the G-BA find no added benefit in this appraisal?",
        "",
        "OPEN THE DOCUMENT. That is the method here, not a shortcut.",
        "The Link column goes straight to it. The Evidence sheet shows the",
        "sentences the pipeline extracted for that appraisal, as a starting",
        "point -- it does NOT show what the pipeline concluded from them.",
        "",
        "In Reasons_You_Found, list EVERY reason that applies, separated by",
        "commas. Most appraisals have more than one.",
        "",
        "   NO_COMPARATIVE_DATA, INSUFFICIENT_FOLLOW_UP",
        "",
        "Use the codes on the Categories sheet. Spelling must match; the",
        "scorer will tell you about anything it cannot read.",
        "",
        "If you cannot work out why it failed, write UNCLEAR and say what",
        "stopped you in Notes.",
        "",
        "The three comparator codes count as one, so do not agonise over",
        "which of them to pick.",
        "",
        "You will not see the pipeline's answer anywhere in this workbook.",
        "Scoring compares your set against it afterwards.",
    ]:
        howto.append([line])
    howto.column_dimensions["A"].width = 74
    howto["A1"].font = Font(bold=True)

    ev = wb.create_sheet("Evidence")
    ev.append(["Appraisal", "Page", "Link", "Sentence_English",
               "Sentence_German"])
    r = 2
    for aid in picked:
        doc, url = where.get(aid, ("unknown", ""))
        for page, text, _code in sorted(grouped[aid][0],
                                        key=lambda x: (x[0] or 0)):
            ev.append([aid, page, "", english.get(text, ""), text])
            cell = ev.cell(row=r, column=3)
            if url:
                cell.value = f"p{page}"
                cell.hyperlink = f"{url}#page={page}" if page else url
                cell.style = "Hyperlink"
            r += 1
    for col, width in {"A": 11, "B": 7, "C": 8, "D": 80, "E": 80}.items():
        ev.column_dimensions[col].width = width
    for cell in ev[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
    for row in ev.iter_rows(min_row=2):
        row[3].alignment = Alignment(wrap_text=True, vertical="top")
        row[4].alignment = Alignment(wrap_text=True, vertical="top")
    ev.freeze_panes = "A2"

    sheet = wb.create_sheet("Coding", 0)
    sheet.append(["Appraisal", "Document", "Link", "Sentences_Extracted",
                  "Reasons_You_Found", "Confident", "Notes"])
    for i, aid in enumerate(picked, 1):
        doc, url = where.get(aid, ("unknown", ""))
        rows, _codes = grouped[aid]
        sheet.append([aid, doc, "", len(rows), "", "", ""])
        cell = sheet.cell(row=i + 1, column=3)
        if url:
            cell.value = "open"
            cell.hyperlink = url
            cell.style = "Hyperlink"
    for col, width in {"A": 12, "B": 22, "C": 8, "D": 20, "E": 54,
                       "F": 12, "G": 60}.items():
        sheet.column_dimensions[col].width = width
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
    for row in sheet.iter_rows(min_row=2):
        row[4].alignment = Alignment(wrap_text=True, vertical="top")
        row[6].alignment = Alignment(wrap_text=True, vertical="top")
    sheet.freeze_panes = "A2"

    dv = DataValidation(type="list", formula1='"yes,no"',
                        allow_blank=True, showDropDown=False)
    sheet.add_data_validation(dv)
    dv.add(f"F2:F{len(picked) + 1}")

    wb.save(args.out)
    counts = [len(grouped[a][0]) for a in picked]
    print(f"{len(grouped):,} appraisals classified, {len(eligible):,} with "
          f"{args.min_sentences}-{args.max_sentences} sentences")
    print(f"sampled {len(picked)}, seed {args.seed}")
    print(f"sentences per appraisal: min {min(counts)}, median "
          f"{sorted(counts)[len(counts) // 2]}, max {max(counts)}")
    print(f"total sentences on the Evidence sheet: {sum(counts)}")
    print(f"\n-> {args.out}")
    print("\nColumn E is yours, comma-separated. Read 'How to code this'.")
    print("The pipeline's answers are NOT in this workbook.")


def score(args):
    try:
        import openpyxl
    except ImportError:
        sys.exit("pip install openpyxl")
    if args.out == OUT or not os.path.exists(args.out):
        args.out = newest(args.out)
    for path in (args.out, args.labels, args.reasons):
        if not os.path.exists(path):
            sys.exit(f"{path} not found.")
    print(f"scoring {args.out}\n")
    grouped = by_appraisal(load(args.reasons), load(args.labels))
    sheet = openpyxl.load_workbook(args.out, data_only=True)["Coding"]
    hdr = [str(c.value or "").strip() for c in sheet[1]]
    ci = {n: hdr.index(n) for n in hdr if n}

    rows, unclear, blank, unknown_tokens = [], [], 0, Counter()
    for row in sheet.iter_rows(min_row=2, values_only=True):
        aid = str(row[ci["Appraisal"]] or "").strip()
        if not aid:
            continue
        raw = str(row[ci["Reasons_You_Found"]] or "").strip()
        note = str(row[ci["Notes"]] or "").strip()
        if not raw:
            blank += 1
            continue
        if "UNCLEAR" in raw.upper():
            unclear.append((aid, note))
            continue
        mine, bad = parse_set(raw)
        for token in bad:
            unknown_tokens[token] += 1
        if not mine:
            print(f"  ! {aid}: could not read any category from {raw!r}")
            continue
        theirs = grouped.get(aid, ([], set()))[1]
        rows.append((aid, mine, theirs, note))

    print(f"appraisals coded : {len(rows)}")
    if unclear:
        print(f"marked UNCLEAR   : {len(unclear)}  (excluded)")
        for aid, note in unclear:
            print(f"    {aid}: {note[:90]}")
    if blank:
        print(f"left blank       : {blank}")
    if unknown_tokens:
        print(f"unreadable words : "
              f"{', '.join(sorted(unknown_tokens))[:200]}")
    if not rows:
        sys.exit("\nNothing to score.")

    exact = sum(1 for _a, m, t, _n in rows if m == t)
    tp = sum(len(m & t) for _a, m, t, _n in rows)
    model_total = sum(len(t) for _a, _m, t, _n in rows)
    mine_total = sum(len(m) for _a, m, _t, _n in rows)
    elo, ehi = wilson(exact, len(rows))

    print(f"\nEXACT SET MATCH   {exact}/{len(rows)} = "
          f"{exact / len(rows):.1%}   95% CI {elo:.1%} to {ehi:.1%}")
    if model_total:
        print(f"PRECISION         {tp}/{model_total} = {tp / model_total:.1%}"
              f"   of the reasons the tool shows, this share is right")
    if mine_total:
        print(f"RECALL            {tp}/{mine_total} = {tp / mine_total:.1%}"
              f"   of the real reasons, this share is found")

    print("\nper appraisal:")
    for aid, mine, theirs, _n in rows:
        missed = sorted(mine - theirs)
        extra = sorted(theirs - mine)
        mark = "ok " if mine == theirs else "   "
        print(f"  {mark} {aid:<10} you {len(mine)}, tool {len(theirs)}, "
              f"agreed {len(mine & theirs)}")
        if missed:
            print(f"        tool missed : {', '.join(show(c) for c in missed)}")
        if extra:
            print(f"        tool invented: {', '.join(show(c) for c in extra)}")

    missed_all = Counter(c for _a, m, t, _n in rows for c in (m - t))
    extra_all = Counter(c for _a, m, t, _n in rows for c in (t - m))
    if missed_all:
        print("\nmost often missed by the tool:")
        for code, n in missed_all.most_common(8):
            print(f"  {n:>3}  {show(code)}")
    if extra_all:
        print("\nmost often invented by the tool:")
        for code, n in extra_all.most_common(8):
            print(f"  {n:>3}  {show(code)}")

    print("\n" + "=" * 70)
    print("This measures what a user sees: the set of reasons shown for an")
    print("appraisal. Precision is how much of it is right; recall is how")
    print("much of the truth it finds. Exact match is the strictest reading.")
    if len(rows) < 8:
        print(f"\n{len(rows)} appraisals is a look, not a measurement.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--reasons", default=REASONS)
    ap.add_argument("--labels", default=LABELS)
    ap.add_argument("--progress", default=PROGRESS)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--n", type=int, default=10)
    ap.add_argument("--min-sentences", type=int, default=3)
    ap.add_argument("--max-sentences", type=int, default=12)
    ap.add_argument("--seed", type=int, default=20261116)
    ap.add_argument("--no-english", action="store_true")
    ap.add_argument("--translate-model", default="claude-haiku-4-5-20251001")
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--score", action="store_true")
    args = ap.parse_args()
    score(args) if args.score else build(args)


if __name__ == "__main__":
    main()
