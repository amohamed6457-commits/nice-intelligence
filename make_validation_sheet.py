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
LABELS = "reason_categories_v6.json"
PROGRESS = "fetch_iqwig_progress.json"
OUT = "validation_sheet.xlsx"

CRITICAL = ["NO_COMPARATIVE_DATA", "POPULATION_MISMATCH", "NOTHING_SUBMITTED",
            "INSUFFICIENT_FOLLOW_UP", "NOT_A_REASON", "REASON_NOT_STATED"]

WHERE = [
    ("IN_THE_SENTENCE", "The reason is in the Sentence column itself."),
    ("IN_THE_CONTEXT", "The reason is in Context_Before or Context_After. "
                       "The model saw these, so this still counts against "
                       "it."),
    ("ELSEWHERE_IN_PDF", "You had to open the document to find it. The "
                         "model never saw it -- this is an extraction "
                         "failure, not a classifier error, and comes out "
                         "of its score."),
    ("NOWHERE", "No reason anywhere you looked. REASON_NOT_STATED is the "
                "correct answer."),
]


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


def key_of(text):
    return hashlib.sha1(text.strip().encode("utf-8")).hexdigest()[:16]


def wilson(hits, n, z=1.96):
    if not n:
        return 0.0, 0.0
    p = hits / n
    d = 1 + z * z / n
    centre = (p + z * z / (2 * n)) / d
    half = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return max(0.0, centre - half), min(1.0, centre + half)


def family(code):
    return FAMILY.get(code, code)


def load(path):
    if not os.path.exists(path):
        sys.exit(f"{path} not found.")
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def normalise(text):
    import re
    return " ".join(re.findall(r"[a-zA-ZÀ-ɏ]+",
                               str(text or "").lower()))


def pool(reasons, labels, wanted):
    seen = defaultdict(list)
    for aid, entries in reasons.items():
        for entry in entries:
            text = str(entry.get("sentence") or "").strip()
            if len(text) < 40:
                continue
            code = labels.get(key_of(text))
            if code not in wanted:
                continue
            seen[key_of(text)].append((
                aid, entry.get("page"), text,
                str(entry.get("before") or "").strip(),
                str(entry.get("after") or "").strip(), code))
    return seen


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
    wanted = set(args.categories.split(",")) if args.categories else set(CRITICAL)
    found = pool(reasons, labels, wanted)
    if not found:
        sys.exit(f"Nothing labelled {sorted(wanted)} in {args.labels}.")

    by_code = defaultdict(list)
    for key, rows in found.items():
        by_code[rows[0][5]].append(key)

    per = max(1, args.n // max(1, len(by_code)))
    random.seed(args.seed)
    picked = []
    for code in sorted(by_code):
        keys = sorted(by_code[code])
        for key in random.sample(keys, min(per, len(keys))):
            picked.append(random.choice(sorted(found[key])))
    random.shuffle(picked)

    where = provenance(args.progress)
    to_do = []
    for a, p, t, b, x, _c in picked:
        for piece in (t, b, x):
            if piece:
                to_do.append((key_of(piece), a, p, piece))
    english = ({} if args.no_english
               else englished(to_do, args.translate_model, args.workers))

    wb = openpyxl.Workbook()

    ref = wb.active
    ref.title = "Categories"
    ref.append(["Code", "What it means", "RWE could address"])
    for code in sorted(CATEGORIES):
        text, rwe = CATEGORIES[code]
        ref.append([code, text, rwe])
    ref.append(["UNSURE", "You cannot tell. Excluded from scoring.", "-"])
    ref.column_dimensions["A"].width = 32
    ref.column_dimensions["B"].width = 88
    ref.column_dimensions["C"].width = 18
    for cell in ref[1]:
        cell.font = Font(bold=True)
    for row in ref.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")

    wsheet = wb.create_sheet("Where")
    wsheet.append(["Answer", "What it means"])
    for code, text in WHERE:
        wsheet.append([code, text])
    wsheet.column_dimensions["A"].width = 22
    wsheet.column_dimensions["B"].width = 92
    for cell in wsheet[1]:
        cell.font = Font(bold=True)
    for row in wsheet.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")

    howto = wb.create_sheet("How to code this")
    for line in [
        "Three columns to fill per row. The third is what makes this work.",
        "",
        "Your_Category  -- which of the 16 codes the SENTENCE should have.",
        "Where_Found    -- where you found the reason (see Where sheet).",
        "Quote          -- paste the exact words that give the reason.",
        "",
        "Open the document if you want to. That is not cheating here:",
        "Where_Found records it, and a reason you had to dig for is scored",
        "as an extraction failure rather than held against the classifier.",
        "",
        "The scorer checks your Quote against the Sentence and Context",
        "columns. If you say IN_THE_SENTENCE or IN_THE_CONTEXT and the quote",
        "is not there, it reclassifies the row rather than scoring it. So",
        "paste the words as they appear, do not paraphrase.",
        "",
        "Comparator codes are scored as one family, so choosing",
        "COMPARATOR_NOT_IMPLEMENTED where the model said",
        "INADEQUATE_THERAPY_CONTINUED is not counted as a disagreement.",
        "",
        "UNSURE is free and excluded. Use it rather than guessing.",
    ]:
        howto.append([line])
    howto.column_dimensions["A"].width = 76
    howto["A1"].font = Font(bold=True)

    sheet = wb.create_sheet("Coding", 0)
    sheet.append(["Item", "Ref", "Appraisal", "Document", "Page", "Link",
                  "Context_Before_EN", "Sentence_English", "Sentence_German",
                  "Context_After_EN", "Your_Category", "Where_Found",
                  "Quote", "Notes"])
    for i, (aid, page, text, before, after, _code) in enumerate(picked, 1):
        doc, url = where.get(aid, ("unknown", ""))
        sheet.append([i, key_of(text), aid, doc, page, "",
                      english.get(before, before),
                      english.get(text, ""), text,
                      english.get(after, after), "", "", "", ""])
        cell = sheet.cell(row=i + 1, column=6)
        if url:
            cell.value = "open"
            cell.hyperlink = f"{url}#page={page}" if page else url
            cell.style = "Hyperlink"

    widths = {"A": 6, "B": 18, "C": 11, "D": 21, "E": 7, "F": 7, "G": 50,
              "H": 70, "I": 70, "J": 50, "K": 31, "L": 20, "M": 46, "N": 34}
    for col, width in widths.items():
        sheet.column_dimensions[col].width = width
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
    for row in sheet.iter_rows(min_row=2):
        for idx in (6, 7, 8, 9, 12, 13):
            row[idx].alignment = Alignment(wrap_text=True, vertical="top")
    sheet.freeze_panes = "H2"

    last = len(CATEGORIES) + 2
    dv = DataValidation(type="list",
                        formula1=f"=Categories!$A$2:$A${last}",
                        allow_blank=True, showDropDown=False)
    dv.error = "Pick a code from the Categories sheet, or UNSURE."
    dv.errorTitle = "Not a category"
    sheet.add_data_validation(dv)
    dv.add(f"K2:K{len(picked) + 1}")

    dw = DataValidation(type="list",
                        formula1=f"=Where!$A$2:$A${len(WHERE) + 1}",
                        allow_blank=True, showDropDown=False)
    dw.error = "Pick one of the four answers on the Where sheet."
    dw.errorTitle = "Not an answer"
    sheet.add_data_validation(dw)
    dw.add(f"L2:L{len(picked) + 1}")

    wb.save(args.out)
    spread = Counter(c for _a, _p, _t, _b, _x, c in picked)
    print(f"{len(found):,} distinct sentences in the chosen categories")
    print(f"sampled {len(picked)}, seed {args.seed}\n")
    for code, n in spread.most_common():
        print(f"  {n:>3} of {len(by_code[code]):>5}   {code}")
    print(f"\n-> {args.out}")
    print("\nColumns K, L and M are yours. Read 'How to code this' first.")


def score(args):
    try:
        import openpyxl
    except ImportError:
        sys.exit("pip install openpyxl")
    if args.out == OUT or not os.path.exists(args.out):
        args.out = newest(args.out)
    for path in (args.out, args.labels):
        if not os.path.exists(path):
            sys.exit(f"{path} not found.")
    print(f"scoring {args.out}\n")
    labels = load(args.labels)
    sheet = openpyxl.load_workbook(args.out, data_only=True)["Coding"]
    hdr = [str(c.value or "").strip() for c in sheet[1]]
    ci = {n: hdr.index(n) for n in hdr if n}

    scoreable, extraction, unsure, blank, unquoted = [], [], 0, 0, []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        ref = str(row[ci["Ref"]] or "").strip()
        if not ref:
            continue
        mine = str(row[ci["Your_Category"]] or "").strip().upper()
        found_in = str(row[ci["Where_Found"]] or "").strip().upper()
        quote = str(row[ci["Quote"]] or "").strip()
        got = labels.get(ref)
        note = str(row[ci["Notes"]] or "") if "Notes" in ci else ""
        alts = {c for c in CATEGORIES
                if c in note.upper().replace(" ", "_")} - {mine}
        item = (row[ci["Item"]], row[ci["Appraisal"]], row[ci["Page"]],
                str(row[ci["Sentence_English"]] or ""), mine, got, quote,
                note, alts)
        if not mine or not found_in:
            blank += 1
            continue
        if mine == "UNSURE":
            unsure += 1
            continue
        if mine not in CATEGORIES:
            sys.exit(f"Item {item[0]}: '{mine}' is not a category code.")
        if not got or str(got).startswith("__"):
            continue
        if found_in == "ELSEWHERE_IN_PDF":
            extraction.append(item)
            continue
        if found_in in ("IN_THE_SENTENCE", "IN_THE_CONTEXT") and quote:
            shown = normalise(" ".join(
                str(row[ci[c]] or "") for c in hdr
                if c.startswith(("Sentence", "Context"))))
            words = normalise(quote).split()
            probe = " ".join(words[:8])
            if probe and probe not in shown:
                unquoted.append(item)
        scoreable.append(item)

    n = len(scoreable)
    print(f"scoreable                  : {n}")
    print(f"reason only in the document : {len(extraction)}  "
          f"(extraction failure, excluded)")
    if unquoted:
        print(f"quote not verifiable       : {len(unquoted)}  (still "
              f"scored -- see note)")
    if unsure:
        print(f"UNSURE                     : {unsure}  (excluded)")
    if blank:
        print(f"incomplete rows            : {blank}")
    if not n:
        sys.exit("\nNothing scoreable.")

    if unquoted:
        print(f"\nNOTE: {len(unquoted)} quote(s) could not be matched "
              f"against the shown text.\nThey are scored anyway, trusting "
              f"your Where_Found answers. If any of those\nreasons came "
              f"from the document rather than the columns, the figure\n"
              f"below is optimistic.")
    agree = sum(1 for *_x, mine, got, _q, _n, _al in scoreable
                if family(mine) == family(got))
    lo, hi = wilson(agree, n)
    print(f"\nCLASSIFIER ACCURACY  {agree}/{n} = {agree / n:.1%}"
          f"   95% CI {lo:.1%} to {hi:.1%}")
    print("(comparator codes scored as one family)")

    per = defaultdict(lambda: [0, 0])
    for *_x, mine, got, _q, _n, _al in scoreable:
        per[got][1] += 1
        if family(mine) == family(got):
            per[got][0] += 1
    print("\nWhere the model said X, how often you agreed:")
    for code in sorted(per, key=lambda c: -per[c][1]):
        hits, total = per[code]
        klo, khi = wilson(hits, total)
        print(f"  {code:<31} {hits:>3}/{total:<3} {hits / total:>5.0%}"
              f"   [{klo:.0%}-{khi:.0%}]")

    lenient = sum(1 for *_x, mine, got, _q, _n, alts in scoreable
                  if family(mine) == family(got)
                  or any(family(a) == family(got) for a in alts))
    noted = sum(1 for *_x, mine, got, _q, _n, alts in scoreable
                if family(mine) != family(got)
                and any(family(a) == family(got) for a in alts))
    if noted:
        llo, lhi = wilson(lenient, n)
        print(f"\nAllowing alternatives you listed in Notes: "
              f"{lenient}/{n} = {lenient / n:.1%}"
              f"   95% CI {llo:.1%} to {lhi:.1%}")
        print(f"({noted} disagreement(s) had the model's answer among your "
              f"noted alternatives)")

    pairs = Counter((got, mine) for *_x, mine, got, _q, _n, _al in scoreable
                    if family(mine) != family(got))
    if pairs:
        print(f"\n{sum(pairs.values())} disagreement(s):")
        for (got, mine), count in pairs.most_common():
            print(f"  {count:>3}x  model {got}  ->  you {mine}")

    if args.show:
        for item, aid, page, text, mine, got, quote, note, alts in scoreable:
            if family(mine) == family(got):
                continue
            print(f"\nitem {item}  {aid} p{page}")
            print(f"  {' '.join(text[:220].split())}")
            print(f"  model: {got}")
            print(f"  you  : {mine}")
            if alts:
                mark = " <- model's answer is in here" if any(
                    family(a) == family(got) for a in alts) else ""
                print(f"  your alternatives: {', '.join(sorted(alts))}{mark}")
            if quote:
                print(f"  quote: {' '.join(quote[:150].split())}")
            if note:
                print(f"  notes: {' '.join(note[:170].split())}")

    print("\n" + "=" * 70)
    if n < 25:
        print(f"{n} scoreable items is too few to quote.")
    elif lo >= 0.80:
        print(f"Lower bound {lo:.0%}. Holds at 80% even pessimistically "
              f"-- safe to put on screen with the figure stated.")
    elif agree / n >= 0.80:
        print(f"Point estimate {agree / n:.0%} but lower bound {lo:.0%}. "
              f"Quote the interval,\nnot the point estimate.")
    else:
        print(f"{agree / n:.0%}. Below the bar. Read the disagreements "
              f"before shipping any of it.")
    if extraction:
        print(f"\n{len(extraction)} of {len(extraction) + n} reasons were "
              f"only findable in the document.\nThat is the extraction "
              f"window, measured: "
              f"{len(extraction) / (len(extraction) + n):.0%} of reasons "
              f"sit outside it.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--reasons", default=REASONS)
    ap.add_argument("--labels", default=LABELS)
    ap.add_argument("--progress", default=PROGRESS)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--categories")
    ap.add_argument("--n", type=int, default=48)
    ap.add_argument("--seed", type=int, default=20261114)
    ap.add_argument("--no-english", action="store_true")
    ap.add_argument("--translate-model", default="claude-haiku-4-5-20251001")
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--score", action="store_true")
    ap.add_argument("--show", action="store_true")
    args = ap.parse_args()
    score(args) if args.score else build(args)


if __name__ == "__main__":
    main()
