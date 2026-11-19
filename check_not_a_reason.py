import argparse
import hashlib
import json
import math
import os
import random
import sys
from collections import Counter, defaultdict

from make_reason_benchmark import englished, provenance

TARGET = "NOT_A_REASON"
REASONS = "reasons.json"
LABELS = "reason_categories_v3.json"
PROGRESS = "fetch_iqwig_progress.json"
OUT = "not_a_reason_check.xlsx"

CHOICES = [
    ("NOT_A_REASON", "Correct. This sentence gives no reason that evidence "
                     "was unusable -- it is a heading, a caption, a scope "
                     "statement, a description, a reported result, or a "
                     "statement of agreement."),
    ("STATES_A_REASON", "Wrong. This sentence DOES state why evidence could "
                        "not be used. Put which reason in Notes."),
    ("UNSURE", "You cannot tell. Excluded from scoring."),
]


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


def load(path):
    if not os.path.exists(path):
        sys.exit(f"{path} not found.")
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def pool_for(reasons, labels, target):
    seen = defaultdict(list)
    for aid, entries in reasons.items():
        for entry in entries:
            text = str(entry.get("sentence") or "").strip()
            if len(text) < 40:
                continue
            if labels.get(key_of(text)) != target:
                continue
            seen[key_of(text)].append((
                aid, entry.get("page"), text,
                str(entry.get("before") or "").strip(),
                str(entry.get("after") or "").strip()))
    return seen


def build(args):
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.worksheet.datavalidation import DataValidation
    except ImportError:
        sys.exit("pip install openpyxl")

    reasons = load(args.reasons)
    labels = load(args.labels)
    pool = pool_for(reasons, labels, args.target)
    if not pool:
        sys.exit(f"Nothing labelled {args.target} in {args.labels}.")
    print(f"{len(pool):,} distinct sentences labelled {args.target}")

    random.seed(args.seed)
    keys = random.sample(sorted(pool), min(args.n, len(pool)))
    rows = [random.choice(sorted(pool[k])) for k in keys]
    random.shuffle(rows)

    where = provenance(args.progress)
    if not where:
        print(f"WARNING: {args.progress} unreadable; no document column.")
    english = ({} if args.no_english
               else englished([(key_of(t), a, p, t) for a, p, t, _b, _x in rows],
                              args.translate_model, args.workers))

    wb = openpyxl.Workbook()
    ref = wb.active
    ref.title = "Answers"
    ref.append(["Answer", "What it means"])
    for code, text in CHOICES:
        ref.append([code, text])
    ref.column_dimensions["A"].width = 20
    ref.column_dimensions["B"].width = 96
    ref["A1"].font = Font(bold=True)
    ref["B1"].font = Font(bold=True)
    for row in ref.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")

    guide = wb.create_sheet("Question")
    for line in [
        f"The classifier labelled all {len(rows)} of these {args.target}.",
        "",
        "ONE QUESTION: does the SENTENCE state a reason that submitted",
        "evidence could not be used?",
        "",
        "Context_Before and Context_After are shown because the model saw",
        "them too. A cause in the context counts -- that is the point of",
        "the context.",
        "",
        "These are NOT reasons, and should stay NOT_A_REASON:",
        "  a heading or table caption",
        "  a statement of what the assessment covers",
        "  a description of what a drug is or does",
        "  a reported RESULT, whichever way it went",
        "  the company agreeing with the G-BA",
        "  a statutory statement that evidence need not be submitted (35a)",
        "  a criticism of Module 3 -- patient numbers or costs, not evidence",
        "",
        "If it DOES state a reason, answer STATES_A_REASON and write the",
        "reason in Notes in your own words. Do not pick a category code.",
        "",
        "Use UNSURE freely. It is excluded, not counted against the model.",
    ]:
        guide.append([line])
    guide.column_dimensions["A"].width = 78
    guide["A1"].font = Font(bold=True)

    sheet = wb.create_sheet("Coding", 0)
    sheet.append(["Item", "Ref", "Appraisal", "Document", "Page", "Link",
                  "Context_Before", "Sentence_English", "Sentence_German",
                  "Context_After", "Your_Answer", "Notes"])
    for i, (aid, page, text, before, after) in enumerate(rows, 1):
        doc, url = where.get(aid, ("unknown", ""))
        sheet.append([i, key_of(text), aid, doc, page, "",
                      before, english.get(text, ""), text, after, "", ""])
        cell = sheet.cell(row=i + 1, column=6)
        if url:
            cell.value = "open"
            cell.hyperlink = f"{url}#page={page}" if page else url
            cell.style = "Hyperlink"

    widths = {"A": 6, "B": 18, "C": 11, "D": 21, "E": 7, "F": 7, "G": 54,
              "H": 72, "I": 72, "J": 54, "K": 20, "L": 44}
    for col, width in widths.items():
        sheet.column_dimensions[col].width = width
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
    for row in sheet.iter_rows(min_row=2):
        for idx in (6, 7, 8, 9, 11):
            row[idx].alignment = Alignment(wrap_text=True, vertical="top")
    sheet.freeze_panes = "H2"

    dv = DataValidation(type="list",
                        formula1=f"=Answers!$A$2:$A${len(CHOICES) + 1}",
                        allow_blank=True, showDropDown=False)
    dv.error = "Pick NOT_A_REASON, STATES_A_REASON or UNSURE."
    dv.errorTitle = "Not an answer"
    sheet.add_data_validation(dv)
    dv.add(f"K2:K{len(rows) + 1}")

    wb.save(args.out)
    docs = Counter(where.get(a, ("unknown", ""))[0] for a, _p, _t, _b, _x in rows)
    ctx = sum(1 for _a, _p, _t, b, x in rows if b or x)
    print(f"\nsampled {len(rows)}, seed {args.seed}")
    print(f"  carrying context: {ctx} of {len(rows)}")
    for doc, n in docs.most_common():
        print(f"  {n:>3}   {doc}")
    print(f"\n-> {args.out}")
    print("\nColumn K is yours. One question per row, three answers. "
          "Read the\nQuestion sheet first -- it lists what does NOT count "
          "as a reason.")


def score(args):
    try:
        import openpyxl
    except ImportError:
        sys.exit("pip install openpyxl")
    if not os.path.exists(args.out):
        sys.exit(f"{args.out} not found. Build it first.")
    sheet = openpyxl.load_workbook(args.out, data_only=True)["Coding"]
    hdr = [str(c.value or "").strip() for c in sheet[1]]
    ci = {n: hdr.index(n) for n in hdr if n}

    right, wrong, unsure, blank = [], [], 0, 0
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row[ci["Ref"]]:
            continue
        ans = str(row[ci["Your_Answer"]] or "").strip().upper()
        item = (row[ci["Item"]], row[ci["Appraisal"]], row[ci["Page"]],
                str(row[ci["Sentence_English"]] or ""),
                str(row[ci["Notes"]] or ""))
        if not ans:
            blank += 1
        elif ans == "UNSURE":
            unsure += 1
        elif ans == "NOT_A_REASON":
            right.append(item)
        elif ans == "STATES_A_REASON":
            wrong.append(item)
        else:
            sys.exit(f"Item {item[0]}: '{ans}' is not one of the three "
                     f"answers.")

    n = len(right) + len(wrong)
    print(f"coded   : {n}")
    if blank:
        print(f"blank   : {blank}")
    if unsure:
        print(f"UNSURE  : {unsure}  (excluded)")
    if not n:
        sys.exit("\nNothing to score.")

    lo, hi = wilson(len(right), n)
    print(f"\n{args.target} PRECISION  {len(right)}/{n} = "
          f"{len(right) / n:.1%}   95% CI {lo:.1%} to {hi:.1%}")

    if wrong:
        print(f"\n{len(wrong)} sentence(s) you say DO state a reason:")
        for item, aid, page, text, note in wrong:
            print(f"\n  item {item}  {aid} p{page}")
            print(f"    {' '.join(text[:200].split())}")
            if note:
                print(f"    reason: {note[:150]}")

    print("\n" + "=" * 70)
    share = len(wrong) / n if n else 0
    print(f"If {share:.0%} of the {args.target} labels are wrong, that is "
          f"roughly\n{share * args.population:,.0f} of the "
          f"{args.population:,} in the full run.")
    if lo < 0.80:
        print(f"\nLower bound {lo:.0%}. Not safe to treat {args.target} as "
              f"a filter.")
    else:
        print(f"\nLower bound {lo:.0%}. Holds up as a filter.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--reasons", default=REASONS)
    ap.add_argument("--labels", default=LABELS)
    ap.add_argument("--progress", default=PROGRESS)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--target", default=TARGET)
    ap.add_argument("--n", type=int, default=40)
    ap.add_argument("--seed", type=int, default=20261114)
    ap.add_argument("--population", type=int, default=1003)
    ap.add_argument("--no-english", action="store_true")
    ap.add_argument("--translate-model", default="claude-haiku-4-5-20251001")
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--score", action="store_true")
    args = ap.parse_args()
    score(args) if args.score else build(args)


if __name__ == "__main__":
    main()
