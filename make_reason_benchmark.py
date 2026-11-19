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
TRANSLATIONS = "translations.json"
PROGRESS = "fetch_iqwig_progress.json"
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


SHORT = [
    ("Addendum zur Nutzenbewertung", "IQWiG addendum"),
    ("Nutzenbewertung", "IQWiG assessment"),
    ("Zusammenfassende Dokumentation", "Zusammenfassende Dok."),
    ("Tragende", "Tragende Gruende"),
]


def provenance(path):
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as fh:
            raw = json.load(fh)
    except (ValueError, OSError):
        return {}
    out = {}
    for aid, info in raw.items():
        if not isinstance(info, dict) or info.get("error"):
            continue
        kind = str(info.get("type") or "")
        label = kind[:40]
        for needle, short in SHORT:
            if needle.lower() in kind.lower():
                label = short
                break
        out[aid] = (label, str(info.get("url") or ""))
    return out


def englished(rows, model, workers):
    import translate_passages as tp
    cache = {}
    if os.path.exists(TRANSLATIONS):
        try:
            with open(TRANSLATIONS, encoding="utf-8") as fh:
                cache = json.load(fh)
        except (ValueError, OSError):
            cache = {}
    need = [t for _k, _a, _p, t in rows if tp.key_of(t) not in cache]
    print(f"translating {len(need)} of {len(rows)} sentence(s) "
          f"({len(rows) - len(need)} already cached)")
    if need:
        try:
            import anthropic
        except ImportError:
            sys.exit("pip install anthropic")
        client = anthropic.Anthropic(api_key=tp.api_key())
        import concurrent.futures as futures
        with futures.ThreadPoolExecutor(max_workers=workers) as pool:
            jobs = {pool.submit(tp.translate_one, client, model, t, "de"): t
                    for t in need}
            bad = 0
            for done_n, fut in enumerate(futures.as_completed(jobs), 1):
                text = jobs[fut]
                try:
                    got = fut.result()
                except Exception as exc:
                    got = f"__ERROR__ {type(exc).__name__}: {exc}"
                if str(got).startswith("__ERROR__"):
                    bad += 1
                else:
                    cache[tp.key_of(text)] = got
                if done_n % 10 == 0 or done_n == len(need):
                    print(f"  {done_n}/{len(need)}, {bad} failed")
        tmp = TRANSLATIONS + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(cache, fh, ensure_ascii=False, indent=1)
        os.replace(tmp, TRANSLATIONS)
    out = {}
    for _k, _a, _p, text in rows:
        out[text] = cache.get(tp.key_of(text), "")
    missing = sum(1 for v in out.values() if not v)
    if missing:
        print(f"{missing} sentence(s) have no English. Code those from the "
              f"German column.")
    return out


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
    ap.add_argument("--progress", default=PROGRESS)
    ap.add_argument("--no-english", action="store_true")
    ap.add_argument("--translate-model", default="claude-haiku-4-5-20251001")
    ap.add_argument("--workers", type=int, default=6)
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

    where = provenance(args.progress)
    if not where:
        print(f"WARNING: {args.progress} not found or unreadable. The sheet "
              f"will not say\n         which document each page refers to.")

    english = {} if args.no_english else englished(
        rows, args.translate_model, args.workers)

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
    sheet.append(["Item", "Ref", "Appraisal", "Document", "Page", "Link",
                  "Sentence_English", "Sentence_German", "Your_Category",
                  "Notes"])
    for i, (key, aid, page, text) in enumerate(rows, 1):
        doc, url = where.get(aid, ("unknown", ""))
        sheet.append([i, key, aid, doc, page, "", english.get(text, ""),
                      text, "", ""])
        cell = sheet.cell(row=i + 1, column=6)
        if url:
            cell.value = "open"
            cell.hyperlink = f"{url}#page={page}" if page else url
            cell.style = "Hyperlink"

    widths = {"A": 6, "B": 18, "C": 11, "D": 23, "E": 7, "F": 7,
              "G": 82, "H": 82, "I": 34, "J": 40}
    for col, width in widths.items():
        sheet.column_dimensions[col].width = width
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
    for row in sheet.iter_rows(min_row=2):
        for idx in (6, 7, 9):
            row[idx].alignment = Alignment(wrap_text=True, vertical="top")
    sheet.freeze_panes = "G2"

    last = len(codes) + 1
    dv = DataValidation(type="list",
                        formula1=f"=Categories!$A$2:$A${last}",
                        allow_blank=True, showDropDown=False)
    dv.error = "Pick a code from the list, or UNSURE."
    dv.errorTitle = "Not a category"
    sheet.add_data_validation(dv)
    dv.add(f"I2:I{len(rows) + 1}")

    wb.save(args.out)

    spread = Counter(
        labels[key] for key, _a, _p, _t in rows)
    print(f"{len(pool):,} distinct classified sentences available")
    print(f"sampled {len(rows)} of them, seed {args.seed}, "
          f"floor {args.floor} per category\n")
    for label in sorted(spread, key=lambda l: -spread[l]):
        have = len(by_label[label])
        print(f"  {spread[label]:>3} of {have:>4}   {label}")
    docs = Counter(where.get(aid, ("unknown", ""))[0]
                   for _k, aid, _p, _t in rows)
    print("\nsampled from:")
    for doc, count in docs.most_common():
        print(f"  {count:>3}   {doc}")

    missing = sorted(set(CATEGORIES) - set(spread))
    if missing:
        print(f"\nNot represented (nothing classified as these): "
              f"{', '.join(missing)}")
    print(f"\n-> {args.out}")
    print("\nThe model's answers are NOT in this file. Column I is yours. "
          "Code all\nof them, then run score_reason_benchmark.py, which "
          "looks the predictions\nup by Ref at scoring time.")
    if not args.no_english:
        print("\nCode from the English column. The German is beside it "
              "because the model\nread the German, not the translation -- "
              "if a translation looks like it\nchanged the meaning, say so "
              "in Notes and that item comes out of the score.")
    print("\nRead the Rules sheet first. The table benchmark lost points to "
          "definitional\ndisagreement rather than error, and that is "
          "avoidable here.")


if __name__ == "__main__":
    main()
