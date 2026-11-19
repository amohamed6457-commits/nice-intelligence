"""
List the appraisals with no committee chapter, so they can be checked by hand.

why_no_chapters.py counts the failure reasons. This names the appraisals
behind one of them and writes a CSV with a clickable NICE URL, the drug and
the year, so the question "do these actually have a committee section?" can
be settled by opening a few rather than by argument.

The bucket and reason logic is imported from why_no_chapters rather than
copied, so the two scripts cannot drift into disagreeing about what counts
as a failure.

    python list_missing_chapters.py                     # the 224, to CSV
    python list_missing_chapters.py --sample 12         # a random dozen
    python list_missing_chapters.py --reason all        # every failure
    python list_missing_chapters.py --reason "done but" # the parse drops

Check a RANDOM sample, not the first ten. The index is in TA order, so the
first ten are consecutive appraisals from one period and will mislead you
about the rest.
"""

import argparse
import csv
import random
import re
import sys

import os

# Make the import work however this is invoked — python -I, a different
# working directory, or a launcher that does not put the script's folder on
# the path.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from why_no_chapters import _BUCKETS, buckets, describe, load, records
except ImportError as exc:
    sys.exit(f"why_no_chapters.py must sit beside this script ({exc}).")

INDEX = "chapter_index.json"
SECTIONS = "chapter_sections.json"
DEFAULT_REASON = "no chapter links"
OUT = "missing_chapters.csv"

WORKBOOKS = ["NICE_v14_updated_2026-09-29.xlsx", "NICE_v14_updated_2026-09-28.xlsx"]


def ta_number(key):
    m = re.search(r"(\d+)", str(key))
    return int(m.group(1)) if m else None


def lookup():
    """{ta number: (drug, indication, year)} from the workbook, if present."""
    import os
    path = next((w for w in WORKBOOKS if os.path.exists(w)), None)
    if not path:
        return {}, None
    try:
        import pandas as pd
        df = pd.read_excel(path, sheet_name="Sheet1")
    except Exception as exc:                                   # noqa: BLE001
        print(f"  (workbook not read: {exc})")
        return {}, None

    def col(*names):
        return next((n for n in names if n in df.columns), None)

    c_id = col("appraisal_id", "ta_id", "id")
    c_drug = col("generic_name", "drug_name", "technology", "inn")
    c_ind = col("indication", "appraised_indication")
    c_year = col("year_start", "year")
    out = {}
    for _, r in df.iterrows():
        n = ta_number(r.get(c_id)) if c_id else None
        if n is None:
            continue
        out[n] = (str(r.get(c_drug, "") or "")[:48],
                  str(r.get(c_ind, "") or "")[:90],
                  str(r.get(c_year, "") or ""))
    return out, path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--index", default=INDEX)
    ap.add_argument("--sections", default=SECTIONS)
    ap.add_argument("--reason", default=DEFAULT_REASON,
                    help="substring of the reason to list, or 'all'")
    ap.add_argument("--sample", type=int,
                    help="pick this many at random instead of listing all")
    ap.add_argument("--seed", type=int, default=7,
                    help="so the same sample comes back on a re-run")
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--show", type=int, default=15, help="rows to print")
    args = ap.parse_args()

    index = load(args.index)
    if index is None:
        sys.exit(f"{args.index} not found. Run this from the folder that holds it.")

    grouped = buckets(index)
    rows, origin = {}, {}
    if grouped:
        for name, entries in grouped:
            for key, value in entries.items():
                rows[key] = value
                origin[key] = str(name).lower()
    else:
        rows = dict(records(index))

    sections = load(args.sections)
    have = set()
    if sections is not None:
        for key, value in records(sections):
            if value:
                have.add(str(key).strip().lower())

    wanted = args.reason.strip().lower()
    hits = []
    for key, value in rows.items():
        norm = str(key).strip().lower()
        if have and norm in have:
            continue
        bucket = _BUCKETS.get(origin.get(key, ""), None)
        if bucket == "ok":
            why = "done but no sections — parse step dropped it"
        else:
            why = describe(value)
            if bucket == "failed" and why in {"no status recorded", "ok"}:
                why = "in failed bucket, no reason recorded"
        if why == "ok":
            continue
        if wanted != "all" and wanted not in why.lower():
            continue
        hits.append((ta_number(key), key, why))

    if not hits:
        sys.exit(f"Nothing matched reason {args.reason!r}. Run why_no_chapters.py "
                 "to see the reasons actually present.")

    hits.sort(key=lambda h: (h[0] is None, -(h[0] or 0)))   # newest TA first
    total = len(hits)
    if args.sample and args.sample < total:
        random.seed(args.seed)
        hits = sorted(random.sample(hits, args.sample),
                      key=lambda h: (h[0] is None, -(h[0] or 0)))
        print(f"{total} match {args.reason!r}; showing a random {len(hits)} "
              f"(seed {args.seed}, so re-running gives the same set)\n")
    else:
        print(f"{total} match {args.reason!r}\n")

    meta, wb = lookup()
    if wb:
        print(f"drug and year from {wb}\n")

    out_rows = []
    for n, key, why in hits:
        drug, indication, year = meta.get(n, ("", "", ""))
        url = f"https://www.nice.org.uk/guidance/ta{n}" if n else ""
        out_rows.append({"ta": f"TA{n}" if n else key, "url": url, "drug": drug,
                         "year": year, "indication": indication, "reason": why})

    for r in out_rows[:args.show]:
        print(f"  {r['ta']:<8} {r['year']:<6} {r['drug'][:30]:<32} {r['url']}")
    if len(out_rows) > args.show:
        print(f"  … {len(out_rows) - args.show} more in the CSV")

    with open(args.out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=["ta", "url", "drug", "year",
                                           "indication", "reason"])
        w.writeheader()
        w.writerows(out_rows)
    print(f"\n-> {args.out} ({len(out_rows)} rows)")

    years = [r["year"][:4] for r in out_rows if r["year"]]
    if years:
        from collections import Counter
        spread = ", ".join(f"{y}:{n}" for y, n in sorted(Counter(years).items()))
        print(f"\nby year: {spread}")
        print("If these cluster in the most recent years, they are very likely "
              "the\nnewer NICE format where the reasoning sits inside chapter 1 "
              "under\n'Why the committee made these recommendations' — recoverable, "
              "not absent.")
    print("\nOpen five or six of the URLs above. For each, note whether the page "
          "has\nany committee reasoning at all, and if so which chapter it is in.")


if __name__ == "__main__":
    main()
