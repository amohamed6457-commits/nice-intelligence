"""
Pilot: do HAS committee opinions yield tables the way G-BA resolutions do?

G-BA table recovery is now measured at 85.8% of endpoints intact. The
question this answers is whether the same machinery works on HAS, where the
situation is different in three ways:

  * the documents are the avis de la Commission de la Transparence, linked
    from Appraisal_Documents (21,152 PDFs across 12,343 appraisals), not the
    economic opinions already text-indexed in HAS_Economic_Pages (229);
  * they are in French, with French number formatting;
  * nothing about their layout has been checked, where G-BA's was.

Every gate and repair is imported from pilot_gba_tables, so this measures
HAS documents against the SAME rules that produced the G-BA figure. If the
rules need changing for French layouts, that is itself the finding.

Run it on three or four before anyone considers 21,000.

    python pilot_has_tables.py
    python pilot_has_tables.py --appraisal CT-15634 --keep-all
    python pilot_has_tables.py --type "CEESP economic opinion" --n 3
"""

import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from pilot_gba_tables import (SETTINGS, alignment_ok, clean, find_workbook,
                                  looks_like_data, merge_split_columns, tidy)
except ImportError as exc:
    sys.exit(f"pilot_gba_tables.py must sit beside this script ({exc}).")

CACHE = "has_pdfs"
DEFAULT_TYPE = "Committee opinion / associated paper"


def fetch(url, name):
    os.makedirs(CACHE, exist_ok=True)
    path = os.path.join(CACHE, re.sub(r"[^A-Za-z0-9_.-]", "_", f"{name}.pdf"))
    if os.path.exists(path) and os.path.getsize(path) > 1000:
        print(f"  cached: {path}")
        return path
    try:
        import requests
    except ImportError:
        sys.exit("requests is not installed: pip install requests")
    print(f"  downloading {url[:110]}")
    resp = requests.get(url, timeout=120,
                        headers={"User-Agent": "hta-intelligence-pilot"})
    resp.raise_for_status()
    if not resp.content.startswith(b"%PDF"):
        print(f"  NOT A PDF — {len(resp.content):,} bytes, starts "
              f"{resp.content[:24]!r}")
        return None
    with open(path, "wb") as fh:
        fh.write(resp.content)
    print(f"  saved {len(resp.content):,} bytes")
    return path


def show(rows, page, index, note, limit=10):
    width = max(len(r) for r in rows)
    print(f"\n  ── page {page}, table {index} ── {note}")
    for r in rows[:limit]:
        cells = [(c[:24] + "…") if len(c) > 25 else c for c in r]
        cells += [""] * (width - len(cells))
        print("    | " + " | ".join(f"{c:<25}" for c in cells))
    if len(rows) > limit:
        print(f"    … {len(rows) - limit} more rows")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook")
    ap.add_argument("--appraisal", action="append")
    ap.add_argument("--type", default=DEFAULT_TYPE,
                    help=f"Document_Type to pull (default: {DEFAULT_TYPE!r})")
    ap.add_argument("--n", type=int, default=3, help="how many to try")
    ap.add_argument("--keep-all", action="store_true")
    ap.add_argument("--seed", type=int, default=20261110)
    args = ap.parse_args()

    try:
        import pdfplumber
    except ImportError:
        sys.exit("pdfplumber is not installed: pip install pdfplumber")
    import pandas as pd
    import random

    wb = find_workbook(args.workbook)
    print(f"workbook: {wb}")
    docs = pd.read_excel(wb, sheet_name="Appraisal_Documents")
    has = docs[docs["Institution"].astype(str).str.contains("HAS", case=False,
                                                            na=False)]
    has = has[has["Source_URL"].fillna("").astype(str).str.endswith(".pdf")]
    if args.type and args.type.lower() != "any":
        picked = has[has["Document_Type"].astype(str) == args.type]
        if not len(picked):
            print(f"\nNo documents of type {args.type!r}. Types available:")
            print(has["Document_Type"].value_counts().head(10).to_string())
            sys.exit(1)
        has = picked

    print(f"HAS PDFs of this type: {len(has):,} "
          f"across {has['Appraisal_ID'].nunique():,} appraisals\n")

    if args.appraisal:
        wanted = {a.strip() for a in args.appraisal}
        rows = has[has["Appraisal_ID"].astype(str).str.strip().isin(wanted)]
        if not len(rows):
            sys.exit(f"None of {', '.join(wanted)} have a PDF of this type.")
    else:
        ids = sorted(set(has["Appraisal_ID"].astype(str)))
        random.seed(args.seed)
        chosen = random.sample(ids, min(args.n, len(ids)))
        rows = has[has["Appraisal_ID"].astype(str).isin(chosen)]
        print(f"random sample (seed {args.seed}): {', '.join(chosen)}\n")

    seen, kept_total, tried = set(), 0, 0
    for _, row in rows.iterrows():
        aid = str(row["Appraisal_ID"]).strip()
        if aid in seen:
            continue
        seen.add(aid)
        tried += 1
        print(f"=== {aid} — {str(row['Document_Type'])[:46]} ===")
        try:
            path = fetch(str(row["Source_URL"]), aid)
        except Exception as exc:                       # noqa: BLE001
            print(f"  FAILED {type(exc).__name__}: {exc}\n")
            continue
        if not path:
            print()
            continue
        kept = rejected = 0
        with pdfplumber.open(path) as pdf:
            print(f"  {len(pdf.pages)} pages")
            for page_no, page in enumerate(pdf.pages, 1):
                try:
                    cands = page.extract_tables(SETTINGS)
                except Exception:                      # noqa: BLE001
                    continue
                for i, raw in enumerate(cands, 1):
                    body = [[clean(c) for c in r] for r in raw]
                    body = [r for r in body if any(r)]
                    if not body:
                        continue
                    body, _notes = tidy(body)
                    if not body:
                        continue
                    body = merge_split_columns(body)
                    ok, why = looks_like_data(body)
                    if ok:
                        aligned, aw = alignment_ok(body)
                        if not aligned:
                            ok, why = False, aw
                    if ok:
                        kept += 1
                        show(body, page_no, i, why)
                    else:
                        rejected += 1
                        if args.keep_all:
                            show(body, page_no, i, f"REJECTED: {why}", limit=4)
        kept_total += kept
        print(f"\n  -> {kept} kept, {rejected} rejected\n")

    print("=" * 62)
    print(f"{tried} HAS document(s) tried, {kept_total} data table(s) kept.")
    if not kept_total:
        print("\nNothing survived. Re-run with --keep-all to see what was found\n"
              "and why it was rejected. Two likely causes: the opinions are\n"
              "prose rather than tabulated, or their tables are drawn without\n"
              "ruling lines.")
    else:
        print("\nCheck whether the numbers landed under the right headings, and\n"
              "whether French decimal commas came through intact.")


if __name__ == "__main__":
    main()
