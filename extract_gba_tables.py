"""
Batch-extract the tables from G-BA resolutions into gba_tables.json.

The app shows each resolution's passage text. Where that passage is a table
in the original, the text is close to unreadable — the line order survives
but the columns do not, so an endpoints table arrives as a run of numbers.
This rebuilds those tables from the PDFs, where the column geometry still
exists, and stores them for the app to render beside the passage.

All of the table repair and the safety gates live in pilot_gba_tables.py and
are imported rather than copied, so there is one definition of what counts as
a trustworthy table. A table that fails any gate is simply not stored, and
the app falls back to the passage and the page link — which are never wrong.

Resumable. Interrupt it and run it again; appraisals already in the output
are skipped unless --refresh is passed. PDFs are cached in gba_pdfs/ and
reused, so a second pass costs no downloads.

    python extract_gba_tables.py --limit 5        # try a handful first
    python extract_gba_tables.py                  # everything, resumable
    python extract_gba_tables.py --appraisal D-1058 --refresh

Expect roughly 1 GB of PDFs and a few hours for the full set. Run it
overnight; it writes after every appraisal, so stopping it loses nothing.
"""

import argparse
import json
import os
import sys
import time
import traceback

OUT = "gba_tables.json"
ERRORS = "gba_tables_errors.json"

try:
    from pilot_gba_tables import (CACHE, SETTINGS, alignment_ok, clean, fetch,
                                  find_workbook, looks_like_data,
                                  merge_split_columns, tidy)
except ImportError as exc:
    sys.exit(f"pilot_gba_tables.py must sit beside this script ({exc}).")


def load(path):
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (ValueError, OSError) as exc:
        sys.exit(f"{path} exists but could not be read ({exc}). Move it aside "
                 "and re-run to start fresh.")


def save(path, payload):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, path)          # atomic: a crash mid-write cannot corrupt


def tables_in(pdf_path):
    """Every trustworthy table in one PDF, as {page, header, rows, notes}."""
    import pdfplumber

    found = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, 1):
            try:
                candidates = page.extract_tables(SETTINGS)
            except Exception:                            # noqa: BLE001
                continue
            for raw in candidates:
                body = [[clean(c) for c in row] for row in raw]
                body = [r for r in body if any(r)]
                if not body:
                    continue
                body, notes = tidy(body)
                if not body:
                    continue
                body = merge_split_columns(body)
                ok, _ = looks_like_data(body)
                if not ok:
                    continue
                aligned, _ = alignment_ok(body)
                if not aligned:
                    continue
                # A header row is one that is not itself data.
                from pilot_gba_tables import _numeric_count
                has_header = (_numeric_count(body[0]) < 2
                              and not (body[0][0] and not any(body[0][1:])))
                found.append({
                    "page": page_no,
                    "header": body[0] if has_header else [],
                    "rows": body[1:] if has_header else body,
                    "notes": notes,
                })
    return found


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook")
    ap.add_argument("--limit", type=int, help="process at most this many")
    ap.add_argument("--appraisal", action="append",
                    help="only these ids; repeatable")
    ap.add_argument("--refresh", action="store_true",
                    help="redo appraisals already in the output")
    ap.add_argument("--pause", type=float, default=0.5,
                    help="seconds between downloads (default 0.5)")
    args = ap.parse_args()

    try:
        import pdfplumber                                # noqa: F401
    except ImportError:
        sys.exit("pdfplumber is not installed: pip install pdfplumber")
    import pandas as pd

    workbook = find_workbook(args.workbook)
    passages = pd.read_excel(workbook, sheet_name="GBA_Decision_Passages")
    passages["appraisal_id"] = passages["appraisal_id"].astype(str).str.strip()

    # One URL per appraisal; the resolution PDF is the same for every passage.
    pairs = (passages.dropna(subset=["source_url"])
             .groupby("appraisal_id")["source_url"].first().to_dict())
    if args.appraisal:
        wanted = {a.strip() for a in args.appraisal}
        pairs = {k: v for k, v in pairs.items() if k in wanted}
        missing = wanted - set(pairs)
        if missing:
            print(f"not in the sheet: {', '.join(sorted(missing))}")

    out = load(OUT)
    errors = load(ERRORS)
    todo = [k for k in pairs if args.refresh or k not in out]
    if args.limit:
        todo = todo[:args.limit]

    print(f"workbook : {workbook}")
    print(f"resolutions with a PDF : {len(pairs):,}")
    print(f"already done           : {len(out):,}")
    print(f"to process now         : {len(todo):,}\n")
    if not todo:
        print("Nothing to do. Pass --refresh to redo.")
        return

    os.makedirs(CACHE, exist_ok=True)
    started = time.time()
    kept_total = 0
    for i, appraisal in enumerate(todo, 1):
        url = str(pairs[appraisal])
        try:
            path = fetch(url, appraisal)
            tables = tables_in(path)
            out[appraisal] = tables
            errors.pop(appraisal, None)
            kept_total += len(tables)
            status = f"{len(tables)} table(s)"
        except KeyboardInterrupt:
            print("\ninterrupted — progress is saved, re-run to continue.")
            break
        except Exception as exc:                          # noqa: BLE001
            errors[appraisal] = {"url": url, "error": f"{type(exc).__name__}: {exc}",
                                 "trace": traceback.format_exc()[-800:]}
            status = f"FAILED {type(exc).__name__}"
        save(OUT, out)
        save(ERRORS, errors)
        rate = (time.time() - started) / i
        left = (len(todo) - i) * rate
        print(f"[{i:>4}/{len(todo)}] {appraisal:<9} {status:<22} "
              f"~{left/60:.0f} min left")
        if args.pause:
            time.sleep(args.pause)

    print(f"\n{len(out):,} resolution(s) in {OUT}; {kept_total:,} table(s) this run.")
    if errors:
        print(f"{len(errors)} failure(s) recorded in {ERRORS} — re-run to retry them.")
    with_tables = sum(1 for v in out.values() if v)
    print(f"{with_tables:,} of {len(out):,} resolutions have at least one table.")


if __name__ == "__main__":
    main()
