import argparse
import json
import os
import shutil
import sys
import time

import pandas as pd

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DEFAULT_WORKBOOK = "NICE_HAS_GBA_updated_2026-11-07.xlsx"
DEFAULT_DIR = "data_parquet"


def safe_name(sheet):
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in str(sheet))


def writeable(frame):
    out = frame.copy()
    out.columns = [str(c) for c in out.columns]
    for col in out.columns:
        if out[col].dtype == object:
            mask = out[col].notna()
            out[col] = out[col].astype(str).where(mask, None)
    return out


def open_workbook(path):
    try:
        return pd.ExcelFile(path, engine="calamine")
    except (ImportError, ValueError):
        print("  (python-calamine not installed - using the slower engine)")
        return pd.ExcelFile(path)


def convert_workbook(path, folder):
    book = open_workbook(path)
    index = {}
    for sheet in book.sheet_names:
        frame = book.parse(sheet)
        target = os.path.join(folder, safe_name(sheet) + ".parquet")
        try:
            frame.to_parquet(target, compression="zstd", index=False)
        except Exception:
            writeable(frame).to_parquet(target, compression="zstd",
                                        index=False)
        index[sheet] = os.path.basename(target)
        print(f"  {sheet:<24}{len(frame):>8,} rows  "
              f"{os.path.getsize(target) / 1e6:6.1f} MB")
    return index


def convert_endpoints(src, folder):
    if not os.path.exists(src):
        print(f"  {src} not found - skipped")
        return None
    sys.path.insert(0, os.path.dirname(os.path.abspath(src)) or ".")
    try:
        import has_endpoints_panel as hep
    except ImportError:
        print("  has_endpoints_panel.py not importable - skipped")
        return None
    frame = hep.read_endpoints(src)
    if frame is None or frame.empty:
        print(f"  {src} produced no rows - skipped")
        return None
    target = os.path.join(folder, "has_endpoints.parquet")
    writeable(frame).to_parquet(target, compression="zstd", index=False)
    print(f"  endpoints{'':<16}{len(frame):>8,} rows  "
          f"{os.path.getsize(target) / 1e6:6.1f} MB")
    return os.path.basename(target)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", default=DEFAULT_WORKBOOK)
    ap.add_argument("--endpoints", default="has_endpoints_full.json")
    ap.add_argument("--out", default=DEFAULT_DIR)
    ap.add_argument("--keep", action="store_true",
                    help="do not clear the output folder first")
    args = ap.parse_args()

    if not os.path.exists(args.workbook):
        raise SystemExit(f"{args.workbook} not found")
    if not args.keep and os.path.isdir(args.out):
        shutil.rmtree(args.out)
    os.makedirs(args.out, exist_ok=True)

    start = time.time()
    print(f"workbook -> {args.out}/")
    index = convert_workbook(args.workbook, args.out)
    print("\nendpoints ->")
    endpoints = convert_endpoints(args.endpoints, args.out)

    manifest = {
        "built": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "workbook": os.path.basename(args.workbook),
        "workbook_mtime": os.path.getmtime(args.workbook),
        "sheets": index,
        "endpoints": endpoints,
    }
    with open(os.path.join(args.out, "manifest.json"), "w",
              encoding="utf-8") as fh:
        json.dump(manifest, fh, indent=1)

    total = sum(os.path.getsize(os.path.join(args.out, f))
                for f in os.listdir(args.out))
    print(f"\n{len(index)} sheets in {time.time() - start:.1f}s")
    print(f"{total / 1e6:.1f} MB in {args.out}/ "
          f"(workbook is {os.path.getsize(args.workbook) / 1e6:.1f} MB)")
    print("\nCommit the folder, and rebuild it whenever the workbook or the")
    print("endpoints change - the app falls back to the workbook if it is")
    print("missing or older, so a stale build cannot serve wrong data.")


if __name__ == "__main__":
    main()
