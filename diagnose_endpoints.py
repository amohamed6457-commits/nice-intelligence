import argparse
import os
import sys

import pandas as pd

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--endpoints", default="has_endpoints_full.json")
    ap.add_argument("--workbook",
                    default="NICE_HAS_GBA_updated_2026-11-07.xlsx")
    ap.add_argument("--sheet", default="HAS_Recent")
    args = ap.parse_args()

    print(f"working directory: {os.getcwd()}\n")

    print("1. module")
    try:
        import has_endpoints_panel as hep
    except Exception as exc:
        print(f"   FAIL cannot import has_endpoints_panel: {exc}")
        return
    print(f"   ok   {hep.__file__}")
    for name in ("ct_tokens", "ids_from_frame", "endpoint_table",
                 "render_endpoint_table"):
        print(f"   {'ok  ' if hasattr(hep, name) else 'MISS'} {name}")
    if not hasattr(hep, "ct_tokens"):
        print("\n   -> this is the OLD module. Overwrite it with the new one.")
        return

    print("\n2. endpoint file")
    if not os.path.exists(args.endpoints):
        print(f"   FAIL {args.endpoints} is not in this folder")
        near = [f for f in os.listdir(".") if f.startswith("has_endpoint")]
        print(f"   files here starting has_endpoint: {near or 'none'}")
        return
    size = os.path.getsize(args.endpoints) / 1e6
    df = hep.read_endpoints(args.endpoints)
    print(f"   ok   {args.endpoints}  {size:.1f} MB")
    if df.empty:
        print("   FAIL loaded but produced zero rows")
        return
    print(f"   {len(df):,} endpoints, {df['aid_key'].nunique():,} appraisals")
    print(f"   {int(df['has_result'].sum()):,} with a result")
    print(f"   key format: {list(df['appraisal_id'].head(3))}")

    print("\n3. workbook frame")
    if not os.path.exists(args.workbook):
        print(f"   FAIL {args.workbook} not found")
        return
    frame = pd.read_excel(args.workbook, args.sheet)
    print(f"   ok   {args.sheet}: {len(frame):,} rows")
    cols = hep.id_columns(frame)
    print(f"   id columns found: {cols or 'NONE'}")
    if not cols:
        print(f"   -> columns are: {list(frame.columns)[:12]}")
        return

    ids = hep.ids_from_frame(frame)
    print(f"   CT tokens extracted: {len(ids):,}")
    print(f"   sample: {ids[:4]}")

    print("\n4. the join")
    keys = set(df["aid_key"])
    hits = [i for i in ids if hep.norm_id(i) in keys]
    print(f"   endpoint file has {len(keys):,} appraisal keys")
    print(f"   frame offers      {len(ids):,} CT tokens")
    print(f"   OVERLAP           {len(hits):,}")
    if not hits:
        print("\n   FAIL nothing matches. Compare the two formats above.")
        print(f"   a frame token normalises to: "
              f"{hep.norm_id(ids[0]) if ids else '-'}")
        print(f"   an endpoint key looks like : {sorted(keys)[0]}")
        return

    print("\n5. the table the panel would render")
    links = {}
    try:
        docs = pd.read_excel(args.workbook, "Appraisal_Documents")
        links = hep.read_links(docs)
    except Exception as exc:
        print(f"   (no links: {exc})")
    table = hep.endpoint_table(df, links, ids, hep.labels_from_frame(frame))
    print(f"   {len(table):,} rows, "
          f"{table['Appraisal'].nunique() if len(table) else 0:,} appraisals")
    if table.empty:
        print("   FAIL table is empty despite an overlap - report this")
        return
    print()
    print(table.head(5)[["Appraisal", "Drug", "Endpoint", "Result",
                         "Page"]].to_string(index=False))
    print("\nThe data and the join are fine. If the app still shows nothing,")
    print("it is the Streamlit cache: press C in the app, or restart it.")


if __name__ == "__main__":
    main()
