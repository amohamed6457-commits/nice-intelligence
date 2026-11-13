"""
Download the IQWiG dossier assessments (and fallbacks) in bulk, resumably.

Why this is a separate script
-----------------------------
There are 412 IQWiG assessments behind the appraisals where no endpoint
category was assessable, plus ~50 that have only a Tragende Gründe. They are
large documents and the download takes hours, so it runs on its own and
everything else carries on while it does.

Resumable by design: a file already on disk is skipped, progress is written
after every document, and the whole thing can be stopped with Ctrl+C and
restarted without losing work or re-downloading anything.

Polite by design: one request at a time with a pause between them. These are
a public body's servers and there is no deadline that justifies hammering
them.

    python fetch_iqwig.py --dry-run      # what it would fetch, no requests
    python fetch_iqwig.py                # the 463 gap appraisals
    python fetch_iqwig.py --all          # every G-BA appraisal with one
    python fetch_iqwig.py --limit 50

Writes into iqwig_pdfs/ and records fetch_iqwig_progress.json.
"""

import argparse
import json
import os
import re
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from endpoint_summary import find_workbook
    from pilot_rejection_reasons import gap_appraisals
except ImportError as exc:
    sys.exit(f"endpoint_summary.py and pilot_rejection_reasons.py must sit "
             f"beside this script ({exc}).")

CACHE = "iqwig_pdfs"
PROGRESS = "fetch_iqwig_progress.json"

# In order of preference. IQWiG's assessment carries the methodological
# reasoning; the others are fallbacks for appraisals that have no IQWiG
# document, which is about one in nine.
PREFERRED = [
    "Nutzenbewertung IQWiG",
    "Addendum zur Nutzenbewertung des IQWiG",
    "Zusammenfassende Dokumentation",
    "Tragende Gründe",
]


def load(path):
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (ValueError, OSError):
        return {}


def save(path, payload):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def plan(workbook, everything):
    """appraisal -> (document type, url), best available document each."""
    import pandas as pd
    docs = pd.read_excel(workbook, sheet_name="Appraisal_Documents")
    docs["aid"] = docs["Appraisal_ID"].astype(str).str.strip()
    docs = docs[docs["Source_URL"].fillna("").astype(str).str.startswith("http")]
    wanted = None if everything else set(gap_appraisals(workbook))

    chosen = {}
    for row in docs.itertuples():
        if wanted is not None and row.aid not in wanted:
            continue
        kind = str(row.Document_Type or "")
        for rank, name in enumerate(PREFERRED):
            if name.lower() in kind.lower():
                best = chosen.get(row.aid)
                if best is None or rank < best[0]:
                    chosen[row.aid] = (rank, kind, str(row.Source_URL))
                break
    return {aid: (kind, url) for aid, (_rank, kind, url) in chosen.items()}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook")
    ap.add_argument("--all", action="store_true",
                    help="every G-BA appraisal, not just the 463 with no "
                         "assessable evidence")
    ap.add_argument("--limit", type=int)
    ap.add_argument("--pause", type=float, default=1.5,
                    help="seconds between requests (default 1.5)")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    wb = find_workbook(args.workbook)
    print(f"workbook: {wb}")
    targets = plan(wb, args.all)
    print(f"appraisals to cover: {len(targets):,}")

    from collections import Counter
    kinds = Counter(kind for kind, _url in targets.values())
    for kind, n in kinds.most_common():
        print(f"   {n:>5}  {kind[:52]}")

    os.makedirs(CACHE, exist_ok=True)
    progress = load(PROGRESS)
    todo = []
    for aid, (kind, url) in sorted(targets.items()):
        path = os.path.join(CACHE, re.sub(r"[^A-Za-z0-9_.-]", "_",
                                          f"{aid}.pdf"))
        if os.path.exists(path) and os.path.getsize(path) > 2000:
            continue
        todo.append((aid, kind, url, path))
    print(f"\nalready on disk : {len(targets) - len(todo):,}")
    print(f"to download now : {len(todo):,}")
    if args.limit:
        todo = todo[:args.limit]
        print(f"limited to      : {len(todo):,}")

    if args.dry_run:
        for aid, kind, url, _path in todo[:12]:
            print(f"   {aid:<9} {kind[:34]:<34} {url[:66]}")
        print("\n(dry run — nothing downloaded)")
        return
    if not todo:
        print("\nNothing to do.")
        return

    try:
        import requests
    except ImportError:
        sys.exit("requests is not installed: pip install requests")

    started, ok, failed, bytes_got = time.time(), 0, 0, 0
    print()
    for i, (aid, kind, url, path) in enumerate(todo, 1):
        try:
            resp = requests.get(
                url, timeout=180,
                headers={"User-Agent": "hta-intelligence-research"})
            resp.raise_for_status()
            if not resp.content.startswith(b"%PDF"):
                raise ValueError(f"not a PDF ({len(resp.content):,} bytes)")
            tmp = path + ".part"
            with open(tmp, "wb") as fh:
                fh.write(resp.content)
            os.replace(tmp, path)
            ok += 1
            bytes_got += len(resp.content)
            progress[aid] = {"type": kind, "url": url,
                             "bytes": len(resp.content)}
        except Exception as exc:                           # noqa: BLE001
            failed += 1
            progress[aid] = {"type": kind, "url": url,
                             "error": f"{type(exc).__name__}: {exc}"[:200]}
        if i % 10 == 0 or i == len(todo):
            save(PROGRESS, progress)
            rate = (time.time() - started) / i
            left = (len(todo) - i) * rate / 60
            print(f"  [{i:>4}/{len(todo)}] {ok} ok, {failed} failed, "
                  f"{bytes_got / 1e6:.0f} MB, ~{left:.0f} min left")
        time.sleep(args.pause)

    save(PROGRESS, progress)
    print(f"\n{ok:,} downloaded ({bytes_got / 1e6:.0f} MB), {failed} failed")
    print(f"-> {CACHE}/   progress in {PROGRESS}")
    if failed:
        print("Re-run to retry the failures; successes are skipped.")
    print("\nAdd iqwig_pdfs/ to .gitignore — this must not be committed.")


if __name__ == "__main__":
    main()
