import argparse
import json
import os
import random
import re
import sys
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE = "https://www.has-sante.fr"
HEADERS = {"User-Agent": "Mozilla/5.0 (link check)"}


def digits_of(aid):
    return len(re.sub(r"\D", "", str(aid)))


def url_of(doc):
    u = doc.get("url") or ""
    if not u:
        return ""
    if u.startswith("http"):
        return u
    return BASE + ("" if u.startswith("/") else "/") + u


def probe(session, aid, url, timeout):
    try:
        r = session.head(url, allow_redirects=True, timeout=timeout,
                         headers=HEADERS)
        if r.status_code in (403, 405, 501):
            r = session.get(url, stream=True, allow_redirects=True,
                            timeout=timeout, headers=HEADERS)
            r.close()
        return aid, r.status_code, url
    except Exception as exc:
        return aid, f"ERR {type(exc).__name__}", url


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sections", default="has_sections.json")
    ap.add_argument("--sample", type=int, default=600)
    ap.add_argument("--workers", type=int, default=12)
    ap.add_argument("--timeout", type=float, default=20.0)
    ap.add_argument("--seed", type=int, default=20261118)
    ap.add_argument("--out", default="has_link_check.json")
    ap.add_argument("--all", action="store_true")
    args = ap.parse_args()

    data = json.load(open(args.sections, encoding="utf-8"))
    targets = [(aid, url_of(doc)) for aid, doc in data.items() if url_of(doc)]
    print(f"{len(data):,} appraisals, {len(targets):,} with a document url")

    buckets = defaultdict(list)
    for aid, url in targets:
        buckets[digits_of(aid)].append((aid, url))
    print("by CT number length:")
    for d in sorted(buckets):
        print(f"  {d}-digit : {len(buckets[d]):>6,}")

    if args.all:
        pick = targets
    else:
        random.seed(args.seed)
        pick = []
        per = max(1, args.sample // max(1, len(buckets)))
        for d in sorted(buckets):
            pick += random.sample(buckets[d], min(per, len(buckets[d])))
    print(f"\nchecking {len(pick):,} links with {args.workers} workers\n")

    done = {}
    if os.path.exists(args.out):
        done = json.load(open(args.out, encoding="utf-8"))
        pick = [(a, u) for a, u in pick if a not in done]
        print(f"resuming - {len(done):,} already checked, {len(pick):,} to go")

    session = requests.Session()
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        jobs = {pool.submit(probe, session, a, u, args.timeout): a
                for a, u in pick}
        for i, fut in enumerate(as_completed(jobs), 1):
            aid, status, url = fut.result()
            done[aid] = {"status": status, "url": url}
            if i % 50 == 0 or i == len(jobs):
                print(f"  {i:,}/{len(jobs):,}")
                json.dump(done, open(args.out, "w", encoding="utf-8"))
    json.dump(done, open(args.out, "w", encoding="utf-8"))

    print(f"\n{len(done):,} links checked -> {args.out}\n")
    rows = defaultdict(lambda: Counter())
    for aid, rec in done.items():
        ok = rec["status"] == 200
        rows[digits_of(aid)]["ok" if ok else "dead"] += 1
    print(f"{'CT digits':<12}{'checked':>9}{'live':>8}{'dead':>7}{'dead %':>9}")
    for d in sorted(rows):
        c = rows[d]
        n = c["ok"] + c["dead"]
        print(f"{d:<12}{n:>9,}{c['ok']:>8,}{c['dead']:>7,}"
              f"{c['dead'] / n:>9.0%}")
    total = sum(c["ok"] + c["dead"] for c in rows.values())
    dead = sum(c["dead"] for c in rows.values())
    print(f"\noverall dead: {dead:,}/{total:,} = {dead / max(1, total):.0%}")

    codes = Counter(str(r["status"]) for r in done.values()
                    if r["status"] != 200)
    if codes:
        print("\nfailure codes:")
        for code, n in codes.most_common(8):
            print(f"  {n:>6,}  {code}")
        bad = next(a for a, r in done.items() if r["status"] != 200)
        print(f"\nexample dead link ({bad}):\n  {done[bad]['url']}")


if __name__ == "__main__":
    main()
