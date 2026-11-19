import argparse
import json
import os
import re
import sys
import time
from collections import Counter, defaultdict

CACHE = "has_ct_pdfs"
OUT = "has_sections.json"
WORKBOOK = "NICE_HAS_GBA_updated_2026-11-07.xlsx"

WANTED = re.compile(r"Committee\s+opinion", re.IGNORECASE)

HEADING = re.compile(
    r"^\s*(\d{1,2}(?:\.\d{1,2}){0,2})\.?\s+"
    r"([A-Za-z\u00c0-\u00ff][^\n]{2,70}?)\s*$")

SYNTHESE = re.compile(r"SYNTH[\u00c8E]SE|\u00c9VALUER\s+LES\s+TECHNOLOGIES",
                      re.IGNORECASE)

SUBHEADING = re.compile(
    r"^\s*(0?\d{1,2}\.\d{1,2})\s+"
    r"([A-ZÀÂÄÇÉÈÊËÎÏÔÖÙÛÜŸ][^\n]{2,55})\s*$")

CAPS_ONLY = re.compile(
    r"^\s*([A-ZÀÂÄÇÉÈÊËÎÏÔÖÙÛÜŸ][A-ZÀÂÄÇÉÈÊËÎÏÔÖÙÛÜŸ\s'’,()\-/]{8,})\s*$")

MARKERS = {
    "SMR verdict": r"service\s+m[ée]dical\s+rendu",
    "ASMR verdict": r"am[ée]lioration\s+du\s+service\s+m[ée]dical",
    "efficacy section": r"\bEFFICACIT[ÉE]\b",
    "safety section": r"TOL[ÉE]RANCE|S[ÉE]CURIT[ÉE]",
    "available data": r"ANALYSE\s+DES\s+DONN[ÉE]ES",
    "comparators": r"COMPARATEURS?\s+CLINIQUEMENT",
    "therapeutic place": r"PLACE\s+DANS\s+LA\s+STRAT[ÉE]GIE",
    "commission conclusions": r"CONCLUSIONS?\s+DE\s+LA\s+COMMISSION",
    "medical need": r"BESOIN\s+M[ÉE]DICAL",
}


def undouble(line):
    stripped = line.strip()
    if len(stripped) < 8:
        return line
    pairs = sum(1 for i in range(0, len(stripped) - 1, 2)
                if stripped[i] == stripped[i + 1])
    if pairs / max(1, len(stripped) // 2) >= 0.7:
        return stripped[::2]
    return line


def normalise(text):
    text = re.sub(r"\s+", " ", text).strip().upper()
    text = re.sub(r"[’']", "'", text)
    return text[:60]


def plan(workbook, per_bucket):
    import pandas as pd
    docs = pd.read_excel(workbook, sheet_name="Appraisal_Documents")
    docs["aid"] = docs["Appraisal_ID"].astype(str)
    docs = docs[docs["aid"].str.startswith("CT")]
    docs = docs[docs["Document_Type"].fillna("").astype(str).str.contains(
        WANTED, regex=True)]
    url = docs["Source_URL"].fillna("").astype(str)
    docs = docs[url.str.contains(r"ct_?\d{4,}\.pdf", case=False, regex=True)]
    year = docs["appraisal_date"].astype(str).str.extract(r"(\d{4})")[0]
    docs = docs.assign(year=year.fillna("unknown"))

    buckets = defaultdict(list)
    for row in docs.itertuples():
        try:
            y = int(row.year)
        except (ValueError, TypeError):
            buckets["unknown"].append(row)
            continue
        buckets[f"{y // 5 * 5}s"].append(row)

    chosen = []
    for bucket in sorted(buckets):
        rows = buckets[bucket]
        step = max(1, len(rows) // per_bucket)
        for row in rows[::step][:per_bucket]:
            chosen.append((bucket, row.aid, str(row.Source_URL)))
    return chosen, {b: len(v) for b, v in buckets.items()}


def headings_in(path, max_pages):
    import pdfplumber
    found, blank, pages, synthese = [], 0, 0, False
    with pdfplumber.open(path) as pdf:
        for number, page in enumerate(pdf.pages[:max_pages], 1):
            pages += 1
            try:
                text = page.extract_text() or ""
            except Exception:
                blank += 1
                continue
            if not text.strip():
                blank += 1
                continue
            if number == 1 and SYNTHESE.search(text[:400]):
                synthese = True
            for raw in text.splitlines():
                line = undouble(raw)
                m = HEADING.match(line) or SUBHEADING.match(line)
                if m and len(m.group(2).split()) <= 10:
                    found.append((number, m.group(1).strip(),
                                  normalise(m.group(2))))
                    continue
                c = CAPS_ONLY.match(line)
                if c and len(c.group(1).split()) <= 8:
                    found.append((number, "", normalise(c.group(1))))
    return found, pages, blank, synthese


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", default=WORKBOOK)
    ap.add_argument("--per-bucket", type=int, default=3)
    ap.add_argument("--max-pages", type=int, default=40)
    ap.add_argument("--pause", type=float, default=1.5)
    ap.add_argument("--cache", default=CACHE)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--check-links", action="store_true")
    ap.add_argument("--dump")
    ap.add_argument("--dump-pages", type=int, default=3)
    ap.add_argument("--show", type=int, default=2)
    args = ap.parse_args()

    if not os.path.exists(args.workbook):
        sys.exit(f"{args.workbook} not found.")
    try:
        import pdfplumber
    except ImportError:
        sys.exit("pip install pdfplumber")
    try:
        import requests
    except ImportError:
        sys.exit("pip install requests")

    if args.dump:
        want = args.dump.strip().upper().replace("_", "-")
        import pandas as pd
        docs = pd.read_excel(args.workbook, sheet_name="Appraisal_Documents")
        docs["aid"] = docs["Appraisal_ID"].astype(str).str.upper()
        row = docs[docs["aid"] == want]
        if row.empty:
            sys.exit(f"{want} not in Appraisal_Documents.")
        path = os.path.join(args.cache,
                            re.sub(r"[^A-Za-z0-9_.-]", "_", f"{want}.pdf"))
        if not os.path.exists(path):
            url = str(row.iloc[0]["Source_URL"])
            print(f"fetching {url[:90]}")
            resp = requests.get(url, timeout=120,
                                headers={"User-Agent":
                                         "hta-intelligence-research"})
            resp.raise_for_status()
            os.makedirs(args.cache, exist_ok=True)
            with open(path, "wb") as fh:
                fh.write(resp.content)
        with pdfplumber.open(path) as pdf:
            print(f"{len(pdf.pages)} pages\n")
            for number, page in enumerate(pdf.pages[:args.dump_pages], 1):
                print(f"{'=' * 66}\nPAGE {number}  "
                      f"(extract_text, raw lines)\n{'=' * 66}")
                text = page.extract_text() or ""
                for i, line in enumerate(text.splitlines()[:45]):
                    print(f"  {i:>3} | {line[:118]}")
                words = page.extract_words() or []
                big = sorted({round(w.get("size", 0), 1) for w in words})
                print(f"\n  font sizes on this page: {big[:12]}")
                if big:
                    top = big[-1]
                    tall = [w["text"] for w in words
                            if round(w.get("size", 0), 1) >= top - 0.6][:20]
                    print(f"  words at the largest size ({top}): "
                          f"{' '.join(tall)[:110]}")
        return

    chosen, sizes = plan(args.workbook, args.per_bucket)
    if not chosen:
        sys.exit("No committee opinions matched. Loosen --per-bucket or "
                 "check Appraisal_Documents.")
    print("committee opinions available, by five-year bucket:")
    for bucket in sorted(sizes):
        print(f"  {bucket:<10} {sizes[bucket]:>6,}")
    print(f"\nsampling {len(chosen)} documents "
          f"({args.per_bucket} per bucket)")
    if args.dry_run:
        for bucket, aid, url in chosen:
            print(f"  {bucket:<10} {aid:<10} {url}")
        return

    if args.check_links:
        print()
        codes = Counter()
        for bucket, aid, url in chosen:
            try:
                r = requests.head(url, timeout=30, allow_redirects=True,
                                  headers={"User-Agent":
                                           "hta-intelligence-research"})
                code = r.status_code
                kind = r.headers.get("content-type", "?").split(";")[0]
            except Exception as exc:
                code, kind = type(exc).__name__, "-"
            codes[code] += 1
            flag = "ok " if code == 200 else "ERR"
            print(f"  {flag} {code}  {kind:<16} {aid:<10} {url[-58:]}")
            time.sleep(0.4)
        print("\nstatus codes:")
        for code, n in codes.most_common():
            print(f"  {n:>4}  {code}")
        return

    os.makedirs(args.cache, exist_ok=True)
    per_doc, failed = {}, []
    counts, by_bucket = Counter(), defaultdict(Counter)
    started = time.time()
    for i, (bucket, aid, url) in enumerate(chosen, 1):
        path = os.path.join(args.cache,
                            re.sub(r"[^A-Za-z0-9_.-]", "_", f"{aid}.pdf"))
        if not (os.path.exists(path) and os.path.getsize(path) > 2000):
            try:
                resp = requests.get(
                    url, timeout=120,
                    headers={"User-Agent": "hta-intelligence-research"})
                resp.raise_for_status()
                if not resp.content.startswith(b"%PDF"):
                    raise ValueError("not a PDF")
                with open(path + ".part", "wb") as fh:
                    fh.write(resp.content)
                os.replace(path + ".part", path)
            except Exception as exc:
                failed.append((aid, f"{type(exc).__name__}: {exc}"[:90]))
                time.sleep(args.pause)
                continue
            time.sleep(args.pause)
        try:
            found, pages, blank, synth = headings_in(path, args.max_pages)
        except Exception as exc:
            failed.append((aid, f"parse {type(exc).__name__}: {exc}"[:90]))
            continue
        per_doc[aid] = {"bucket": bucket, "url": url, "pages": pages,
                        "blank_pages": blank, "synthese": synth,
                        "headings": [{"page": p, "number": n, "text": t}
                                     for p, n, t in found]}
        for _p, _n, text in found:
            counts[text] += 1
            by_bucket[bucket][text] += 1
        if i % 5 == 0 or i == len(chosen):
            rate = (time.time() - started) / i
            print(f"  {i}/{len(chosen)}  ({rate:.1f}s each)")

    with open(args.out, "w", encoding="utf-8") as fh:
        json.dump(per_doc, fh, ensure_ascii=False, indent=1)

    print(f"\nread {len(per_doc)} document(s), {len(failed)} failed")
    for aid, why in failed[:8]:
        print(f"    {aid}: {why}")
    if not per_doc:
        sys.exit("Nothing parsed.")

    lengths = sorted(d["pages"] for d in per_doc.values())
    synth = [a for a, d in per_doc.items() if d.get("synthese")]
    if synth:
        print(f"{len(synth)} document(s) are a SYNTHESE, not a full avis: "
              f"{', '.join(synth[:10])}")
    short = [a for a, d in per_doc.items()
             if d["pages"] <= 3 and not d.get("synthese")]
    print(f"page count: min {lengths[0]}, median "
          f"{lengths[len(lengths) // 2]}, max {lengths[-1]}")
    if short:
        print(f"{len(short)} document(s) are 3 pages or fewer -- probably "
              f"synopses, not full opinions:\n    {', '.join(short[:10])}")
    numbered = sum(1 for d in per_doc.values()
                   if any(h["number"] for h in d["headings"]))
    print(f"documents with NUMBERED headings: {numbered} of {len(per_doc)}")
    blanks = sum(d["blank_pages"] for d in per_doc.values())
    print(f"pages that produced no text: {blanks} "
          f"(scanned documents would show here)")

    print(f"\nheadings seen in at least 3 documents:")
    for text, n in counts.most_common():
        docs = sum(1 for d in per_doc.values()
                   if any(h["text"] == text for h in d["headings"]))
        if docs >= 3:
            print(f"  {docs:>3} docs  {text}")

    print("\nwhich landmarks each document has:")
    print(f"  {'marker':<26} {'docs':>5}   {'share':>6}")
    for label, pat in MARKERS.items():
        hit = sum(1 for d in per_doc.values()
                  if any(re.search(pat, h["text"], re.IGNORECASE)
                         for h in d["headings"]))
        print(f"  {label:<26} {hit:>5}   {hit / len(per_doc):>6.0%}")

    print("\nby era, do the numbered headings hold:")
    for bucket in sorted(by_bucket):
        docs = [d for d in per_doc.values() if d["bucket"] == bucket]
        num = sum(1 for d in docs if any(h["number"] for h in d["headings"]))
        print(f"  {bucket:<10} {num}/{len(docs)} numbered, "
              f"median {sorted(len(d['headings']) for d in docs)[len(docs) // 2]} "
              f"headings, median {sorted(d['pages'] for d in docs)[len(docs) // 2]} pages")

    shown = 0
    for aid, d in per_doc.items():
        if shown >= args.show:
            break
        shown += 1
        print(f"\n{'=' * 66}\n{aid}  ({d['bucket']}, {d['pages']} pages)")
        for h in d["headings"][:26]:
            print(f"  p{h['page']:<3} {h['number']:<6} {h['text']}")

    print(f"\n-> {args.out}")
    print("\nIf the numbered headings hold across eras, the clinical data can")
    print("be read from named sections with a page number for each, which is")
    print("what HAS is missing. If they do not, say so before anything is")
    print("built on top of them.")
    print(f"\nAdd {args.cache}/ to .gitignore.")


if __name__ == "__main__":
    main()
