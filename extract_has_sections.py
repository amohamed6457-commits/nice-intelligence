import argparse
import json
import os
import re
import sys
import time
from collections import Counter, defaultdict

CACHE = "has_ct_pdfs"
OUT = "has_clinical_sections.json"
WORKBOOK = "NICE_HAS_GBA_updated_2026-11-07.xlsx"

PREFERENCE = [
    (r"avisdef[_.]", 0),
    (r"avis\d[_.]", 1),
    (r"avis[_.]|_avis", 2),
    (r"^ct_?\d{4,}\.pdf$", 3),
    (r"\.pdf$", 4),
]
NOT_AVIS = re.compile(r"synthese|summary|transcription", re.IGNORECASE)
RENEWAL = re.compile(r"_RI_|_ri_|renouvellement", re.IGNORECASE)
PROJET = re.compile(r"CT[_-]?AP[_-]", re.IGNORECASE)

HEADING = re.compile(r"^\s*(\d{1,2}(?:\.\d{1,2}){0,2})\.?\s+"
                     r"([A-Za-zÀ-ÿ][^\n]{2,70}?)\s*$")

NUMBER_ONLY = re.compile(r"^\s*(\d{1,2}(?:\.\d{1,2}){0,2})\.?\s*$")

TITLE_LINE = re.compile(r"^\s*([A-Za-z\u00c0-\u00ff][^\n]{2,70}?)\s*$")

SYNOPSIS = re.compile(
    r"^\s*(SMR|ASMR|ISP)\b\s*(?:[:.\u2013-]\s*)?(.{0,60})$")

VERDICT = re.compile(
    r"\b(important|mod[ée]r[ée]|faible|insuffisant)\b|"
    r"ASMR\s+(I{1,3}V?|V|IV)\b|"
    r"n'apporte\s+pas\s+d'am[ée]lioration", re.IGNORECASE)

SYNTHESE = re.compile(r"^\s*SYNTH[ÈE]SE\b",
                      re.IGNORECASE | re.MULTILINE)

ROLES = [
    ("asmr", r"AM[EÉ]LIORATION\s+DU\s+SERVICE\s+M[EÉ]DICAL"),
    ("smr", r"SERVICE\s+M[EÉ]DICAL\s+RENDU"),
    ("clinical_evidence",
     r"ANALYSE\s+DES\s+(?:NOUVELLES\s+)?DONN[EÉ]ES|DONN[EÉ]ES\s+CLINIQUES|"
     r"DONN[EÉ]ES\s+ISSUES\s+DES\s+[EÉ]TUDES|"
     r"PR[EÉ]SENTATION\s+DES\s+PRINCIPALES\s+[EÉ]TUDES|"
     r"ESSAIS\s+PR[EÉ]SENT|EFFICACIT|R[EÉ]ACTUALISATION\s+DES\s+DONN[EÉ]ES|"
     r"QUALIT[EÉ]\s+DE\s+VIE|R[EÉ]SULTATS"),
    ("discussion", r"R[EÉ]SUM[EÉ]\s*&?\s*DISCUSSION|"
                   r"R[EÉ]SUM[EÉ]\s+ET\s+DISCUSSION"),
    ("safety", r"TOL[EÉ]RANCE|S[EÉ]CURIT[EÉ]|EFFETS\s+IND[EÉ]SIRABLES|"
               r"ISSUES\s+DES\s+PSUR|PLAN\s+DE\s+GESTION\s+DES\s+RISQUES|\bPGR\b"),
    ("comparators",
     r"M[EÉ]DICAMENTS?\s+COMPARABLES|COMPARATEURS?\s+CLINIQUEMENT|"
     r"COMPARATEURS?\s+NON\s+M[EÉ]DICAMENTEUX|"
     r"M[EÉ]DICAMENTS?\s+DE\s+COMPARAISON|M[EÉ]DICAMENTS?\s+[AÀ]\s+M[EÊ]ME|"
     r"M[EÉ]DICAMENTS?\s+DE\s+M[EÊ]ME\s+CLASSE|EVALUATION\s+CONCURRENTIELLE"),
    ("therapeutic_place", r"PLACE\s+DANS\s+LA\s+STRAT[EÉ]GIE"),
    ("target_population", r"POPULATION\s+CIBLE"),
    ("conclusions", r"CONCLUSIONS?\s+DE\s+LA\s+COMMISSION|^CONCLUSIONS?$"),
    ("recommendations", r"RECOMMANDATIONS"),
    ("medical_need", r"BESOIN\s+(?:M[EÉ]DICAL|TH[EÉ]RAPEUTIQUE)|"
                     r"INT[EÉ]R[EÊ]T\s+DE\s+SANT[EÉ]"),
    ("study_programme", r"PROGRAMME\s+D'?[EÉ]TUDES"),
    ("product_info", r"DONN[EÉ]ES\s+ISSUES\s+DU\s+RCP|"
                     r"INFORMATIONS\s+SUR\s+LE\s+M[EÉ]DICAMENT"),
    ("indication", r"CARACT[EÉ]RISTIQUES\s+DU\s+M[EÉ]DICAMENT|^INDICATIONS?$|"
                   r"INDICATIONS?\s+TH[EÉ]RAPEUTIQUES?|"
                   r"PRINCIPE\s+ACTIF|POSOLOGIE|ORIGINALIT"),
]

CLINICAL = {"clinical_evidence", "safety", "smr", "asmr", "conclusions",
            "therapeutic_place"}


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
    return re.sub(r"[’']", "'", text)[:70]


def role_of(title):
    for role, pat in ROLES:
        if re.search(pat, title, re.IGNORECASE):
            return role
    return "other"


def walk(candidates, first_allowed, max_step):
    accepted, path = [], []
    for page, number, title in candidates:
        try:
            parts = [int(x) for x in number.split(".")]
        except ValueError:
            continue
        depth = len(parts)
        if depth > len(path) + 1:
            continue
        if not path:
            if depth != 1 or not 1 <= parts[0] <= first_allowed:
                continue
        elif depth == len(path) + 1:
            if parts[-1] != 1 or parts[:-1] != path:
                continue
        else:
            if parts[:depth - 1] != path[:depth - 1]:
                continue
            step = parts[-1] - path[depth - 1]
            if not 1 <= step <= max_step:
                continue
        path = parts
        accepted.append((page, number, title))
    return accepted


def validated(candidates):
    strict = walk(candidates, 1, 1)
    if len(strict) >= 4 or len(candidates) < 6:
        return strict, False
    relaxed = walk(candidates, 3, 3)
    if len(relaxed) > len(strict):
        return relaxed, True
    return strict, False


def read_document(path, max_pages, max_chars):
    import pdfplumber
    pages, candidates, blank, synth = [], [], 0, False
    with pdfplumber.open(path) as pdf:
        for number, page in enumerate(pdf.pages[:max_pages], 1):
            try:
                text = page.extract_text() or ""
            except Exception:
                text = ""
            if not text.strip():
                blank += 1
                pages.append("")
                continue
            if number == 1 and SYNTHESE.search(text[:400]):
                synth = True
            repaired = "\n".join(undouble(line) for line in text.splitlines())
            pages.append(repaired)
            lines = repaired.splitlines()
            for index, line in enumerate(lines):
                m = HEADING.match(line)
                if m and len(m.group(2).split()) <= 10:
                    candidates.append((number, m.group(1).strip(),
                                       normalise(m.group(2))))
                    continue
                n = NUMBER_ONLY.match(line)
                if not n:
                    continue
                for ahead in lines[index + 1:index + 3]:
                    if not ahead.strip():
                        continue
                    t = TITLE_LINE.match(ahead)
                    if t and len(t.group(1).split()) <= 10:
                        candidates.append((number, n.group(1).strip(),
                                           normalise(t.group(1))))
                    break
    synopsis = []
    for index, text in enumerate(pages[:6], 1):
        lines = text.splitlines()
        for i, line in enumerate(lines):
            m = SYNOPSIS.match(line)
            if not m:
                continue
            body = " ".join(
                [m.group(2).strip()]
                + [l.strip() for l in lines[i + 1:i + 14] if l.strip()])
            synopsis.append({"label": m.group(1).upper(), "page": index,
                             "verdict": bool(VERDICT.search(body[:400])),
                             "text": body[:1400]})
    good, relaxed = validated(candidates)
    sections = []
    for i, (page, number, title) in enumerate(good):
        depth = len(number.split("."))
        end = len(pages)
        for later_page, later_number, _t in good[i + 1:]:
            if len(later_number.split(".")) <= depth:
                end = later_page
                break
        body = "\n".join(pages[page - 1:end])
        sections.append({"number": number, "title": title,
                         "role": role_of(title), "page": page,
                         "page_end": end,
                         "chars": len(body),
                         "text": body[:max_chars]})
    return {"pages": len(pages), "blank_pages": blank, "synthese": synth,
            "candidates": len(candidates), "sections": sections,
            "synopsis": synopsis, "relaxed": relaxed}


def rank_of(name):
    for pat, rank in PREFERENCE:
        if re.search(pat, name, re.IGNORECASE):
            return rank
    return 99


def plan(workbook, since_year, limit, sample=None, renewals=False,
         seed=20261116, opinions_only=True):
    import pandas as pd
    docs = pd.read_excel(workbook, sheet_name="Appraisal_Documents")
    docs["aid"] = docs["Appraisal_ID"].astype(str)
    docs = docs[docs["aid"].str.startswith("CT")].copy()
    url = docs["Source_URL"].fillna("").astype(str)
    docs["file"] = url.str.rsplit("/", n=1).str[-1]
    docs["url"] = url
    if opinions_only:
        kind = docs["Document_Type"].fillna("").astype(str)
        docs = docs[kind.str.contains(r"Committee\s+opinion", case=False,
                                      regex=True)]
    docs = docs[~docs["file"].str.contains(NOT_AVIS, regex=True)]
    docs = docs.assign(rank=docs["file"].map(rank_of))
    docs = docs[docs["rank"] < 99]
    docs = docs[~docs["aid"].str.contains(PROJET, regex=True)]
    renewal_count = int(docs["file"].str.contains(RENEWAL, regex=True).sum())
    if not renewals:
        docs = docs[~docs["file"].str.contains(RENEWAL, regex=True)]
    year = docs["appraisal_date"].astype(str).str.extract(r"(\d{4})")[0]
    docs = docs.assign(year=year.fillna("0").astype(int))
    docs = docs[docs["year"] >= since_year]
    docs = docs.sort_values(["aid", "rank", "file"])
    docs = docs.drop_duplicates("aid", keep="first")
    chosen = [(r.aid, r.url, int(r.year)) for r in docs.itertuples()]
    if sample:
        import random
        random.seed(seed)
        chosen = random.sample(chosen, min(sample, len(chosen)))
        chosen.sort(key=lambda row: row[2])
    elif limit:
        chosen = chosen[:limit]
    return chosen, renewal_count


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", default=WORKBOOK)
    ap.add_argument("--cache", default=CACHE)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--since", type=int, default=2015)
    ap.add_argument("--limit", type=int)
    ap.add_argument("--sample", type=int)
    ap.add_argument("--include-renewals", action="store_true")
    ap.add_argument("--any-document", action="store_true")
    ap.add_argument("--max-pages", type=int, default=60)
    ap.add_argument("--max-chars", type=int, default=12000)
    ap.add_argument("--pause", type=float, default=1.5)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--show", type=int, default=0)
    args = ap.parse_args()

    for mod in ("pdfplumber", "requests", "pandas"):
        try:
            __import__(mod)
        except ImportError:
            sys.exit(f"pip install {mod}")
    import requests

    chosen, renewals = plan(args.workbook, args.since, args.limit,
                            args.sample, args.include_renewals,
                            opinions_only=not args.any_document)
    if not chosen:
        sys.exit("No full-avis documents matched.")
    if renewals and not args.include_renewals:
        print(f"excluded {renewals:,} renewal (_RI_) documents -- those "
              f"carry an SMR\nverdict but no clinical analysis. "
              f"--include-renewals keeps them.\n")
    years = Counter(y for _a, _u, y in chosen)
    print(f"{len(chosen):,} appraisals with a full avis, {args.since}+")
    for y in sorted(years):
        print(f"  {y}  {years[y]:>5,}")
    if args.dry_run:
        for aid, url, year in chosen[:12]:
            print(f"  {aid:<10} {year}  {url[-70:]}")
        return

    if (args.limit or args.sample) and args.out == OUT:
        print(f"\nthis is a sample run, so writing to {OUT}.sample "
              f"instead of {OUT}\n(pass --out to override)")
        args.out = OUT + ".sample"

    os.makedirs(args.cache, exist_ok=True)
    done, failed = {}, []
    roles, started = Counter(), time.time()
    for i, (aid, url, year) in enumerate(chosen, 1):
        path = os.path.join(args.cache,
                            re.sub(r"[^A-Za-z0-9_.-]", "_", f"{aid}.pdf"))
        if not (os.path.exists(path) and os.path.getsize(path) > 2000):
            try:
                resp = requests.get(
                    url, timeout=180,
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
            got = read_document(path, args.max_pages, args.max_chars)
        except Exception as exc:
            failed.append((aid, f"parse {type(exc).__name__}: {exc}"[:90]))
            continue
        got["url"] = url
        got["year"] = year
        done[aid] = got
        for s in got["sections"]:
            roles[s["role"]] += 1
        if i % 25 == 0 or i == len(chosen):
            left = (len(chosen) - i) * (time.time() - started) / i / 60
            print(f"  [{i:>5}/{len(chosen)}] {len(done):,} parsed, "
                  f"{len(failed)} failed, ~{left:.0f} min left")
            with open(args.out, "w", encoding="utf-8") as fh:
                json.dump(done, fh, ensure_ascii=False)

    with open(args.out, "w", encoding="utf-8") as fh:
        json.dump(done, fh, ensure_ascii=False)

    print(f"\nparsed {len(done):,}, failed {len(failed)}")
    for aid, why in failed[:8]:
        print(f"    {aid}: {why}")
    if not done:
        sys.exit("Nothing parsed.")

    fell_back = sum(1 for d in done.values() if d.get("relaxed"))
    if fell_back:
        print(f"needed the relaxed sequence rule: {fell_back:,}")
    kept = sum(len(d["sections"]) for d in done.values())
    seen = sum(d["candidates"] for d in done.values())
    print(f"heading candidates seen : {seen:,}")
    print(f"survived sequence check : {kept:,} "
          f"({kept / max(1, seen):.0%})")
    synth = [a for a, d in done.items() if d["synthese"]]
    if synth:
        print(f"flagged as SYNTHESE     : {len(synth)}  "
              f"({', '.join(synth[:6])})")
    nosec = [a for a, d in done.items() if not d["sections"]]
    print(f"no sections found        : {len(nosec)}"
          + (f"  ({', '.join(nosec[:6])})" if nosec else ""))

    print("\nsections by role:")
    for role, n in roles.most_common():
        docs = sum(1 for d in done.values()
                   if any(s["role"] == role for s in d["sections"]))
        print(f"  {n:>6}  in {docs:>5} docs ({docs / len(done):>4.0%})  "
              f"{role}")

    full = {a: d for a, d in done.items()
            if not d["synthese"] and d["sections"]}
    syn = sum(1 for d in done.values()
              if any(s.get("verdict") for s in d.get("synopsis", [])))
    print(f"\nfull avis (not a synthese, sections found): {len(full)} of "
          f"{len(done)}")
    print(f"with a front-matter verdict block        : {syn} of "
          f"{len(done)}  {syn / len(done):.0%}")
    covered = sum(1 for a, d in done.items()
                  if any(s["role"] in ("smr", "conclusions")
                         for s in d["sections"])
                  or any(x.get("verdict") for x in d.get("synopsis", [])))
    print(f"verdict findable either way              : {covered} of "
          f"{len(done)}  {covered / len(done):.0%}")
    print(f"\nclinical landmarks, among the {len(full)} full avis:")
    for role in ("clinical_evidence", "smr", "asmr", "conclusions",
                 "comparators", "safety", "therapeutic_place"):
        docs = sum(1 for d in full.values()
                   if any(s["role"] == role for s in d["sections"]))
        print(f"  {role:<20} {docs:>5} of {len(full):>5}  "
              f"{docs / max(1, len(full)):>4.0%}")

    for aid in list(done)[:args.show]:
        d = done[aid]
        print(f"\n{'=' * 68}\n{aid}  ({d['year']}, {d['pages']} pages)")
        for s in d["sections"]:
            flag = "*" if s["role"] in CLINICAL else " "
            print(f" {flag} p{s['page']:>3}-{s['page_end']:<3} "
                  f"{s['number']:<7} {s['role']:<18} {s['title'][:40]}")

    print(f"\n-> {args.out}")
    print("\n* marks sections that carry clinical content. Each has a page")
    print("range, so a quotation from it can cite its page -- which is what")
    print("HAS has been missing.")


if __name__ == "__main__":
    main()
