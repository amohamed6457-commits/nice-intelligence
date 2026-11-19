import argparse
import concurrent.futures as futures
import hashlib
import json
import os
import random
import re
import sys
import time
from collections import Counter

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SECTIONS = "has_clinical_sections.json"
OUT = "has_endpoints.json"
FAILLOG = "has_endpoint_failures.json"

CATEGORIES = ["mortality", "morbidity", "quality_of_life", "safety",
              "surrogate", "pharmacokinetic", "other"]

DIRECTIONS = ["favours_drug", "favours_comparator", "no_difference",
              "not_stated"]

SYSTEM = """You read one section of a French HTA assessment (an avis of the
Commission de la Transparence) and list the clinical endpoints it reports.

Return ONLY a JSON array. No prose, no code fence, no explanation.

Each element is an object with exactly these keys:
  "endpoint"   the endpoint as the document names it, in French
  "endpoint_en" the same endpoint in English
  "category"   one of: mortality, morbidity, quality_of_life, safety,
               surrogate, pharmacokinetic, other
  "study"      the study name or code if the section gives one, else ""
  "comparator" what the drug was compared against, else ""
  "result"     the quantitative result exactly as written, including any
               HR, OR, RR, difference, confidence interval and p-value.
               Empty string if the section reports no number for it.
  "direction"  one of: favours_drug, favours_comparator, no_difference,
               not_stated
  "primary"    true if the section calls it a primary endpoint, else false

RULES
1. Only endpoints the section actually reports. Do not add endpoints you
   would expect to see.
2. Return [] when the section reports no clinical endpoint. That is common
   and correct: bioequivalence studies, generics, and older avis often
   contain no comparative clinical data. An empty array is a valid answer
   and is better than inventing one.
3. A bioequivalence or pharmacokinetic comparison is category
   "pharmacokinetic", not mortality or morbidity.
4. Copy numbers exactly, including French decimal commas. Never convert,
   round or recompute.
5. "survie globale" is mortality. "survie sans progression", tumour
   response and biomarker measures are surrogate. Symptom scores, function
   scores and event rates are morbidity. Named instruments such as
   EORTC QLQ-C30, FACT-O, EQ-5D are quality_of_life. Adverse events,
   tolerance and discontinuation are safety.
6. If the same endpoint is reported for several studies or subgroups,
   return one element per study.
7. Do not summarise the Commission's opinion. Only the endpoints and their
   results."""


def key_of(text):
    return hashlib.sha1(text.strip().encode("utf-8")).hexdigest()[:16]


def load(path, default=None):
    if not os.path.exists(path):
        return default
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (ValueError, OSError) as exc:
        sys.exit(f"{path} could not be read ({exc}).")


def save(path, payload):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False)
    os.replace(tmp, path)


def api_key():
    for var in ("ANTHROPIC_API_KEY", "CLAUDE_API_KEY"):
        if os.environ.get(var):
            return os.environ[var]
    secrets = os.path.join(".streamlit", "secrets.toml")
    if os.path.exists(secrets):
        pattern = re.compile(r'^\s*[\w.-]*api[_-]?key\s*=\s*["\'](.+?)["\']',
                             re.IGNORECASE)
        for line in open(secrets, encoding="utf-8"):
            m = pattern.match(line)
            if m and m.group(1).strip():
                return m.group(1).strip()
    sys.exit("No API key. Set ANTHROPIC_API_KEY or put it in "
             ".streamlit/secrets.toml.")


def parse_array(raw):
    text = str(raw or "").strip()
    text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.MULTILINE).strip()
    start = text.find("[")
    if start < 0:
        return None
    depth, end = 0, None
    for i, ch in enumerate(text[start:], start):
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                end = i + 1
                break
    if end is None:
        return None
    try:
        got = json.loads(text[start:end])
    except ValueError:
        return None
    return got if isinstance(got, list) else None


def clean(items, section):
    out = []
    for item in items:
        if not isinstance(item, dict):
            continue
        name = str(item.get("endpoint") or "").strip()
        if not name:
            continue
        category = str(item.get("category") or "other").strip().lower()
        direction = str(item.get("direction") or "not_stated").strip().lower()
        out.append({
            "endpoint": name[:200],
            "endpoint_en": str(item.get("endpoint_en") or "").strip()[:200],
            "category": category if category in CATEGORIES else "other",
            "study": str(item.get("study") or "").strip()[:120],
            "comparator": str(item.get("comparator") or "").strip()[:200],
            "result": str(item.get("result") or "").strip()[:300],
            "direction": direction if direction in DIRECTIONS else "not_stated",
            "primary": bool(item.get("primary")),
            "page": section["page"],
            "page_end": section["page_end"],
            "section": section["number"],
        })
    return out


WANTED_ROLES = ("clinical_evidence", "safety", "discussion")


def sections_of(doc, min_chars=300):
    best = {}
    for s in doc.get("sections", []):
        if s["role"] not in WANTED_ROLES:
            continue
        if s["chars"] < min_chars or not s["text"].strip():
            continue
        kept = best.get(s["role"])
        if kept is None or s["chars"] > kept["chars"]:
            best[s["role"]] = s
    return [best[r] for r in WANTED_ROLES if r in best]


def extract_one(client, model, section, max_retry=3):
    delay, last = 2.0, "no response"
    body = section["text"]
    for attempt in range(max_retry):
        try:
            resp = client.messages.create(
                model=model,
                max_tokens=4000,
                temperature=0 if attempt == 0 else 0.2,
                system=SYSTEM,
                messages=[{"role": "user", "content": body}],
            )
            raw = "".join(b.text for b in resp.content
                          if b.type == "text")
            got = parse_array(raw)
            if got is not None:
                return clean(got, section)
            stop = getattr(resp, "stop_reason", None)
            last = (f"truncated at 4000 tokens" if stop == "max_tokens"
                    else f"no JSON array in {raw[:80]!r}")
        except Exception as exc:
            last = f"{type(exc).__name__}: {exc}"
        if attempt == max_retry - 1:
            break
        time.sleep(delay * (0.5 + random.random()))
        delay = min(delay * 2, 20)
    return f"__ERROR__ {last}"


def audit(done):
    flat = [e for v in done.values() if isinstance(v, list) for e in v]
    empty = sum(1 for v in done.values() if isinstance(v, list) and not v)
    appraisals = {k.split("#")[0] for k in done}
    with_any = {k.split("#")[0] for k, v in done.items()
                if isinstance(v, list) and v}
    print(f"sections read        : {len(done):,}")
    print(f"  returned endpoints : {len(done) - empty:,}")
    print(f"  returned []        : {empty:,}  "
          f"({empty / max(1, len(done)):.0%})")
    print(f"appraisals covered   : {len(appraisals):,}")
    print(f"  with >=1 endpoint  : {len(with_any):,}  "
          f"({len(with_any) / max(1, len(appraisals)):.0%})")
    print(f"endpoints extracted  : {len(flat):,}")
    if not flat:
        return
    print(f"  per appraisal      : "
          f"{len(flat) / max(1, len(with_any)):.1f} average")
    print(f"  per section read   : "
          f"{len(flat) / max(1, len(done) - empty):.1f} average\n")
    print("by category:")
    for cat, n in Counter(e["category"] for e in flat).most_common():
        hit = {k.split("#")[0] for k, v in done.items()
               if isinstance(v, list)
               and any(e["category"] == cat for e in v)}
        print(f"  {n:>7,}  in {len(hit):>5,} appraisals   {cat}")
    print("\nby direction:")
    for d, n in Counter(e["direction"] for e in flat).most_common():
        print(f"  {n:>7,}  {d}")
    quant = sum(1 for e in flat if e["result"])
    named = sum(1 for e in flat if e["study"])
    prim = sum(1 for e in flat if e["primary"])
    print(f"\nwith a quantitative result : {quant:,} "
          f"({quant / len(flat):.0%})")
    print(f"with a named study         : {named:,} "
          f"({named / len(flat):.0%})")
    print(f"marked primary             : {prim:,} "
          f"({prim / len(flat):.0%})")
    print("\ncommonest endpoints (English):")
    for name, n in Counter(e["endpoint_en"].lower() for e in flat
                           if e["endpoint_en"]).most_common(12):
        print(f"  {n:>6,}  {name[:58]}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sections", default=SECTIONS)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--model", default="claude-haiku-4-5-20251001")
    ap.add_argument("--sample", type=int)
    ap.add_argument("--limit", type=int)
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--seed", type=int, default=20261117)
    ap.add_argument("--min-year", type=int)
    ap.add_argument("--max-year", type=int)
    ap.add_argument("--audit", action="store_true")
    args = ap.parse_args()

    if args.audit:
        done = load(args.out, {})
        if not done:
            sys.exit(f"{args.out} is empty. Nothing extracted yet.")
        audit(done)
        return

    data = load(args.sections)
    if data is None:
        sys.exit(f"{args.sections} not found. Run extract_has_sections.py "
                 f"first.")

    todo = []
    for aid, doc in sorted(data.items()):
        if doc.get("synthese"):
            continue
        year = int(doc.get("year") or 0)
        if args.min_year and year < args.min_year:
            continue
        if args.max_year and year > args.max_year:
            continue
        for section in sections_of(doc):
            todo.append((aid, section))

    print(f"{len(data):,} appraisals in {args.sections}")
    print(f"sections to read         : {len(todo):,} across "
          f"{len({a for a, _s in todo}):,} appraisals")
    by_role = Counter(s["role"] for _a, s in todo)
    for role, n in by_role.most_common():
        print(f"    {n:>6,}  {role}")

    if args.sample:
        random.seed(args.seed)
        ids = sorted({a for a, _s in todo})
        keep = set(random.sample(ids, min(args.sample, len(ids))))
        todo = [(a, s) for a, s in todo if a in keep]
    elif args.limit:
        todo = todo[:args.limit]

    if (args.sample or args.limit) and args.out == OUT:
        args.out = OUT + ".sample"
        print(f"sample run, writing to {args.out}")

    done = load(args.out, {})
    failures = load(FAILLOG, {})
    todo = [(a, s) for a, s in todo
            if f"{a}#{s['number']}" not in done]
    print(f"already extracted        : {len(done):,}")
    print(f"to do now                : {len(todo):,}\n")
    if not todo:
        print("Nothing to do. Run --audit for the distribution.")
        return

    try:
        import anthropic
    except ImportError:
        sys.exit("pip install anthropic")
    client = anthropic.Anthropic(api_key=api_key())

    started, errors = time.time(), 0
    with futures.ThreadPoolExecutor(max_workers=args.workers) as pool:
        jobs = {pool.submit(extract_one, client, args.model, s):
                (a, s["number"]) for a, s in todo}
        for i, fut in enumerate(futures.as_completed(jobs), 1):
            aid, number = jobs[fut]
            slot = f"{aid}#{number}"
            try:
                got = fut.result()
            except Exception as exc:
                got = f"__ERROR__ {type(exc).__name__}: {exc}"
            if isinstance(got, str):
                errors += 1
                failures[slot] = got[10:].strip()
            else:
                done[slot] = got
                failures.pop(slot, None)
            if i % 25 == 0 or i == len(todo):
                save(args.out, done)
                save(FAILLOG, failures)
                rate = (time.time() - started) / i
                print(f"  [{i:>5}/{len(todo)}] {len(done):,} done, "
                      f"{errors} failed, "
                      f"~{(len(todo) - i) * rate / 60:.0f} min left")

    save(args.out, done)
    save(FAILLOG, failures)
    print(f"\n{len(done):,} appraisals -> {args.out}")
    if errors:
        print(f"{errors} failed -> {FAILLOG}. Re-run to retry them.")
    print()
    audit(done)
    print("\nBenchmark this against ten appraisals you read yourself before")
    print("any of it reaches the tool.")


if __name__ == "__main__":
    main()
