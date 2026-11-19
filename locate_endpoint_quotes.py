import argparse
import gc
import json
import os
import re
import sys
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

NUMBER = re.compile(r"\d+(?:[.,]\d+)?")
SENTENCE = re.compile(
    r"(?<=[.!?])\s+(?=[A-ZÀ-Ü])"
    r"|\s[-\u2013\u2014\u2022]\s+"
    r"|\n\s*[-\u2013\u2014\u2022]\s*"
    r"|\n{2,}")
DECIMAL = re.compile(r"(?<=\d)[.,](?=\d)")
GUARD = "\x00"
WS = re.compile(r"\s+")

NOISE = {"95", "0", "1", "2", "3", "5", "100"}


def norm_id(x):
    return re.sub(r"[^A-Za-z0-9]+", "", str(x or "")).upper()


def numbers_in(text):
    clean = str(text or "").replace(GUARD, ",")
    return [n.replace(",", ".").rstrip(".") for n in NUMBER.findall(clean)]


def signal_numbers(text):
    nums = [n for n in numbers_in(text)]
    strong = [n for n in nums if n not in NOISE]
    return strong or nums


def tidy(text):
    return WS.sub(" ", str(text or "")).strip()


def sentences(text):
    safe = DECIMAL.sub(GUARD, str(text or ""))
    parts = [tidy(p).replace(GUARD, ",") for p in SENTENCE.split(safe)]
    return [p for p in parts if len(p) > 20]


def best_sentence(section_text, result, endpoint):
    wanted = signal_numbers(result)
    if not wanted:
        return None, 0.0, 0
    target = set(wanted)
    words = [w for w in re.findall(r"[a-zà-ü]{5,}", str(endpoint or "").lower())]
    best, best_score, best_hits = None, 0.0, 0
    for sent in sentences(section_text):
        have = set(numbers_in(sent))
        hits = len(target & have)
        if not hits:
            continue
        score = hits / len(target)
        lowered = sent.lower()
        score += 0.08 * sum(1 for w in words if w in lowered)
        if score > best_score:
            best, best_score, best_hits = sent, score, hits
    return best, min(best_score, 1.0), best_hits


def pad(section_text, sentence, before=1, after=1):
    parts = sentences(section_text)
    if sentence not in parts:
        return sentence
    i = parts.index(sentence)
    lo = max(0, i - before)
    hi = min(len(parts), i + after + 1)
    return " ".join(parts[lo:hi])


def shortest_complete(section_text, result, endpoint):
    """Prefer the tightest fragment carrying every figure: a list item about
    one endpoint beats the paragraph that runs five of them together."""
    sent, score, hits = best_sentence(section_text, result, endpoint)
    if sent is None:
        return None, 0.0, 0, False
    total = len(signal_numbers(result))
    complete = hits >= total
    return sent, score, hits, complete


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sections", default="has_clinical_sections.json")
    ap.add_argument("--endpoints", default="has_endpoints_full.json")
    ap.add_argument("--out", default="has_endpoint_quotes.json")
    ap.add_argument("--context", type=int, default=1)
    ap.add_argument("--max-chars", type=int, default=700)
    ap.add_argument("--min-hits", type=int, default=1)
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    for path in (args.sections, args.endpoints):
        if not os.path.exists(path):
            raise SystemExit(f"{path} not found - run this beside your data")

    done = json.load(open(args.endpoints, encoding="utf-8"))
    print(f"{args.endpoints}: {len(done):,} sections")

    needed = set()
    for key, items in done.items():
        if isinstance(items, list) and items:
            aid, _, number = str(key).partition("#")
            needed.add((norm_id(aid), number))
    print(f"sections carrying endpoints: {len(needed):,}")

    size = os.path.getsize(args.sections) / 1e6
    print(f"\nreading {args.sections} ({size:,.0f} MB) - this takes a minute")
    sections = json.load(open(args.sections, encoding="utf-8"))
    text_of = {}
    seen = 0
    for aid, doc in sections.items():
        key_aid = norm_id(aid)
        for sec in doc.get("sections", []):
            seen += 1
            pair = (key_aid, str(sec.get("number")))
            if pair in needed:
                text_of[pair] = sec.get("text", "")
    del sections
    gc.collect()
    print(f"  {seen:,} sections in file, {len(text_of):,} kept\n")
    if not text_of:
        raise SystemExit(
            "no section text matched. Run probe_sections.py on this file to "
            "check its shape, and confirm it is the file extract_has_endpoints "
            "was run against.")

    quotes = {}
    stat = Counter()
    missing_sections = set()
    for i, (key, items) in enumerate(done.items()):
        if args.limit and i >= args.limit:
            break
        if not isinstance(items, list) or not items:
            continue
        aid, _, number = str(key).partition("#")
        body = text_of.get((norm_id(aid), number))
        if body is None:
            missing_sections.add(key)
            stat["no section text"] += len(items)
            continue
        for pos, e in enumerate(items):
            result = (e.get("result") or "").strip()
            if not result:
                stat["no result to match"] += 1
                continue
            sent, score, hits = best_sentence(body, result,
                                              e.get("endpoint_en") or
                                              e.get("endpoint"))
            if sent is None or hits < args.min_hits:
                stat["figures not found"] += 1
                continue
            complete = hits >= len(signal_numbers(result))
            pad_by = 0 if complete else args.context
            quote = pad(body, sent, pad_by, pad_by)
            if len(quote) > args.max_chars:
                quote = sent[:args.max_chars]
            quotes[f"{key}|{pos}"] = {
                "quote": quote,
                "match": round(score, 2),
                "figures_matched": hits,
                "figures_total": len(signal_numbers(result)),
            }
            stat["located" if score >= 0.6 else "partial match"] += 1

    json.dump(quotes, open(args.out, "w", encoding="utf-8"),
              ensure_ascii=False)

    total = sum(stat.values())
    print(f"{total:,} endpoints examined")
    for label in ("located", "partial match", "figures not found",
                  "no result to match", "no section text"):
        if stat[label]:
            print(f"  {label:<22} {stat[label]:>7,}  "
                  f"({stat[label] / total:>4.0%})")
    checkable = stat["located"] + stat["partial match"] + stat["figures not found"]
    if checkable:
        found = stat["located"] + stat["partial match"]
        print(f"\nof {checkable:,} endpoints whose figures could be checked, "
              f"{found:,} ({found / checkable:.0%}) appear in the source text")
    if missing_sections:
        print(f"\n{len(missing_sections):,} sections had no text in "
              f"{args.sections}; e.g. {sorted(missing_sections)[:3]}")
    size = os.path.getsize(args.out) / 1e6
    print(f"\n{len(quotes):,} quotes -> {args.out}  ({size:.1f} MB)")


if __name__ == "__main__":
    main()
