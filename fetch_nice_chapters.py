"""
Stage 1 of the chapter-first pipeline: discover and fetch NICE committee
discussion chapters.

Established by survey of 12 appraisals:
  * Chapter NAME is stable ('N-Committee-discussion') but N varies (3 or 4),
    so chapter links must be discovered from the overview page, not constructed.
  * Appraisals below ~TA400 publish no chapters at all. Don't waste requests.
  * 'committee' alone is too loose a match: it also hits
    'N-Evaluation-committee-members-and-NICE-project-team', a list of names.
  * Newer appraisals under the evaluation process may name things differently,
    so several variants are scored rather than one pattern assumed.

Also prints a structural analysis of the fetched markup, so the chunk parser
can be written against real HTML instead of assumptions.

Usage:
  python fetch_nice_chapters.py --analyse 805 993 710      # fetch + describe markup
  python fetch_nice_chapters.py <excel> --limit 20 --min-ta 400
"""

import argparse
import json
import os
import re
import time
from collections import Counter
from urllib.parse import urljoin

import requests

BASE = "https://www.nice.org.uk"
HTML_DIR = "nice_chapters"
INDEX_FILE = "chapter_index.json"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-GB,en;q=0.9",
}

CHAPTER_HREF = re.compile(r'href="(/guidance/[Tt][Aa]\d+/chapter/[^"#]+)"')

# Scored highest-first. Deliberately excludes member/project-team pages.
# The bare '<n>-Discussion' form appears in NICE's newer evaluation process
# (e.g. TA1152 semaglutide), where the word 'committee' is dropped entirely.
# Anchoring these with $ was a real bug: NICE publishes slugs like
# "4-Committee-discussion-during-the-original-appraisal", which is plainly a
# committee discussion chapter and which "committee-discussion$" rejects.
# Matching a PREFIX and letting the blocklist exclude the wrong chapters is
# both more permissive and easier to reason about.
CHAPTER_PATTERNS = [
    (100, r"\d+-committee-discussion\b"),
    (98,  r"\d+-discussion\b"),
    (95,  r"\d+-evaluation-committee-discussion\b"),
    (90,  r"\d+-committee-discussion-and-recommendations\b"),
    (85,  r"\d+-the-committee(?:'s|s)?-discussion\b"),
    (80,  r"\d+-committee-considerations?\b"),
    (78,  r"\d+-(?:the-)?committees?-(?:view|conclusion|consideration)"),
    (75,  r"\d+-(?:the-)?evaluation\b"),
    (40,  r"\d+-(?:evidence|clinical|cost)[-\w]*$"),
]
# Order matters against CHAPTER_PATTERNS: the blocklist runs first, so a
# term here beats any pattern. "recommendations$" used to kill
# "3-Committee-discussion-and-recommendations", which is a discussion
# chapter. Block a chapter that is ONLY recommendations, not one that
# merely ends with the word.
CHAPTER_BLOCKLIST = re.compile(
    r"members|project-team|implementation|information-about|"
    r"^\d+-recommendations?$|appendi|glossary|references",
    re.IGNORECASE,
)


def log(m):
    print(m, flush=True)


def get(session, url):
    try:
        r = session.get(url, headers=HEADERS, timeout=25, allow_redirects=True)
    except requests.exceptions.RequestException as e:
        return None, type(e).__name__
    if r.status_code != 200:
        return None, f"HTTP {r.status_code}"
    return r.text, None


def discover_chapters(session, ta):
    """Return (ordered_chapter_paths, error)."""
    html, err = get(session, f"{BASE}/guidance/ta{ta}")
    if err:
        return [], err
    paths = list(dict.fromkeys(CHAPTER_HREF.findall(html)))
    return paths, None if paths else "no chapter links on overview"


def pick_discussion_chapter(paths):
    """Score candidates; return (path, score) or (None, None)."""
    best, best_score = None, 0
    for p in paths:
        tail = p.rstrip("/").rsplit("/", 1)[-1]
        if CHAPTER_BLOCKLIST.search(tail):
            continue
        for score, pattern in CHAPTER_PATTERNS:
            if re.search(pattern, tail, re.IGNORECASE):
                if score > best_score:
                    best, best_score = p, score
                break
    return best, (best_score or None)


def analyse_markup(html, ta):
    """Describe how numbered sections are expressed, to inform the parser."""
    log(f"\n  --- markup analysis TA{ta} ---")
    log(f"  html length: {len(html):,}")

    # Which tags hold text starting with a NICE section number like 3.14 ?
    tag_hits = Counter()
    samples = []
    for m in re.finditer(r"<(p|li|div|h[1-6])\b([^>]*)>(.{0,220}?)</\1>", html,
                         re.IGNORECASE | re.DOTALL):
        tag, attrs, inner = m.group(1).lower(), m.group(2), m.group(3)
        text = re.sub(r"<[^>]+>", "", inner).strip()
        if re.match(r"^\d+\.\d+\b", text):
            tag_hits[tag] += 1
            if len(samples) < 4:
                samples.append((tag, attrs.strip()[:90], text[:150]))

    if tag_hits:
        log(f"  tags holding numbered sections: {dict(tag_hits)}")
        for tag, attrs, text in samples:
            log(f"    <{tag} {attrs}>")
            log(f"      {text}")
    else:
        log("  no tag directly wraps text beginning '<n>.<n>'")
        nums = re.findall(r"\b\d+\.\d+\b", html)[:12]
        loose_count = len(re.findall(r"[>\s]\d+\.\d+\s", html))
        log(f"  bare section-number occurrences in html: {loose_count}")
        log(f"  first few seen: {nums}")

    # Does an id/name anchor carry the section number? Useful for deep links.
    anchors = re.findall(r'(?:id|name)="([^"]{0,40}\d+[-.]\d+[^"]{0,20})"', html)
    log(f"  anchors containing a section number: {len(anchors)}"
        + (f"  e.g. {anchors[:5]}" if anchors else ""))

    # Main content container, so the parser can scope itself.
    for pat in (r'<(article|main)\b[^>]*>', r'<div[^>]*class="[^"]*(?:chapter|content|guidance)[^"]*"[^>]*>'):
        found = re.findall(pat, html, re.IGNORECASE)[:3]
        if found:
            log(f"  container candidates: {[f[:80] if isinstance(f,str) else f for f in found]}")
            break


def fetch_one(session, ta, analyse=False):
    paths, err = discover_chapters(session, ta)
    if err:
        return False, err

    chosen, score = pick_discussion_chapter(paths)
    if not chosen:
        tails = [p.rsplit("/", 1)[-1][:40] for p in paths]
        return False, f"no discussion chapter among {len(paths)}: {tails}"

    time.sleep(0.4)
    html, err = get(session, urljoin(BASE, chosen))
    if err:
        return False, f"chapter fetch {err}"

    os.makedirs(HTML_DIR, exist_ok=True)
    out = os.path.join(HTML_DIR, f"TA{str(ta).zfill(4)}_discussion.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)

    if analyse:
        analyse_markup(html, ta)

    return True, {
        "ta": str(ta),
        "chapter_path": chosen,
        "match_score": score,
        "file": out,
        "chars": len(html),
        "n_chapters": len(paths),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("excel", nargs="?", default="NICE_v14_updated_2026-09-28.xlsx")
    ap.add_argument("--analyse", nargs="+", metavar="TA",
                    help="fetch these TAs and describe their markup, then exit")
    ap.add_argument("--limit", type=int, default=20)
    ap.add_argument("--min-ta", type=int, default=400,
                    help="below ~TA400 NICE publishes no chapters (default 400)")
    ap.add_argument("--delay", type=float, default=1.2)
    ap.add_argument("--retry-failed", action="store_true",
                    help="re-attempt the TAs recorded as failed. Without this "
                         "they are skipped forever, because a failure is "
                         "remembered exactly like a success")
    ap.add_argument("--why-failed", action="store_true",
                    help="group the recorded failures by reason and exit, so "
                         "you can see whether retrying is worth the time")
    args = ap.parse_args()

    session = requests.Session()

    if args.analyse:
        for ta in args.analyse:
            ta = ta.upper().replace("TA", "").lstrip("0")
            log(f"\nTA{ta}")
            ok, info = fetch_one(session, ta, analyse=True)
            if ok:
                log(f"  saved {info['file']}  ({info['chars']:,} chars)")
                log(f"  chapter: {info['chapter_path']}  score={info['match_score']}")
            else:
                log(f"  failed: {info}")
            time.sleep(args.delay)
        log("\nPaste the markup analysis back so the parser can be written "
            "against real HTML.")
        return

    import pandas as pd
    try:
        df = pd.read_excel(args.excel, sheet_name="Sheet1")
    except FileNotFoundError:
        log(f"Excel not found: {args.excel}")
        return
    nums = sorted({
        int(s) for s in (str(i).upper().replace("TA", "").strip()
                         for i in df["appraisal_id"].dropna().unique())
        if s.isdigit()
    })

    index = {}
    if os.path.exists(INDEX_FILE):
        try:
            index = json.load(open(INDEX_FILE))
        except (json.JSONDecodeError, OSError):
            index = {}
    done = index.get("done", {})
    failed = index.get("failed", {})

    if args.why_failed:
        import collections
        # Group on the reason's SHAPE, not its full text, or every appraisal
        # is its own category because the drug name is in the chapter slug.
        groups = collections.defaultdict(list)
        for ta, info in sorted(failed.items(), key=lambda kv: int(kv[0])):
            text = (info if isinstance(info, str) else str(info)).strip()
            if "no chapter links" in text:
                key = "no chapter links found on the overview page"
            elif "no discussion chapter" in text:
                key = "chapters found, but none matched a discussion pattern"
            else:
                key = text[:70]
            groups[key].append((ta, text))

        log(f"{len(failed)} recorded failures\n")
        for key, items in sorted(groups.items(), key=lambda kv: -len(kv[1])):
            log(f"{len(items):>4}x  {key}")
            log(f"        TAs: {', '.join('TA' + t for t, _ in items[:10])}"
                + (f" ... (+{len(items) - 10} more)" if len(items) > 10 else ""))
            # One full, untruncated example. The chapter list is the evidence;
            # cutting it off is what hid the cause last time.
            if len(items[0][1]) > 80:
                log(f"        full example (TA{items[0][0]}):")
                log(f"          {items[0][1]}")
            log("")

        probe = [t for t, _ in groups.get(
            "no chapter links found on the overview page", [])][:3]
        if probe:
            log("Investigate the biggest group with:")
            log("  python fetch_nice_chapters.py --analyse "
                + " ".join(probe))
        log(f"\nRetry all with:  --retry-failed --limit {len(failed)}")
        return

    skip = set(done)
    if not args.retry_failed:
        skip |= set(failed)
    todo = [n for n in nums if n >= args.min_ta and str(n) not in skip]
    todo.sort(reverse=True)

    log(f"In workbook: {len(nums)}   at/above TA{args.min_ta}: "
        f"{sum(1 for n in nums if n >= args.min_ta)}")
    log(f"Fetched: {len(done)}   failed: {len(failed)}   queued: {len(todo)}")
    if args.retry_failed:
        log(f"retrying failures; --limit is {args.limit}, so raise it to cover "
            f"all {len(failed)} of them")
    log("")

    batch = todo[: args.limit]
    ok_n = bad_n = 0
    for i, ta in enumerate(batch, 1):
        log(f"[{i}/{len(batch)}] TA{str(ta).zfill(4)}")
        ok, info = fetch_one(session, ta)
        if ok:
            done[str(ta)] = info
            failed.pop(str(ta), None)
            ok_n += 1
            log(f"    {info['chapter_path'].rsplit('/',1)[-1]}  {info['chars']:,} chars")
        else:
            failed[str(ta)] = info
            bad_n += 1
            log(f"    -- {info}")
        if i % 5 == 0:
            json.dump({"done": done, "failed": failed}, open(INDEX_FILE, "w"), indent=2)
        time.sleep(args.delay)

    json.dump({"done": done, "failed": failed}, open(INDEX_FILE, "w"), indent=2)
    log(f"\n{'='*46}\nfetched {ok_n}, failed {bad_n}, total {len(done)}")


if __name__ == "__main__":
    main()
