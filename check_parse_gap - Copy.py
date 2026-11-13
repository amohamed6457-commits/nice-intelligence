"""
Why did 22 fetched appraisals parse to zero sections?

chapter_index.json lists them as done — the fetch worked — but nothing for
them reached chapter_sections.json. TA959 has committee documents on
nice.org.uk, so this is not missing data. It is HTML already sitting in
nice_chapters/ that the parse step could not split.

A chapter that parsed correctly looks like this:

    <h2>3 Committee discussion</h2>
      <h3 id="inpagenav-new-treatment-option">New treatment option</h3>
      <h3 id="inpagenav-clinical-effectiveness">Clinical effectiveness</h3>
      ...

so the parser is splitting on the h3 subsections. A page with no h3s — one
continuous discussion, or a newer layout using different tags — yields zero
sections even though the text is all there.

This opens each affected file and reports what it actually contains, which
distinguishes the three possibilities:

  * the file is missing or empty      -> refetch it
  * it has text but no h3 subsections -> change the parser, no refetch
  * it is an error or cookie page     -> refetch it

    python check_parse_gap.py
    python check_parse_gap.py --ta 959 --dump
"""

import argparse
import glob
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from why_no_chapters import _BUCKETS, buckets, load, records
except ImportError as exc:
    sys.exit(f"why_no_chapters.py must sit beside this script ({exc}).")

INDEX = "chapter_index.json"
SECTIONS = "chapter_sections.json"
CHAPTER_DIRS = ["nice_chapters", "chapters", "."]


def strip_tags(raw):
    raw = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", raw)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", raw)).strip()


def find_files(ta):
    """Any local HTML whose name carries this TA number."""
    hits = []
    for d in CHAPTER_DIRS:
        if not os.path.isdir(d):
            continue
        for pattern in (f"*{ta}*.html", f"*ta{ta}*.html", f"*TA{ta}*.html",
                        f"*{int(ta):04d}*.html"):
            hits.extend(glob.glob(os.path.join(d, pattern)))
    # A bare number matches too loosely (959 inside 1959), so keep only
    # files where it stands alone.
    keep, seen = [], set()
    for h in hits:
        base = os.path.basename(h)
        if h in seen:
            continue
        seen.add(h)
        if re.search(rf"(?<!\d){int(ta)}(?!\d)", base) or \
           re.search(rf"(?<!\d){int(ta):04d}(?!\d)", base):
            keep.append(h)
    return sorted(keep)


def inspect(path):
    raw = open(path, encoding="utf-8", errors="ignore").read()
    text = strip_tags(raw)
    h2 = [re.sub(r"<[^>]+>", "", " ".join(h.split()))[:60]
          for h in re.findall(r"(?is)<h2[^>]*>(.*?)</h2>", raw)]
    h3 = [re.sub(r"<[^>]+>", "", " ".join(h.split()))[:60]
          for h in re.findall(r"(?is)<h3[^>]*>(.*?)</h3>", raw)]
    h4 = len(re.findall(r"(?is)<h4[^>]*>", raw))
    nav = len(re.findall(r'id="inpagenav-', raw))
    return {"bytes": len(raw), "text_chars": len(text), "h2": h2, "h3": h3,
            "h4": h4, "inpagenav": nav, "raw": raw, "text": text}


def verdict(info):
    if info["bytes"] < 2000 or info["text_chars"] < 400:
        return "too small — looks like an error or redirect page; refetch"
    if info["inpagenav"] or info["h3"]:
        return (f"has {len(info['h3'])} h3 / {info['inpagenav']} inpagenav — "
                "should have parsed; check the parser's chapter filter")
    if any("committee" in h.lower() or "discussion" in h.lower()
           for h in info["h2"]):
        return ("committee text present but NO h3 subsections — the parser "
                "splits on h3, so it found nothing. Fix the parser, do not refetch")
    if info["h4"]:
        return f"no h3 but {info['h4']} h4 headings — parser is splitting on the wrong tag"
    return "no recognisable section headings — inspect with --dump"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--index", default=INDEX)
    ap.add_argument("--sections", default=SECTIONS)
    ap.add_argument("--ta", help="inspect one appraisal instead of all 22")
    ap.add_argument("--dump", action="store_true",
                    help="print the headings and first text of each file")
    ap.add_argument("--limit", type=int, default=25)
    args = ap.parse_args()

    index = load(args.index)
    if index is None:
        sys.exit(f"{args.index} not found. Run this from the folder that holds it.")

    grouped = buckets(index)
    rows, origin = {}, {}
    if grouped:
        for name, entries in grouped:
            for key, value in entries.items():
                rows[key], origin[key] = value, str(name).lower()
    else:
        rows = dict(records(index))

    have = set()
    sections = load(args.sections)
    if sections is not None:
        for key, value in records(sections):
            if value:
                have.add(str(key).strip().lower())

    if args.ta:
        targets = [str(args.ta).strip().lower().lstrip("ta")]
    else:
        targets = [k for k in rows
                   if _BUCKETS.get(origin.get(k, ""), None) == "ok"
                   and str(k).strip().lower() not in have]
        targets.sort(key=lambda k: -int(re.sub(r"\D", "", k) or 0))

    if not targets:
        sys.exit("No appraisals are marked done without parsed sections. "
                 "Nothing to explain.")

    found = next((d for d in CHAPTER_DIRS if os.path.isdir(d)), None)
    print(f"{len(targets)} appraisal(s) fetched but unparsed")
    print(f"chapter html directory: {found or 'NOT FOUND — is nice_chapters/ here?'}\n")

    tally = {}
    for ta in targets[:args.limit]:
        n = re.sub(r"\D", "", str(ta))
        files = find_files(n) if n else []
        if not files:
            line, v = f"TA{n}", "no local HTML found — refetch"
        else:
            info = inspect(files[0])
            v = verdict(info)
            line = (f"TA{n}  {info['bytes']:>7,}b  h2={len(info['h2'])} "
                    f"h3={len(info['h3'])} nav={info['inpagenav']}")
            if args.dump:
                print(f"\n{line}\n   verdict: {v}")
                print(f"   h2: {info['h2'][:4]}")
                print(f"   h3: {info['h3'][:8]}")
                print(f"   text: {info['text'][:220]}…")
                tally[v] = tally.get(v, 0) + 1
                continue
        tally[v] = tally.get(v, 0) + 1
        print(f"  {line:<46} {v}")

    print("\nsummary:")
    for v, n in sorted(tally.items(), key=lambda kv: -kv[1]):
        print(f"  {n:4d}  {v}")
    print("\nIf most say the text is present but there are no h3 subsections, "
          "these are\nrecoverable from the HTML you already have — no refetching, "
          "no network.")


if __name__ == "__main__":
    main()
