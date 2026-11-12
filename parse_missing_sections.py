import argparse
import glob
import json
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from why_no_chapters import _BUCKETS, buckets, load, records
except ImportError as exc:
    sys.exit(f"why_no_chapters.py must sit beside this script ({exc}).")

INDEX = "chapter_index.json"
SECTIONS = "chapter_sections.json"
CHAPTER_DIRS = ["nice_chapters", "chapters"]

# NICE publishes committee discussions in two markups, and both appear among
# the 22. They disagree about the tag, the class and where the number lives:
#
#   <div id="ta710_3_1">
#     <p class="numbered-paragraph"><span class="paragraph-number">3.1 </span>text
#
#   <article id="ta993-3_1" class="numbered-paragraph">
#     <h5>3.1</h5><div><p>text
#
# What they agree on is the id: ta<number><sep><major><sep><minor>, which is
# also exactly the anchor the stored records use. So the id is the anchor to
# split on, and the leading '3.1' is stripped from the text afterwards —
# which works whether it arrived in a span or an h5.
_PARA_ID = re.compile(r'(?is)<(?:div|article|section|p)\b[^>]*\bid="ta(\d+)[-_](\d+)[-_](\d+)"[^>]*>')
_LEAD_NUMBER = re.compile(r"^\s*\d+\.\d+\s*")
_H2 = re.compile(r"(?is)<h2[^>]*>(.*?)</h2>")


def clean(fragment):
    fragment = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", fragment)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", fragment)).strip()


def chapter_headline(raw):
    """The committee-discussion h2, e.g. '3 Committee discussion'."""
    for h in _H2.findall(raw):
        text = clean(h)
        if "committee" in text.lower() and "discussion" in text.lower():
            return text
    for h in _H2.findall(raw):
        text = clean(h)
        if text.lower() != "on this page":
            return text
    return ""


def parse_paragraphs(raw, ta_hint=None):
    """One record per numbered paragraph, in the file's own shape."""
    headline = chapter_headline(raw)
    marks = list(_PARA_ID.finditer(raw))
    out = []
    for i, m in enumerate(marks):
        ta, major, minor = m.group(1), m.group(2), m.group(3)
        end = marks[i + 1].start() if i + 1 < len(marks) else len(raw)
        section = f"{major}.{minor}"
        # The number is shown inside the paragraph — in a span in one markup,
        # an h5 in the other. Stripping it after the tags are gone handles
        # both without caring which.
        body = _LEAD_NUMBER.sub("", clean(raw[m.end():end]))
        if len(body) < 40:
            continue
        out.append({
            "ta": str(int(ta_hint or ta)),
            "section": section,
            "headline": headline,
            "theme": None,
            "anchor": f"ta{int(ta_hint or ta)}-{section.replace('.', '_')}",
            "text": body,
        })
    return out


def find_html(ta):
    for d in CHAPTER_DIRS:
        if not os.path.isdir(d):
            continue
        for pattern in (f"*{int(ta):04d}*.html", f"*{int(ta)}*.html"):
            for path in sorted(glob.glob(os.path.join(d, pattern))):
                base = os.path.basename(path)
                if re.search(rf"(?<!\d){int(ta)}(?!\d)", base) or \
                   re.search(rf"(?<!\d){int(ta):04d}(?!\d)", base):
                    return path
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--index", default=INDEX)
    ap.add_argument("--sections", default=SECTIONS)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    index = load(args.index)
    sections = load(args.sections)
    if index is None or sections is None:
        sys.exit("chapter_index.json and chapter_sections.json must both be here.")
    if not isinstance(sections, list):
        sys.exit(f"Expected a list of paragraph records, found "
                 f"{type(sections).__name__}. Send me the first 200 characters.")

    template = next((r for r in sections if isinstance(r, dict)), None)
    if not template:
        sys.exit("chapter_sections.json has no dict records to copy the shape from.")
    fields = list(template)
    print(f"{args.sections}: {len(sections):,} paragraph records")
    print(f"record fields: {fields}")

    present = {(str(r.get("ta", "")).strip(), str(r.get("section", "")).strip())
               for r in sections if isinstance(r, dict)}
    have_ta = {str(r.get("ta", "")).strip() for r in sections if isinstance(r, dict)}
    print(f"appraisals represented: {len(have_ta):,}\n")

    grouped = buckets(index)
    rows, origin = {}, {}
    if grouped:
        for name, entries in grouped:
            for key, value in entries.items():
                rows[key], origin[key] = value, str(name).lower()
    else:
        rows = dict(records(index))

    todo = [k for k in rows
            if _BUCKETS.get(origin.get(k, ""), None) == "ok"
            and str(int(re.sub(r"\D", "", k) or 0)) not in have_ta]
    todo.sort(key=lambda k: -int(re.sub(r"\D", "", k) or 0))
    if not todo:
        sys.exit("Every done appraisal already has paragraphs. Nothing to do.")

    print(f"{len(todo)} appraisal(s) to parse\n")
    new, skipped = [], []
    for key in todo:
        n = re.sub(r"\D", "", str(key))
        path = find_html(n) if n else None
        if not path:
            skipped.append((n, "no local HTML"))
            continue
        raw = open(path, encoding="utf-8", errors="ignore").read()
        parsed = [r for r in parse_paragraphs(raw, ta_hint=n)
                  if (r["ta"], r["section"]) not in present]
        if not parsed:
            skipped.append((n, "no numbered paragraphs found"))
            continue
        # Keep only the fields this file actually uses, in its order.
        new.extend({f: r.get(f) for f in fields} for r in parsed)
        print(f"  TA{n:<6} {len(parsed):>3} paragraphs   "
              f"{parsed[0]['section']}–{parsed[-1]['section']}   "
              f"{parsed[0]['headline'][:34]}")

    print(f"\nparsed {len(new)} paragraphs from {len(todo) - len(skipped)} appraisals")
    for n, why in skipped:
        print(f"  skipped TA{n}: {why}")
    if not new:
        sys.exit("Nothing parsed; chapter_sections.json untouched.")

    merged = list(sections) + new
    if len(merged) < len(sections):
        sys.exit("Refusing to write: result is smaller than the original.")

    if args.dry_run:
        print(f"\n--dry-run: would grow {len(sections):,} -> {len(merged):,} records "
              f"({len(have_ta):,} -> {len(have_ta) + len(todo) - len(skipped):,} "
              "appraisals). Nothing written.")
        print("\nsample new record:")
        print(json.dumps(new[0], ensure_ascii=False)[:420])
        return

    backup = args.sections + ".backup2"
    shutil.copy2(args.sections, backup)
    tmp = args.sections + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(merged, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, args.sections)
    print(f"\nWrote {args.sections}: {len(sections):,} -> {len(merged):,} records")
    print(f"backup: {backup}")
    print("\nRe-run why_no_chapters.py to confirm, then restart Streamlit.")


if __name__ == "__main__":
    main()
