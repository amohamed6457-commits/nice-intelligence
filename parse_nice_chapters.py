import argparse
import glob
import json
import os
import re

from bs4 import BeautifulSoup

# ta710_3_1 | ta993-3_1 | ta805_3_19
SECTION_ID = re.compile(r"^ta\d+[-_](\d+)_(\d+)$", re.IGNORECASE)


def log(m):
    print(m, flush=True)


def _nearest_section_titles(el):
    """Walk up for the committee conclusion headline and the broader theme."""
    headline = theme = None
    node = el.parent
    while node is not None and getattr(node, "get", None):
        classes = node.get("class") or []
        if "section" in classes or "chapter" in classes:
            title = (node.get("title") or "").strip()
            if title:
                if headline is None:
                    headline = title
                elif theme is None:
                    theme = title
                    break
        node = node.parent
    return headline, theme


def _clean_text(el, section_number):
    """Full subtree text, with the leading section number removed."""
    text = el.get_text(" ", strip=True)
    text = re.sub(r"\s+", " ", text)
    # The number is repeated at the start ("3.1 Atypical haemolytic ...").
    text = re.sub(rf"^{re.escape(section_number)}\s*", "", text)
    return text.strip()


def parse_chapter(html, source_name=""):
    """Return (sections, diagnostics)."""
    soup = BeautifulSoup(html, "lxml")

    root = soup.find("div", class_="chapter") or soup.find("main") or soup
    chapter_title = (root.get("title") if hasattr(root, "get") else None) or ""

    ta_match = re.search(r"\bTA(\d+)\b", str(soup.title.string if soup.title else "")) \
        or re.search(r"ta(\d+)[-_]", str(root.get("id") or "")) if hasattr(root, "get") else None
    ta = ta_match.group(1) if ta_match else None

    sections = []
    seen = set()

    for el in root.find_all(id=SECTION_ID):
        m = SECTION_ID.match(el["id"])
        chap_no, sec_no = m.group(1), m.group(2)
        number = f"{chap_no}.{sec_no}"
        if number in seen:
            continue
        seen.add(number)

        headline, theme = _nearest_section_titles(el)
        text = _clean_text(el, number)

        sections.append({
            "ta": ta,
            "section": number,
            "sort_key": (int(chap_no), int(sec_no)),
            "headline": headline,
            "theme": theme,
            "anchor": el["id"],
            "text": text,
            "chars": len(text),
        })

    sections.sort(key=lambda s: s["sort_key"])
    for s in sections:
        del s["sort_key"]

    diagnostics = {
        "source": source_name,
        "chapter_title": chapter_title,
        "n_sections": len(sections),
        "strategy": "element-id",
        "n_numbered_paragraphs": len(soup.find_all(class_="numbered-paragraph")),
        "missing_headline": sum(1 for s in sections if not s["headline"]),
        "empty_text": sum(1 for s in sections if len(s["text"]) < 40),
    }
    return sections, diagnostics


def verify(sections, diag):
    """Sanity checks that would catch an unrecognised markup variant."""
    problems = []
    n = diag["n_sections"]
    expected = diag["n_numbered_paragraphs"]

    if n == 0:
        problems.append("no sections extracted - unrecognised markup")
    elif expected and n != expected:
        problems.append(f"found {n} sections but {expected} numbered paragraphs")

    if diag["empty_text"]:
        problems.append(f"{diag['empty_text']} section(s) with almost no text")
    if diag["missing_headline"]:
        problems.append(f"{diag['missing_headline']} section(s) with no headline")

    # Section numbers should be contiguous from .1
    nums = [int(s["section"].split(".")[1]) for s in sections]
    if nums and nums != list(range(1, len(nums) + 1)):
        gaps = sorted(set(range(1, max(nums) + 1)) - set(nums))
        if gaps:
            problems.append(f"gaps in section numbering: {gaps[:8]}")

    return problems


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default="nice_chapters")
    ap.add_argument("--out", default="chapter_sections.json")
    ap.add_argument("--verify", action="store_true", help="print per-file checks and samples")
    args = ap.parse_args()

    files = sorted(glob.glob(os.path.join(args.dir, "*.html")))
    if not files:
        log(f"No .html files in {args.dir}/")
        return

    all_sections = []
    all_clean = True

    for path in files:
        html = open(path, encoding="utf-8").read()
        sections, diag = parse_chapter(html, os.path.basename(path))
        problems = verify(sections, diag)
        all_clean &= not problems

        log(f"\n{os.path.basename(path)}")
        log(f"  chapter   : {diag['chapter_title'][:60]}")
        log(f"  sections  : {diag['n_sections']}  "
            f"(numbered paragraphs in page: {diag['n_numbered_paragraphs']})")
        if sections:
            lens = [s["chars"] for s in sections]
            log(f"  text chars: min {min(lens)}  median {sorted(lens)[len(lens)//2]}  max {max(lens)}")
            log(f"  range     : {sections[0]['section']} .. {sections[-1]['section']}")
        if problems:
            for p in problems:
                log(f"  PROBLEM   : {p}")
        else:
            log("  checks    : ok")

        if args.verify and sections:
            s = sections[0]
            log(f"  -- first section --")
            log(f"     {s['section']}  theme={s['theme']!r}")
            log(f"     headline: {s['headline']}")
            log(f"     text    : {s['text'][:170]}...")
            long = max(sections, key=lambda x: x["chars"])
            log(f"  -- longest ({long['section']}, {long['chars']} chars) --")
            log(f"     {long['text'][:170]}...")

        all_sections.extend(sections)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(all_sections, f, ensure_ascii=False, indent=1)

    log(f"\n{'='*52}")
    log(f"files parsed   : {len(files)}")
    log(f"sections total : {len(all_sections)}")
    log(f"written to     : {args.out}")
    log(f"status         : {'all checks passed' if all_clean else 'PROBLEMS - see above'}")


if __name__ == "__main__":
    main()
