"""
Pilot: why was the evidence unassessable? Read it from the Tragende Gründe.

The question
------------
468 G-BA appraisals have no assessable evidence in ANY endpoint category.
The Beschluss says only "Es liegen keine bewertbaren Daten vor" -- it states
the conclusion and not the cause. The Tragende Gründe zum Beschluss, which
99% of those 468 have, explains WHY.

Why the cause matters more than the count
-----------------------------------------
"77% of appraisals had no usable quality-of-life evidence" is a statistic.
"Of the 468 appraisals with nothing assessable, N failed because the
comparator arm was not the one G-BA specified, and M because the study was
single-arm" is a finding someone can act on -- because those two causes have
completely different implications for whether real-world evidence could fix
them. A wrong comparator arm cannot be rescued by registry data; an absent
long-term outcome often can.

What this pilot does
--------------------
Fetches the Tragende Gründe for a few of those appraisals, finds the
sentences that give a reason, and prints them with their page. It does NOT
classify yet: the point is to look at real text first and find out what the
reasons actually say, rather than inventing categories and then matching
them. That mistake has already cost us one benchmark this week.

    python pilot_rejection_reasons.py                 # 5 from the gap list
    python pilot_rejection_reasons.py --n 20
    python pilot_rejection_reasons.py --appraisal D-1155
"""

import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from endpoint_summary import CATEGORIES, find_workbook, verdict
except ImportError as exc:
    sys.exit(f"endpoint_summary.py must sit beside this script ({exc}).")

CACHE = "gba_reasoning_pdfs"
DOC_TYPE = "Tragende Gründe"

# Phrases that introduce a reason rather than a conclusion. Kept deliberately
# broad: this is a look, not a classifier.
REASON_CUES = [
    r"entspricht\s+nicht\s+der\s+zweckm[aä][sß]igen\s+Vergleichstherapie",
    r"nicht\s+der\s+zweckm[aä][sß]igen\s+Vergleichstherapie",
    r"einarmig",
    r"keine\s+(?:direkt\s+)?vergleichenden\s+Studien",
    r"kein\s+adjustierter\s+indirekter\s+Vergleich",
    r"nicht\s+bewertbar",
    r"keine\s+(?:geeigneten|verwertbaren|bewertbaren)\s+Daten",
    r"wurden\s+nicht\s+vorgelegt",
    r"nicht\s+f[uü]r\s+die\s+Nutzenbewertung\s+herangezogen",
    r"entspricht\s+nicht\s+dem\s+(?:zugelassenen\s+)?Anwendungsgebiet",
    r"hohe[rn]?\s+Anteil\s+(?:an\s+)?Studienabbr",
    r"nicht\s+validiert",
    r"Operationalisierung",
    r"Verzerrungspotenzial",
    # English courtesy translations
    r"does\s+not\s+correspond\s+to\s+the\s+appropriate\s+comparator",
    r"single-?arm",
    r"no\s+suitable\s+data",
    r"were\s+not\s+submitted",
    r"not\s+used\s+for\s+the\s+benefit\s+assessment",
    r"high\s+(?:proportion|rate)\s+of\s+(?:study\s+)?discontinuation",
    r"not\s+validated",
    r"risk\s+of\s+bias",
]
REASON_CUES = [re.compile(p, re.IGNORECASE) for p in REASON_CUES]


def gap_appraisals(workbook):
    """Appraisals where every endpoint category was unassessable."""
    import pandas as pd
    subs = pd.read_excel(workbook, sheet_name="GBA_Subgroups")
    subs["aid"] = subs["appraisal_id"].astype(str).str.strip()
    out = []
    for aid, group in subs.groupby("aid"):
        verdicts = [verdict(getattr(row, field))
                    for row in group.itertuples() for field, _ in CATEGORIES]
        verdicts = [v for v in verdicts if v]
        if verdicts and all(v == "NO DATA" for v in verdicts):
            out.append(aid)
    return sorted(out)


def documents(workbook):
    import pandas as pd
    docs = pd.read_excel(workbook, sheet_name="Appraisal_Documents")
    docs["aid"] = docs["Appraisal_ID"].astype(str).str.strip()
    sel = docs[docs["Document_Type"].astype(str)
               .str.contains(DOC_TYPE, case=False, na=False)]
    sel = sel[sel["Source_URL"].fillna("").astype(str).str.startswith("http")]
    return {r.aid: str(r.Source_URL) for r in sel.itertuples()}


def fetch(url, name):
    os.makedirs(CACHE, exist_ok=True)
    path = os.path.join(CACHE, re.sub(r"[^A-Za-z0-9_.-]", "_", f"{name}.pdf"))
    if os.path.exists(path) and os.path.getsize(path) > 1000:
        return path
    try:
        import requests
    except ImportError:
        sys.exit("requests is not installed: pip install requests")
    print(f"  downloading {url[:100]}")
    resp = requests.get(url, timeout=180,
                        headers={"User-Agent": "hta-intelligence-pilot"})
    resp.raise_for_status()
    if not resp.content.startswith(b"%PDF"):
        print(f"  NOT A PDF ({len(resp.content):,} bytes)")
        return None
    with open(path, "wb") as fh:
        fh.write(resp.content)
    print(f"  saved {len(resp.content):,} bytes")
    return path


def reasons_in(path, limit=6):
    """Sentences that give a reason, with their page."""
    import pdfplumber
    found = []
    with pdfplumber.open(path) as pdf:
        pages = len(pdf.pages)
        for number, page in enumerate(pdf.pages, 1):
            try:
                text = page.extract_text() or ""
            except Exception:                              # noqa: BLE001
                continue
            flat = " ".join(text.split())
            for cue in REASON_CUES:
                for match in cue.finditer(flat):
                    start = flat.rfind(".", 0, match.start()) + 1
                    end = flat.find(".", match.end())
                    sentence = flat[start:end + 1 if end > 0 else len(flat)]
                    sentence = sentence.strip()
                    if 40 <= len(sentence) <= 420:
                        found.append((number, sentence))
                    break
            if len(found) >= limit:
                break
    return pages, found


def outline(path, chars=260):
    """What is actually in this document: its headings, page by page.

    Needed because three of the first five Tragende Gründe matched no reason
    cue at all, and one was only three pages. Before widening the cue list
    -- which would be guessing twice -- look at what these documents
    contain.
    """
    import pdfplumber
    out = []
    heading = re.compile(r"^\s*(\d+(?:\.\d+)*)\s+(\S.{3,70})$")
    with pdfplumber.open(path) as pdf:
        for number, page in enumerate(pdf.pages, 1):
            try:
                text = page.extract_text() or ""
            except Exception:                              # noqa: BLE001
                continue
            heads = [m.group(0).strip() for line in text.splitlines()
                     for m in [heading.match(line)] if m]
            first = " ".join(text.split())[:chars]
            out.append((number, heads, first))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook")
    ap.add_argument("--appraisal", action="append")
    ap.add_argument("--n", type=int, default=5)
    ap.add_argument("--seed", type=int, default=20261112)
    ap.add_argument("--show", action="store_true",
                    help="print each document's headings and opening text "
                         "instead of searching for reason sentences")
    ap.add_argument("--type", default=DOC_TYPE,
                    help="Document_Type to pull "
                         "(e.g. 'Zusammenfassende Dokumentation', "
                         "'Nutzenbewertung IQWiG')")
    args = ap.parse_args()
    globals()["DOC_TYPE"] = args.type

    try:
        import pdfplumber                                  # noqa: F401
    except ImportError:
        sys.exit("pdfplumber is not installed: pip install pdfplumber")

    wb = find_workbook(args.workbook)
    print(f"workbook: {wb}")
    links = documents(wb)
    if args.appraisal:
        chosen = [a.strip() for a in args.appraisal]
    else:
        gap = [a for a in gap_appraisals(wb) if a in links]
        print(f"appraisals with nothing assessable AND a Tragende Gründe: "
              f"{len(gap):,}")
        import random
        random.seed(args.seed)
        chosen = random.sample(gap, min(args.n, len(gap)))
        print(f"random sample (seed {args.seed}): {', '.join(chosen)}")
    print()

    for aid in chosen:
        url = links.get(aid)
        print(f"=== {aid} ===")
        if not url:
            print("  no Tragende Gründe linked\n")
            continue
        try:
            path = fetch(url, aid)
        except Exception as exc:                           # noqa: BLE001
            print(f"  FAILED {type(exc).__name__}: {exc}\n")
            continue
        if not path:
            print()
            continue
        if args.show:
            pages = outline(path)
            print(f"  {len(pages)} pages\n")
            for number, heads, first in pages:
                print(f"  --- page {number} ---")
                for head in heads[:6]:
                    print(f"      § {head[:78]}")
                print(f"      {first[:240]}")
            print(f"  {url[:110]}\n")
            continue
        pages, found = reasons_in(path)
        print(f"  {pages} pages · {len(found)} reason sentence(s)")
        for number, sentence in found:
            print(f"    [p{number}] {sentence[:300]}")
        print(f"  {url[:110]}\n")

    print("=" * 62)
    print("Read these before anyone classifies them. The categories should "
          "come\nfrom what the documents actually say, not from a list "
          "written first.")


if __name__ == "__main__":
    main()
