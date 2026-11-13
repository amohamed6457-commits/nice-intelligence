"""
Read the downloaded IQWiG assessments and surface WHY evidence was rejected.

This does not classify. That is the point.

The table classifier reached 75% precision because the categories were
written first and the text was matched to them afterwards. The same mistake
here would be worse, because a wrong reason attached to an appraisal is a
claim about why a regulator rejected evidence -- the kind of thing someone
repeats in a meeting.

So this runs in two passes:

    --vocabulary   read every document, pull the sentences that state a
                   reason, and report how they are PHRASED, ranked by
                   frequency. Categories are then written from what the
                   documents say.
    (default)      once the categories exist, extract per appraisal and
                   write reasons.json with the sentence and its page.

Run --vocabulary first and read the output. It is the only way to find out
what reasons G-BA and IQWiG actually give, as opposed to the ones that seem
plausible from the outside.

    python extract_rejection_reasons.py --vocabulary
    python extract_rejection_reasons.py --vocabulary --limit 50
    python extract_rejection_reasons.py --appraisal D-1177
"""

import argparse
import json
import os
import re
import sys
from collections import Counter

CACHE = "iqwig_pdfs"
OUT = "reasons.json"

# Where an assessment states its conclusions. Used to narrow the search to
# the parts of a 200-page document that carry a verdict, rather than reading
# every page of study description.
SECTION_CUES = re.compile(
    r"Fazit|Zusammenfassung\s+der\s+Bewertung|Ergebnis\s+der\s+Bewertung|"
    r"Kommentar\s+zum|Bewertung\s+der\s+Evidenz|Schlussfolgerung|"
    r"conclusion|summary\s+of\s+the\s+assessment", re.IGNORECASE)

# Sentences that state WHY something could not be used. Deliberately broad
# and overlapping: this pass is meant to over-collect so the vocabulary is
# visible, not to decide anything.
REASON_CUES = re.compile(
    r"nicht\s+geeignet|ungeeignet|nicht\s+bewertbar|keine\s+Aussage|"
    r"liegen\s+keine|nicht\s+vorgelegt|nicht\s+herangezogen|"
    r"entspricht\s+nicht|nicht\s+umgesetzt|einarmig|"
    r"keine\s+vergleichenden|hohes?\s+Verzerrungspotenzial|"
    r"nicht\s+validiert|nicht\s+ausreichend|unzureichend|"
    r"Operationalisierung|nicht\s+interpretierbar|"
    r"no\s+suitable|not\s+assessable|single-?arm|not\s+submitted|"
    r"does\s+not\s+correspond|high\s+risk\s+of\s+bias|not\s+validated",
    re.IGNORECASE)

# A reason has to be ABOUT the evidence. Without this, any sentence with a
# negation near an assessment word qualifies, and the 20-document pass
# returned things like "Rapiscan® ist ein selektiver koronarer
# Vasodilatator" -- a drug description -- as a reason for rejection.
EVIDENCE_TERMS = re.compile(
    r"Studie|Daten|Evidenz|Vergleich|Endpunkt|Nachweis|Recherche|Analyse|"
    r"Bewertung|Operationalisierung|Studienpool|Population|RCT|"
    r"Auswertung|Ergebnis|Effekt|\bpU\b|Meta-?Analyse|Verzerrung|"
    r"stud(?:y|ies)|data|evidence|comparison|endpoint|analysis|population",
    re.IGNORECASE)

# Noise that is never a reason: page furniture, boilerplate, and -- the two
# that cost the first pass -- statements of AGREEMENT with the G-BA, and
# definitions of what a drug is.
NOISE = re.compile(
    r"Institut\s+f[uü]r\s+Qualit[aä]t|Dossierbewertung\s+[A-Z]\d|"
    r"Version\s+\d|Seite\s+\d+\s+von|^\s*\d+\s*$|"
    r"Gemeinsamer\s+Bundesausschuss|"
    r"in\s+[UÜ]bereinstimmung\s+mit|stimmt\s+.{0,20}\s+zu|"
    r"\bist\s+ein\s+(?:selektiver|humaner|monoklonaler|rekombinanter)|"
    r"wird\s+als\s+\w+\s+(?:f[uü]r|bei)\s|"
    r"angewendet\s+(?:zur|bei)\s", re.IGNORECASE)


def sentences(text):
    flat = " ".join(str(text or "").split())
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+(?=[A-ZÄÖÜ])", flat)]


def harvest(path, max_pages=None):
    """(page, sentence) for every sentence that states a reason."""
    import pdfplumber
    found = []
    with pdfplumber.open(path) as pdf:
        pages = pdf.pages[:max_pages] if max_pages else pdf.pages
        for number, page in enumerate(pages, 1):
            try:
                text = page.extract_text() or ""
            except Exception:                              # noqa: BLE001
                continue
            near_conclusion = bool(SECTION_CUES.search(text))
            for sentence in sentences(text):
                if not (40 <= len(sentence) <= 400):
                    continue
                if NOISE.search(sentence):
                    continue
                # Both required: something inadequate, AND it is the
                # evidence that is inadequate.
                if (REASON_CUES.search(sentence)
                        and EVIDENCE_TERMS.search(sentence)):
                    found.append((number, sentence, near_conclusion))
    return found


def shape(sentence):
    """A sentence reduced to its shape, so phrasings can be counted.

    Drug names, study names, numbers and endpoints vary in every document;
    the REASON does not. Stripping the specifics makes the recurring
    phrasings visible instead of 400 unique sentences.
    """
    text = sentence
    text = re.sub(r"\b[A-ZÄÖÜ][\w-]{3,}\b", "·", text)     # proper nouns
    text = re.sub(r"\d+[\d.,%]*", "#", text)
    text = re.sub(r"[·\s]{2,}", " · ", text)
    return " ".join(text.split())[:150]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cache", default=CACHE)
    ap.add_argument("--vocabulary", action="store_true")
    ap.add_argument("--appraisal")
    ap.add_argument("--limit", type=int)
    ap.add_argument("--max-pages", type=int, default=60,
                    help="pages to read per document (default 60; the "
                         "conclusions are near the front and the back "
                         "matter is study tables)")
    ap.add_argument("--top", type=int, default=40)
    args = ap.parse_args()

    try:
        import pdfplumber                                  # noqa: F401
    except ImportError:
        sys.exit("pdfplumber is not installed: pip install pdfplumber")
    if not os.path.isdir(args.cache):
        sys.exit(f"{args.cache}/ not found — run fetch_iqwig.py first.")

    files = sorted(f for f in os.listdir(args.cache) if f.endswith(".pdf"))
    if args.appraisal:
        want = args.appraisal.strip().replace("-", "_")
        files = [f for f in files if f.split(".")[0].replace("-", "_") == want]
        if not files:
            sys.exit(f"No PDF for {args.appraisal} in {args.cache}/.")
    if args.limit:
        files = files[:args.limit]
    print(f"{len(files):,} document(s) in {args.cache}/\n")

    if args.appraisal:
        got = harvest(os.path.join(args.cache, files[0]), args.max_pages)
        print(f"{len(got)} reason sentence(s)\n")
        for number, sentence, near in got[:40]:
            mark = "★" if near else " "
            print(f" {mark}[p{number}] {sentence[:280]}")
        print("\n★ = on a page that also carries a conclusion heading")
        return

    import time
    started = time.time()
    shapes, examples, per_doc, sentences_by_doc = Counter(), {}, {}, {}
    for i, name in enumerate(files, 1):
        try:
            got = harvest(os.path.join(args.cache, name), args.max_pages)
        except Exception as exc:                           # noqa: BLE001
            print(f"  {name}: FAILED {type(exc).__name__}: {exc}")
            continue
        per_doc[name] = len(got)
        sentences_by_doc[name] = [(n, s) for n, s, _ in got]
        for _number, sentence, _near in got:
            key = shape(sentence)
            shapes[key] += 1
            examples.setdefault(key, sentence)
        if i % 10 == 0 or i == len(files):
            rate = (time.time() - started) / i
            left = (len(files) - i) * rate / 60
            print(f"  read {i}/{len(files)}  ({rate:.1f}s each, "
                  f"~{left:.0f} min left)")

    print(f"\ndocuments read      : {len(per_doc):,}")
    print(f"with any reason     : {sum(1 for v in per_doc.values() if v):,}")
    print(f"reason sentences    : {sum(per_doc.values()):,}")
    print(f"distinct phrasings  : {len(shapes):,}\n")

    if args.vocabulary:
        print("=" * 68)
        print("HOW THE REASONS ARE PHRASED — write the categories from these")
        print("=" * 68)
        for key, n in shapes.most_common(args.top):
            print(f"\n{n:>5}x  {key}")
            print(f"        e.g. {examples[key][:200]}")
        print("\n" + "=" * 68)
        print("Read these before writing any category list. The categories "
              "have to\ncome from what the documents say, not from what "
              "sounds plausible.")
        return

    # Keep the sentences themselves, with their page: this is the input the
    # classifier reads, and the quote a citation will point at.
    out = {}
    for name, pairs in sentences_by_doc.items():
        aid = name.split(".")[0].replace("_", "-")
        out[aid] = [{"page": page, "sentence": sentence}
                    for page, sentence in pairs]
    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)
    print(f"-> {OUT}  ({sum(len(v) for v in out.values()):,} sentences from "
          f"{len(out):,} appraisals)")
    print("\nRun with --vocabulary to see the phrasings before classifying.")


if __name__ == "__main__":
    main()
