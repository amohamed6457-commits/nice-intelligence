import argparse
import json
import os
import re
import sys
from collections import Counter

CACHE = "iqwig_pdfs"
OUT = "reasons.json"

SECTION_CUES = re.compile(
    r"Fazit|Zusammenfassung\s+der\s+Bewertung|Ergebnis\s+der\s+Bewertung|"
    r"Kommentar\s+zum|Bewertung\s+der\s+Evidenz|Schlussfolgerung|"
    r"conclusion|summary\s+of\s+the\s+assessment", re.IGNORECASE)

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

EVIDENCE_TERMS = re.compile(
    r"Studie|Daten|Evidenz|Vergleich|Endpunkt|Nachweis|Recherche|Analyse|"
    r"Bewertung|Operationalisierung|Studienpool|Population|RCT|"
    r"Auswertung|Ergebnis|Effekt|\bpU\b|Meta-?Analyse|Verzerrung|"
    r"stud(?:y|ies)|data|evidence|comparison|endpoint|analysis|population",
    re.IGNORECASE)

NOISE = re.compile(
    r"Institut\s+f[uü]r\s+Qualit[aä]t|Dossierbewertung\s+[A-Z]\d|"
    r"Version\s+\d|Seite\s+\d+\s+von|^\s*\d+\s*$|"
    r"Gemeinsamer\s+Bundesausschuss|"
    r"in\s+[UÜ]bereinstimmung\s+mit|stimmt\s+.{0,20}\s+zu|"
    r"\bist\s+ein\s+(?:selektiver|humaner|monoklonaler|rekombinanter)|"
    r"wird\s+als\s+\w+\s+(?:f[uü]r|bei)\s|"
    r"angewendet\s+(?:zur|bei)\s", re.IGNORECASE)


MODULE3 = re.compile(
    r"Jahrestherapiekosten|Kosten\s+(?:der|f(?:ue|[uü])r|pro|je)\s|"
    r"\bcosts?\s+(?:for|of|per)\s|"
    r"Apothekenverkaufspreis|Zuzahlung|Rabatt|Packungsgr|"
    r"Verbrauch\s+(?:pro|je)|Wirkst[aä]rke|"
    r"Anzahl\s+der\s+(?:GKV-)?Patient|Zielpopulation|"
    r"Patientenzahl|Patientenpopulation\s+(?:in|der)\s|"
    r"Pr(?:ae|[aä])valenz|Inzidenz|Versorgungsanteil|"
    r"annual\s+treatment\s+costs?|number\s+of\s+(?:GKV\s+)?patients|"
    r"target\s+population\s+size|prevalence|incidence", re.IGNORECASE)

EVIDENCE_CRITIQUE = re.compile(
    r"Zusatznutzen|zweckm[aä]ssige|zweckm[aä]ßige|Vergleichstherapie|"
    r"Evidenzlage|Studienlage|Evidenznachweis|"
    r"Verzerrungspotenzial|Endpunkt|einarmig|indirekt|RCT|"
    r"Studienpool|Operationalisierung\s+des\s+Endpunkt|"
    r"added\s+benefit|comparator\s+therapy|risk\s+of\s+bias|"
    r"single-?arm|indirect\s+comparison", re.IGNORECASE)


def is_module3(sentence):
    return bool(MODULE3.search(sentence)
                and not EVIDENCE_CRITIQUE.search(sentence))


def sentences(text):
    flat = " ".join(str(text or "").split())
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+(?=[A-ZÄÖÜ])", flat)]


def harvest(path, max_pages=None, window=1, keep_module3=False):
    import pdfplumber
    found, dropped = [], []
    with pdfplumber.open(path) as pdf:
        pages = pdf.pages[:max_pages] if max_pages else pdf.pages
        for number, page in enumerate(pages, 1):
            try:
                text = page.extract_text() or ""
            except Exception:
                continue
            near_conclusion = bool(SECTION_CUES.search(text))
            items = sentences(text)
            for idx, sentence in enumerate(items):
                if not (40 <= len(sentence) <= 400):
                    continue
                if NOISE.search(sentence):
                    continue
                if not (REASON_CUES.search(sentence)
                        and EVIDENCE_TERMS.search(sentence)):
                    continue
                if is_module3(sentence) and not keep_module3:
                    dropped.append((number, sentence))
                else:
                    before = " ".join(items[max(0, idx - window):idx])
                    after = " ".join(items[idx + 1:idx + 1 + window])
                    found.append((number, sentence, near_conclusion,
                                  before[-600:], after[:600]))
    return found, dropped


def shape(sentence):
    text = sentence
    text = re.sub(r"\b[A-ZÄÖÜ][\w-]{3,}\b", "·", text)
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
    ap.add_argument("--window", type=int, default=1)
    ap.add_argument("--keep-module3", action="store_true")
    ap.add_argument("--show-dropped", action="store_true")
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()

    try:
        import pdfplumber
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
        got, cut = harvest(os.path.join(args.cache, files[0]),
                           args.max_pages, args.window, args.keep_module3)
        print(f"{len(got)} reason sentence(s)\n")
        for number, sentence, near, before, after in got[:40]:
            mark = "★" if near else " "
            print(f" {mark}[p{number}] {sentence[:280]}")
            if before:
                print(f"      before: ...{before[-150:]}")
            if after:
                print(f"      after : {after[:150]}...")
        print("\n★ = on a page that also carries a conclusion heading")
        if cut:
            print(f"\n{len(cut)} dropped as Module 3 (costs, patient "
                  f"numbers):")
            for number, sentence in cut[:20]:
                print(f"   [p{number}] {sentence[:190]}")
        return

    import time
    started = time.time()
    shapes, examples, per_doc, sentences_by_doc = Counter(), {}, {}, {}
    module3, dropped_examples = 0, []
    for i, name in enumerate(files, 1):
        try:
            got, cut = harvest(os.path.join(args.cache, name),
                               args.max_pages, args.window,
                               args.keep_module3)
            module3 += len(cut)
            if args.show_dropped:
                for number, sentence in cut:
                    dropped_examples.append((name, number, sentence))
        except Exception as exc:
            print(f"  {name}: FAILED {type(exc).__name__}: {exc}")
            continue
        per_doc[name] = len(got)
        sentences_by_doc[name] = [(n, s, b, a) for n, s, _x, b, a in got]
        for _number, sentence, _near, _b, _a in got:
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
    print(f"distinct phrasings  : {len(shapes):,}")
    print(f"dropped as Module 3 : {module3:,}"
          f"{' (kept, --keep-module3)' if args.keep_module3 else ''}\n")
    if args.show_dropped and dropped_examples:
        print("=" * 68)
        print("DROPPED AS MODULE 3 -- read these before trusting the filter")
        print("=" * 68)
        for name, number, sentence in dropped_examples[:40]:
            aid = name.split(".")[0].replace("_", "-")
            print(f"\n  {aid} p{number}")
            print(f"    {sentence[:200]}")
        print()

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

    out = {}
    for name, pairs in sentences_by_doc.items():
        aid = name.split(".")[0].replace("_", "-")
        out[aid] = [{"page": page, "sentence": sentence,
                     "before": before, "after": after}
                    for page, sentence, before, after in pairs]
    if args.limit and args.out == OUT:
        print(f"NOT WRITTEN: --limit was set and --out is the default "
              f"{OUT}.\n   This was a {len(files)}-document inspection run; "
              f"writing it would\n   replace the full corpus with a "
              f"fragment. Re-run without --limit,\n   or pass --out "
              f"something_else.json to keep this sample.")
        return
    with open(args.out, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)
    carried = sum(1 for v in out.values() for e in v
                  if e["before"] or e["after"])
    total = sum(len(v) for v in out.values())
    print(f"-> {args.out}  ({total:,} sentences from {len(out):,} appraisals)")
    print(f"   {carried:,} of {total:,} carry neighbouring context "
          f"(window {args.window})")
    print("\nRun with --vocabulary to see the phrasings before classifying.")


if __name__ == "__main__":
    main()
