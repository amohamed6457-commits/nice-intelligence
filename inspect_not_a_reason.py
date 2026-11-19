import argparse
import hashlib
import json
import os
import re
import sys
from collections import Counter, defaultdict

REASONS = "reasons.json"
LABELS = "reason_categories_v4.json"
TARGET = "NOT_A_REASON"

UNUSABLE = re.compile(
    r"nicht\s+geeignet|ungeeignet|nicht\s+bewertbar|nicht\s+verwertbar|"
    r"nicht\s+interpretierbar|nicht\s+aussagekr[aä]ftig|"
    r"keine\s+(?:geeigneten|verwertbaren|validen)\s|"
    r"nicht\s+herangezogen\s+werden|lassen\s+sich\s+keine\s+Aussagen|"
    r"nicht\s+ausreichend|unzureichend|nicht\s+plausibel|"
    r"nicht\s+nachvollziehbar|nicht\s+sachgerecht|nicht\s+relevant|"
    r"\birrelevant|nicht\s+valide|nicht\s+belastbar|"
    r"not\s+suitable|unsuitable|not\s+assessable|cannot\s+be\s+used|"
    r"not\s+interpretable|not\s+sufficient|insufficient",
    re.IGNORECASE)

SIGNPOST = re.compile(
    r"(?:nachfolgend|im\s+Folgenden|folgend)[^.]{0,80}"
    r"(?:beschrieben|begr[uü]ndet|erl[aä]utert|dargestellt|dargelegt)|"
    r"(?:beschrieben|begr[uü]ndet|erl[aä]utert)[^.]{0,40}"
    r"(?:nachfolgend|im\s+Folgenden)|"
    r"In\s+Abschnitt\s+[\d.]+\s+(?:werden|wird)|"
    r"Dies\s+wird\s+(?:nachfolgend|im\s+Folgenden)|"
    r"siehe\s+(?:hierzu\s+)?Abschnitt\s+[\d.]+\s*\)?\s*\.?\s*$|"
    r"wird\s+nachfolgend\s+(?:begr[uü]ndet|erl[aä]utert)",
    re.IGNORECASE)

VERDICT_SUBJECT = re.compile(
    r"(?:vom\s+pU\s+)?(?:vorgelegten?|eingereichten?|herangezogenen?)\s+"
    r"(?:Daten|Unterlagen|Nachweise|Analysen|Studien|Evidenz|"
    r"[UÜ]berlegungen|Auswertungen)|"
    r"(?:Daten|Unterlagen|Analysen|Nachweise)\s+des\s+pU|"
    r"Angaben\s+des\s+pU|Begr[uü]ndung\s+des\s+pU|"
    r"data\s+(?:submitted|presented)\s+by", re.IGNORECASE)

DEFERS = re.compile(
    r"aus\s+den\s+(?:in|oben|vorgenannten)[^.]{0,60}Gr[uü]nden|"
    r"wie\s+(?:oben|zuvor|bereits)\s+(?:beschrieben|dargelegt|ausgef)|"
    r"siehe\s+(?:Abschnitt|Kapitel|oben|unten)|"
    r"\bsomit\b|\bdaher\b|\bdemnach\b|\bdeshalb\b|\bfolglich\b|"
    r"aufgrund\s+der\s+(?:oben|vorgenannten|genannten)|"
    r"\btherefore\b|\bthus\b|\bconsequently\b|"
    r"for\s+the\s+reasons\s+(?:described|stated|given)|"
    r"as\s+(?:described|outlined|stated)\s+above", re.IGNORECASE)

HEADING = re.compile(r"^\s*(?:[IVX]+\s+)?\d+(?:\.\d+){0,3}\s+[A-ZÄÖÜ]")

RESULT = re.compile(
    r"statistisch\s+signifikant|zeigte\s+sich|ergab\s+sich|"
    r"Hazard\s*Ratio|\bHR\s*[=:]|\bRR\s*[=:]|\bOR\s*[=:]|"
    r"\b95\s*%\s*(?:-?\s*)?(?:KI|CI)|p\s*[<=>]\s*0[.,]|"
    r"statistically\s+significant", re.IGNORECASE)

AGREE = re.compile(
    r"in\s+[UÜ]bereinstimmung\s+mit|stimmt\s+.{0,25}\s+zu|"
    r"ist\s+nachvollziehbar|wird\s+zugestimmt|sachgerecht|"
    r"folgt\s+der\s+Einsch[aä]tzung", re.IGNORECASE)

STATUTE = re.compile(r"§\s*35a|Paragraph\s*35a|SGB\s*V", re.IGNORECASE)


def key_of(text):
    return hashlib.sha1(text.strip().encode("utf-8")).hexdigest()[:16]


def load(path):
    if not os.path.exists(path):
        sys.exit(f"{path} not found.")
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def shape_of(sentence):
    words = sentence.split()
    digits = sum(c.isdigit() for c in sentence)
    tags = []
    if HEADING.match(sentence):
        tags.append("starts with a section number")
    if words and digits / max(1, len(sentence)) > 0.12:
        tags.append("number-dense (likely a table row)")
    if UNUSABLE.search(sentence):
        tags.append("asserts evidence is unusable")
    if SIGNPOST.search(sentence):
        tags.append("signpost: announces reasons follow")
    elif VERDICT_SUBJECT.search(sentence) and UNUSABLE.search(sentence):
        tags.append("verdict on submitted evidence, no cause given")
    elif DEFERS.search(sentence):
        tags.append("defers its reason elsewhere")
    if RESULT.search(sentence):
        tags.append("reports a result")
    if AGREE.search(sentence):
        tags.append("states agreement")
    if STATUTE.search(sentence):
        tags.append("cites statute")
    return tags


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--reasons", default=REASONS)
    ap.add_argument("--labels", default=LABELS)
    ap.add_argument("--target", default=TARGET)
    ap.add_argument("--show", type=int, default=8)
    ap.add_argument("--only")
    args = ap.parse_args()

    reasons = load(args.reasons)
    labels = load(args.labels)

    seen = {}
    for aid, entries in reasons.items():
        for entry in entries:
            text = str(entry.get("sentence") or "").strip()
            if len(text) < 40:
                continue
            key = key_of(text)
            if labels.get(key) != args.target:
                continue
            seen.setdefault(key, (aid, entry.get("page"), text,
                                  str(entry.get("before") or ""),
                                  str(entry.get("after") or "")))
    if not seen:
        sys.exit(f"Nothing labelled {args.target} in {args.labels}.")

    buckets, counts = defaultdict(list), Counter()
    for key, row in seen.items():
        tags = shape_of(row[2])
        if not tags:
            tags = ["no mechanical signal"]
        for tag in tags:
            counts[tag] += 1
            buckets[tag].append(row)

    total = len(seen)
    print(f"{total:,} distinct sentences labelled {args.target}\n")
    print("Mechanical shape of those sentences (a sentence can carry more")
    print("than one signal, so these do not sum to the total):\n")
    for tag, n in counts.most_common():
        print(f"  {n:>5}  {n / total:>5.1%}  {tag}")

    signpost = {k: r for k, r in seen.items() if SIGNPOST.search(r[2])}
    verdict = {k: r for k, r in seen.items()
               if k not in signpost
               and VERDICT_SUBJECT.search(r[2]) and UNUSABLE.search(r[2])}
    other_unusable = {k: r for k, r in seen.items()
                      if k not in signpost and k not in verdict
                      and UNUSABLE.search(r[2])}
    print(f"\nSplit of the bucket:")
    print(f"  {len(signpost):>5}  {len(signpost) / total:>5.1%}  "
          f"signposts -- correctly {args.target}")
    print(f"  {len(verdict):>5}  {len(verdict) / total:>5.1%}  "
          f"verdict on submitted evidence with no cause")
    print(f"        -> these belong in REASON_NOT_STATED, not "
          f"{args.target}")
    print(f"  {len(other_unusable):>5}  {len(other_unusable) / total:>5.1%}  "
          f"other sentences asserting something is unusable")
    print(f"        -> needs reading; the phrase can describe patients or "
          f"sections,\n           not only evidence")
    wrong = dict(verdict)

    ctx_has_cause = sum(
        1 for r in wrong.values()
        if UNUSABLE.search(r[3] + " " + r[4])
        or re.search(r"\bda\b|\bweil\b|\baufgrund\b|\bsince\b|\bbecause\b",
                     r[3] + " " + r[4], re.IGNORECASE))
    if wrong:
        print(f"\nof those {len(wrong):,}, the context contains a causal "
              f"clause in {ctx_has_cause:,} "
              f"({ctx_has_cause / len(wrong):.0%})\n  -- the model had "
              f"something to work with and still said {args.target}")

    order = [args.only] if args.only else [t for t, _n in counts.most_common()]
    for tag in order:
        rows = buckets.get(tag, [])
        if not rows:
            print(f"\nno sentences tagged {tag!r}")
            continue
        print("\n" + "=" * 70)
        print(f"{tag.upper()}  ({len(rows):,})")
        print("=" * 70)
        for aid, page, text, before, after in rows[:args.show]:
            print(f"\n  {aid} p{page}")
            print(f"    {' '.join(text[:230].split())}")
            if before:
                print(f"    [before] ...{' '.join(before[-130:].split())}")
            if after:
                print(f"    [after ] {' '.join(after[:130].split())}...")

    print("\n" + "=" * 70)
    print("All of the above is pattern matching on the sentences, not a "
          "judgement\nabout whether any individual label is right. It says "
          "what shape the\nsentences in this bucket have.")


if __name__ == "__main__":
    main()
