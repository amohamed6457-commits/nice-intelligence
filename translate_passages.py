"""
Translate G-BA and HAS passages into English, cached to translations.json.

The tool's readers work in English; the sources are German and French. This
produces an English rendering ALONGSIDE the original — never instead of it.
That distinction is the whole design:

  * the original text stays the record, and the page link still opens the
    official PDF in its own language;
  * a translation is labelled as machine translation wherever it appears;
  * the cited-answer flow keeps quoting the ORIGINAL inside quotation marks,
    because a quotation that has been through a model is not a quotation.

Terminology is pinned by a glossary in the system prompt. These are terms of
art, not vocabulary: "Zusatznutzen ist nicht belegt" means a specific
regulatory finding, and a loose rendering ("no additional use was proven")
would misreport a decision. The glossary is the reason a small model is
viable here at all.

Cost, measured against the real corpus (19.4M characters):

    Haiku 5.5    ~$3.40       Sonnet 5.5    ~$67

so the choice is about quality, not budget. Start with --sample, which now
reports three failure modes rather than printing output for you to eyeball:
an empty or truncated answer, a number that was reformatted, and a passage
that came back untranslated.

    python translate_passages.py --sample 20      # eyeball before committing
    python translate_passages.py --limit 200
    python translate_passages.py                  # everything, resumable
    python translate_passages.py --model claude-sonnet-5-5

Needs ANTHROPIC_API_KEY in the environment or .streamlit/secrets.toml.
"""

import argparse
import concurrent.futures as futures
import hashlib
import json
import os
import random
import re
import sys
import time

OUT = "translations.json"
WORKBOOKS = ["NICE_HAS_GBA_updated_2026-11-07.xlsx",
             "NICE_v14_updated_2026-09-29.xlsx",
             "NICE_v14_updated_2026-09-28.xlsx"]

# Every column a reader actually sees on screen, not just the long passages.
# The record body -- indication, comparator, patient population, the reasons
# -- is what someone reads FIRST, and leaving it out meant the language
# switch appeared to do nothing at all.
SHEETS = [
    # long cited passages
    ("GBA_Decision_Passages", "source_text", "de"),
    ("HAS_Economic_Pages", "source_text", "fr"),
    ("Economic_Findings", "source_context", "fr"),
    # The G-BA record body. These live on G-BA_Recent, NOT on GBA_Subgroups:
    # same column names in places, but the G-BA_Recent values are the
    # subgroups joined together, so they hash differently and a translation
    # of the subgroup row does not match what the record view renders.
    ("G-BA_Recent", "indication", "de"),
    ("G-BA_Recent", "decision_raw", "de"),
    ("G-BA_Recent", "comparator_therapy", "de"),
    ("G-BA_Recent", "patient_population", "de"),
    ("G-BA_Recent", "decision_reasoning", "de"),
    ("G-BA_Recent", "added_benefit_assessment", "de"),
    ("GBA_Subgroups", "patient_subgroup", "de"),
    ("GBA_Subgroups", "comparator_therapy", "de"),
    ("GBA_Subgroups", "benefit_conclusion", "de"),
    ("GBA_Subgroups", "committee_reasoning", "de"),
    ("GBA_Subgroups", "appraised_indication", "de"),
    ("GBA_Subgroups", "licensed_indication", "de"),
    ("GBA_Subgroups", "mortality", "de"),
    ("GBA_Subgroups", "morbidity", "de"),
    ("GBA_Subgroups", "quality_of_life", "de"),
    ("GBA_Subgroups", "adverse_events", "de"),
    # the HAS record body
    ("HAS_Recent", "indication", "fr"),
    ("HAS_Recent", "decision_raw", "fr"),
    ("HAS_Recent", "comparator", "fr"),
    ("HAS_Recent", "patient_population", "fr"),
    ("HAS_Recent", "decision_reasoning", "fr"),
    ("HAS_Ratings", "reasoning", "fr"),
]

GLOSSARY_DE = """Zusatznutzen = added benefit (the specific thing a G-BA assessment decides)
Ein Zusatznutzen ist nicht belegt = an added benefit is not proven
zweckmäßige Vergleichstherapie = appropriate comparator therapy
Anhaltspunkt / Hinweis / Beleg = indication / hint / proof (ascending evidence certainty; keep the distinction)
Beschluss = resolution
Jahrestherapiekosten = annual treatment costs
Gesamtüberleben = overall survival
Morbidität / Mortalität = morbidity / mortality
unerwünschte Ereignisse (UE) = adverse events (AE)
schwerwiegende UE (SUE) = serious adverse events (SAE)
keine verwertbaren Daten = no usable data
k. A. = not reported
nicht belegt = not proven"""

GLOSSARY_FR = """service médical rendu (SMR) = actual clinical benefit (sets the reimbursement rate)
amélioration du service médical rendu (ASMR) = improvement in clinical benefit (drives price)
avis d'efficience = efficiency opinion
Commission de la Transparence = Transparency Committee
rapport différentiel coût-résultat (RDCR) = incremental cost-effectiveness ratio (ICER)
année de vie gagnée (AVG) = life year gained
données de vie réelle = real-world data
critère de jugement = endpoint
survie globale (SG) = overall survival
effets indésirables = adverse events
non estimable / n.d. = not estimable"""

SYSTEM = """You translate health technology assessment documents into English for
health economists who will act on the result.

RULES
1. Translate accurately and literally. This is a regulatory record, not prose
   to improve. Do not summarise, shorten, reorder or explain.
2. Use the glossary below for every term it covers. These are terms of art
   with specific regulatory meaning; a natural-sounding alternative is wrong.
3. COPY numbers, never convert them. Every number, unit, confidence interval,
   p-value, percentage, patient count and date must appear in the translation
   character for character as it appears in the source: the same decimal
   comma, the same spaces used as thousands separators, the same brackets,
   and the currency symbol in the same position.
       83,5%              stays  83,5%              NOT 83.5%
       846 933 €/QALY     stays  846 933 €/QALY     NOT €846,933/QALY
       9 931 €            stays  9 931 €            NOT €9,931
       0,79 [0,62; 1,01]  stays  0,79 [0,62; 1,01]
   Reformatting a number is an error even when its value is unchanged. The
   reader checks these against the original PDF side by side.
4. Keep drug names, trial names, study acronyms and document references in
   their original form.
5. Preserve the line structure of the input. One line in, one line out.
6. If a passage is already in English, return it unchanged.
7. Output ONLY the translation. No preamble, no notes, no quotation marks
   around it.

GLOSSARY
{glossary}"""


def key_of(text):
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:16]


def load(path):
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (ValueError, OSError) as exc:
        sys.exit(f"{path} exists but could not be read ({exc}).")


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
        # Any key name ending in api_key / api-key / apikey, any case, so the
        # file the deployed app already uses works without being renamed.
        pattern = re.compile(r'^\s*[\w.-]*api[_-]?key\s*=\s*["\'](.+?)["\']',
                             re.IGNORECASE)
        for line in open(secrets, encoding="utf-8"):
            m = pattern.match(line)
            if m and m.group(1).strip():
                return m.group(1).strip()
    sys.exit("No API key found.\n"
             "  Either:  $env:ANTHROPIC_API_KEY = 'sk-ant-...'\n"
             "  or add to .streamlit/secrets.toml:  "
             'ANTHROPIC_API_KEY = "sk-ant-..."')


def find_workbook(explicit=None):
    if explicit:
        if os.path.exists(explicit):
            return explicit
        sys.exit(f"Workbook not found: {explicit}")
    for w in WORKBOOKS:
        if os.path.exists(w):
            return w
    sys.exit("No workbook found. Pass --workbook.")


# Function words separate these three languages far more reliably than
# subject-matter phrases do, because every passage of real prose contains
# plenty of them whatever it is about.
_STOP = {
    "en": {"the", "of", "and", "to", "in", "for", "with", "is", "are", "be",
           "this", "that", "which", "as", "on", "by", "from", "an", "not",
           "has", "have", "been", "was", "were", "or", "at", "it"},
    "de": {"der", "die", "das", "und", "ist", "für", "nicht", "den", "dem",
           "des", "ein", "eine", "einer", "mit", "von", "zu", "im", "auf",
           "bei", "wird", "werden", "sind", "sich", "als", "oder", "durch"},
    "fr": {"le", "la", "les", "des", "est", "pour", "dans", "une", "un", "du",
           "de", "et", "par", "sur", "au", "aux", "ce", "cette", "qui", "que",
           "sont", "été", "avec", "plus", "ou", "en", "a"},
}
_WORD = re.compile(r"[a-zà-öø-ÿ]+", re.IGNORECASE)


def language_of(text):
    """Best guess at a passage's language, or None when there is too little
    prose to tell (a table of numbers, a heading, a fragment)."""
    words = [w.lower() for w in _WORD.findall(text)]
    if len(words) < 12:
        return None
    score = {lang: sum(w in stop for w in words) / len(words)
             for lang, stop in _STOP.items()}
    best = max(score, key=score.get)
    runner = max(k for k in score if k != best)
    # Needs real density AND a clear margin, so a German table with a few
    # English words in it is not mistaken for English.
    if score[best] < 0.06 or score[best] < score[runner] * 1.8:
        return None
    return best


def already_english(text):
    """True only when a passage is clearly English prose.

    Nearly half the G-BA corpus is published in English already, and sending
    it to a translator costs money to receive it back unchanged. Anything
    ambiguous is translated: paying twice is cheaper than leaving German on
    the screen.
    """
    return language_of(text) == "en"


def collect(workbook, min_chars):
    """Every distinct passage worth translating, with its language."""
    import pandas as pd
    jobs = {}
    skipped_en = 0
    for sheet, column, lang in SHEETS:
        try:
            df = pd.read_excel(workbook, sheet_name=sheet)
        except Exception:                                  # noqa: BLE001
            print(f"  (no sheet {sheet}, skipping)")
            continue
        if column not in df.columns:
            print(f"  (no column {column} in {sheet}, skipping)")
            continue
        n = 0
        for value in df[column].dropna().astype(str):
            text = value.strip()
            if len(text) < min_chars:
                continue
            if already_english(text):
                skipped_en += 1
                continue
            k = key_of(text)
            if k not in jobs:
                jobs[k] = (text, lang)
                n += 1
        print(f"  {sheet:24s} {n:6,} distinct passages")
    if skipped_en:
        print(f"  {'(already English, skipped)':24s} {skipped_en:6,}")
    return jobs


_NUMBER = re.compile(r"\d[\d   ]*(?:[.,]\d+)?")


def numbers_lost(source, translated):
    """Source numbers that do not survive verbatim into the translation.

    A translated figure that has been reformatted — 83,5% into 83.5%, or
    846 933 € into €846,933 — carries the same value but no longer matches
    the PDF the page link opens, which is the whole point of showing them
    side by side. This catches that; it is not a check on the prose.
    """
    if not translated:
        return []
    lost = []
    for token in _NUMBER.findall(source):
        token = token.strip("   .,")
        if len(token) < 2 or token in lost:
            continue
        if token not in translated:
            lost.append(token)
    return lost


def translate_one(client, model, text, lang, max_retry=5):
    glossary = GLOSSARY_DE if lang == "de" else GLOSSARY_FR
    # max_tokens is a ceiling, not a charge — only tokens actually generated
    # are billed — so it is set far above any real need. The previous budget
    # (400 + len/2) truncated number-dense French tables mid-row, because
    # "1 597,52 €" costs several tokens while contributing ten characters.
    budget = max(1024, min(8000, len(text) + 256))
    delay, last = 2.0, "no response"
    for attempt in range(max_retry):
        try:
            resp = client.messages.create(
                model=model,
                max_tokens=budget,
                system=SYSTEM.format(glossary=glossary),
                messages=[{"role": "user", "content": text}],
            )
            out = "".join(b.text for b in resp.content if b.type == "text").strip()
            stop = getattr(resp, "stop_reason", None)
            if out and stop != "max_tokens":
                return out
            if stop == "max_tokens":
                # Retrying at the same ceiling would truncate identically.
                last = f"truncated at {budget} tokens"
                budget = min(8000, budget * 2)
            else:
                kinds = ",".join(sorted({b.type for b in resp.content})) or "none"
                last = f"empty output (stop_reason={stop}, blocks={kinds})"
        except Exception as exc:                           # noqa: BLE001
            last = f"{type(exc).__name__}: {exc}"
        if attempt == max_retry - 1:
            break
        time.sleep(delay * (0.5 + random.random()))
        delay = min(delay * 2, 30)
    # An empty or truncated answer is a FAILURE, not a blank translation.
    # Returning None here would have let a full run drop thousands of
    # passages while reporting no errors at all.
    return f"__ERROR__ {last}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook")
    ap.add_argument("--model", default="claude-haiku-5-5")
    ap.add_argument("--limit", type=int)
    ap.add_argument("--sample", type=int,
                    help="translate N random passages and PRINT them beside "
                         "the original, writing nothing")
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--min-chars", type=int, default=40)
    ap.add_argument("--seed", type=int, default=20261110)
    args = ap.parse_args()

    try:
        import anthropic
    except ImportError:
        sys.exit("pip install anthropic")

    workbook = find_workbook(args.workbook)
    print(f"workbook: {workbook}")
    jobs = collect(workbook, args.min_chars)
    done = load(OUT)
    todo = [(k, t, l) for k, (t, l) in jobs.items() if k not in done]
    print(f"\ndistinct passages : {len(jobs):,}")
    print(f"already translated: {len(done):,}")
    print(f"to do now         : {len(todo):,}")

    client = anthropic.Anthropic(api_key=api_key())

    if args.sample:
        random.seed(args.seed)
        picked = random.sample(todo or list(
            (k, t, l) for k, (t, l) in jobs.items()), min(args.sample, len(jobs)))
        print(f"\nSAMPLE of {len(picked)} — nothing will be written")
        print(f"model: {args.model}\n")
        failed, drifted, echoed = [], [], []
        for i, (_k, text, lang) in enumerate(picked, 1):
            # The WHOLE passage, not a 1,500-character slice: the sample has
            # to put the same load through the model as the full run will.
            out = translate_one(client, args.model, text, lang)
            print(f"--- {i} [{lang}] {len(text):,} chars " + "-" * 40)
            print("  SOURCE :", " ".join(text[:300].split()))
            if str(out).startswith("__ERROR__"):
                failed.append(i)
                print("  FAILED :", str(out)[10:])
            else:
                print("  ENGLISH:", " ".join(str(out)[:300].split()))
                tail = " ".join(str(out)[-120:].split())
                if len(str(out)) > 300:
                    print("  ...ends:", tail)
                lost = numbers_lost(text, out)
                if lost:
                    drifted.append(i)
                    print("  NUMBERS CHANGED:", ", ".join(lost[:8]) +
                          (f" (+{len(lost) - 8} more)" if len(lost) > 8 else ""))
                if out.strip()[:80] == text.strip()[:80] and lang != "en":
                    echoed.append(i)
                    print("  NOT TRANSLATED: came back as the source text")
            print()

        n = len(picked)
        clean = n - len(set(failed) | set(drifted) | set(echoed))
        print("=" * 62)
        print(f"  {clean}/{n} clean")
        if failed:
            print(f"  {len(failed)} failed outright      : {failed}")
        if drifted:
            print(f"  {len(drifted)} reformatted a number : {drifted}")
        if echoed:
            print(f"  {len(echoed)} not translated       : {echoed}")
        print("\nNUMBERS CHANGED means the value is probably right but no longer\n"
              "matches the PDF the page link opens. A few on passages that are\n"
              "pure prose is noise; a pattern on tables is not.")
        print("\nAlso read the German and French yourself for: Zusatznutzen,\n"
              "zweckmäßige Vergleichstherapie, SMR, ASMR, RDCR.")
        return

    if args.limit:
        todo = todo[:args.limit]
    if not todo:
        print("\nNothing to do.")
        return

    chars = sum(len(t) for _k, t, _l in todo)
    print(f"characters        : {chars:,}")
    print(f"model             : {args.model}\n")

    started, errors = time.time(), 0
    with futures.ThreadPoolExecutor(max_workers=args.workers) as pool:
        pending = {pool.submit(translate_one, client, args.model, t, l): (k, t)
                   for k, t, l in todo}
        for i, fut in enumerate(futures.as_completed(pending), 1):
            k, _text = pending[fut]
            try:
                out = fut.result()
            except Exception as exc:                       # noqa: BLE001
                out = f"__ERROR__ {type(exc).__name__}: {exc}"
            if out and out.startswith("__ERROR__"):
                errors += 1
            elif out:
                done[k] = out
            if i % 25 == 0 or i == len(todo):
                save(OUT, done)
                rate = (time.time() - started) / i
                left = (len(todo) - i) * rate / 60
                print(f"  [{i:>6}/{len(todo)}] {len(done):,} cached, "
                      f"{errors} error(s), ~{left:.0f} min left")
    save(OUT, done)
    print(f"\n{len(done):,} translations in {OUT}")
    if errors:
        print(f"{errors} failed — re-run to retry them.")
    print(f"file size: {os.path.getsize(OUT)/1e6:.1f} MB")


if __name__ == "__main__":
    main()
