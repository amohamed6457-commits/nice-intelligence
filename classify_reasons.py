import argparse
import concurrent.futures as futures
import hashlib
import json
import os
import random
import re
import sys
import time

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

REASONS = "reasons.json"
OUT = "reason_categories_v6.json"
FAILLOG = "reason_failures.json"

CATEGORIES = {
    "COMPARATOR_NOT_IMPLEMENTED": (
        "The study did not implement the comparator therapy the G-BA "
        "specified (zweckmaessige Vergleichstherapie). Use this unless one "
        "of the two narrower comparator codes below fits.", "no"),
    "INADEQUATE_THERAPY_CONTINUED": (
        "Narrower comparator case: the control arm continued an "
        "already-inadequate therapy when escalation or dose optimisation "
        "was available and was not used.", "no"),
    "SINGLE_COMPARATOR_INSUFFICIENT": (
        "Narrower comparator case: the G-BA specified therapy at the "
        "physician's discretion, and a trial with one fixed comparator "
        "cannot implement that.", "partly"),
    "NO_COMPARATIVE_DATA": (
        "Single-arm study, or no direct or indirect comparison was "
        "presented at all.", "yes"),
    "INSUFFICIENT_FOLLOW_UP": (
        "Follow-up or study duration was too short to say whether effects "
        "persist, or long-term data are missing.", "yes"),
    "UNSUITABLE_STATISTICAL_METHOD": (
        "The statistical approach was not appropriate -- wrong model, "
        "unsuitable regression, inadequate adjustment, undocumented "
        "confounder selection, no balance or overlap diagnostics.",
        "partly"),
    "INDIRECT_COMPARISON_UNUSABLE": (
        "An indirect comparison or network meta-analysis was presented but "
        "is methodologically unsound or uninterpretable.", "partly"),
    "INCOMPARABLE_STUDIES": (
        "Two or more studies being compared are not comparable to each "
        "other -- different baseline populations, different endpoint "
        "definitions, different timepoints. Not about one study versus the "
        "appraised indication.", "partly"),
    "ENDPOINT_NOT_PATIENT_RELEVANT": (
        "The endpoint itself was not operationalised in a patient-relevant "
        "way, or cannot support a statement about added benefit.", "no"),
    "SEARCH_INCOMPLETE": (
        "The literature search or study pool is incomplete or cannot be "
        "shown to be complete.", "no"),
    "POPULATION_MISMATCH": (
        "The population studied does not match the appraised indication, "
        "including transfer of results from one population to another "
        "(adults to adolescents).", "yes"),
    "EFFECT_TOO_SMALL_VS_BIAS": (
        "The observed effect is not large enough to rule out that it arises "
        "from systematic bias. This is the dramatic-effect test for "
        "single-arm evidence, NOT a general 'looks biased' category.", "no"),
    "NOTHING_SUBMITTED": (
        "No evidence at all was submitted where the company was required "
        "to submit it. Not for 'submitted plenty, none of it suitable'.",
        "yes"),
    "NO_ADVANTAGE_SHOWN": (
        "The evidence WAS usable and was assessed, and it showed no "
        "advantage over the comparator. A substantive verdict, not an "
        "evidence failure.", "no"),
    "REASON_NOT_STATED": (
        "The sentence says evidence was unsuitable, unusable or could not "
        "be used, but gives no cause for it.", "unknown"),
    "NOT_A_REASON": (
        "Affirmatively one of: a heading, a table caption, a scope or "
        "signpost statement, a description of a drug or a study, a "
        "statement of agreement, a statutory statement, or a criticism of "
        "patient numbers or treatment costs. NOT a fallback for anything "
        "unclear.", "n/a"),
}

FAMILY = {
    "COMPARATOR_NOT_IMPLEMENTED": "comparator",
    "INADEQUATE_THERAPY_CONTINUED": "comparator",
    "SINGLE_COMPARATOR_INSUFFICIENT": "comparator",
}

SYSTEM = """You classify single sentences from German HTA dossier assessments
(IQWiG Dossierbewertungen) by WHY the submitted evidence could not be used.

You are given one sentence, sometimes with the sentence before and after it.
Return exactly one category code from the list, and nothing else -- no
explanation, no punctuation, no quotes.

RULES
1. Classify what the SENTENCE says, not what you infer about the appraisal.
2. NOT_A_REASON is an AFFIRMATIVE finding, never a fallback. Use it only
   when the sentence IS one of these: a heading; a table caption or table
   row; a statement of what the assessment covers; a signpost that reasons
   follow ("nachfolgend wird begruendet, warum..."); a description of a drug
   or of a study's design; a statement that the company AGREED with the
   G-BA; a statement of what the law requires or exempts; or a criticism of
   patient numbers, prevalence or treatment costs rather than of evidence.
   If none of those fit, do NOT use NOT_A_REASON.
3. If the sentence is about evidence being unusable but you cannot tell why,
   return REASON_NOT_STATED. That is the code for uncertainty, not
   NOT_A_REASON. "The data submitted are not suitable for the assessment",
   with no cause, is REASON_NOT_STATED even when a heading is glued to the
   front of it and even when the next sentence says the reason follows.
4. You may be shown CONTEXT BEFORE and CONTEXT AFTER. Classify the
   SENTENCE. Use the context only to resolve a cause the sentence defers to
   ("somit", "daher", "aus den oben genannten Gruenden"). If the context is
   about a different study, endpoint or question, ignore it. Context that
   merely announces a reason without giving it does not make a cause.
5. A sentence reporting a RESULT is not a reason. If it reports that the
   evidence was assessed and showed no advantage, that is
   NO_ADVANTAGE_SHOWN. If it reports a measured effect either way, that is
   NOT_A_REASON.
6. Comparator precedence. If the sentence says optimisation or escalation
   was available and not used, return INADEQUATE_THERAPY_CONTINUED. If it
   says a single fixed comparator cannot implement therapy at the
   physician's discretion, return SINGLE_COMPARATOR_INSUFFICIENT.
   Otherwise, for any comparator arm that is not the one specified, return
   COMPARATOR_NOT_IMPLEMENTED.
7. COMPARATOR_NOT_IMPLEMENTED vs NO_COMPARATIVE_DATA: a comparative study
   exists but with the wrong comparator is the first; no comparison at all,
   a single-arm study, is the second.
8. POPULATION_MISMATCH vs INCOMPARABLE_STUDIES: one study's population not
   matching the appraised indication is POPULATION_MISMATCH. Two studies
   not matching EACH OTHER is INCOMPARABLE_STUDIES.
9. NOTHING_SUBMITTED means nothing was submitted. "No SUITABLE data were
   submitted" is not that: if the sentence says why nothing was suitable,
   use that cause; if it does not, use REASON_NOT_STATED.
10. Two opposite statements are made with Paragraph 35a. Read the VERB,
   not the paragraph number:
   - "Nachweise ... muessen nicht vorgelegt werden" -- evidence need NOT be
     submitted, by statute, usually the orphan exemption. NOT_A_REASON.
   - "der pU hat die erforderlichen Nachweise nicht vorgelegt", "hat kein
     Dossier eingereicht", "wurden dem G-BA nicht vorgelegt" -- the company
     DID NOT submit what it was required to submit. NOTHING_SUBMITTED.
11. "Auch wenn", "selbst wenn", "even if", "regardless of" introduce a
   concession. The reason is what follows, not what is conceded.
12. If a sentence gives more than one reason, return the one it leads with.

CATEGORIES
{categories}"""


def key_of(text):
    return hashlib.sha1(text.strip().encode("utf-8")).hexdigest()[:16]


def load(path, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (ValueError, OSError) as exc:
        sys.exit(f"{path} could not be read ({exc}).")


def save(path, payload):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def api_key():
    for var in ("ANTHROPIC_API_KEY", "CLAUDE_API_KEY"):
        if os.environ.get(var):
            return os.environ[var]
    secrets = os.path.join(".streamlit", "secrets.toml")
    if os.path.exists(secrets):
        pattern = re.compile(r'^\s*[\w.-]*api[_-]?key\s*=\s*["\'](.+?)["\']',
                             re.IGNORECASE)
        for line in open(secrets, encoding="utf-8"):
            m = pattern.match(line)
            if m and m.group(1).strip():
                return m.group(1).strip()
    sys.exit("No API key. Set ANTHROPIC_API_KEY or put it in "
             ".streamlit/secrets.toml.")


def prompt():
    lines = [f"{code}: {text}" for code, (text, _rwe) in CATEGORIES.items()]
    return SYSTEM.format(categories="\n".join(lines))


def payload(sentence, before, after):
    if not (before or after):
        return sentence
    parts = []
    if before:
        parts.append(f"CONTEXT BEFORE: {before}")
    parts.append(f"SENTENCE: {sentence}")
    if after:
        parts.append(f"CONTEXT AFTER: {after}")
    return "\n\n".join(parts)


NUDGE = ("\n\nAnswer with one category code from the list and nothing "
         "else. No explanation.")


def extract_code(raw):
    text = str(raw or "").upper()
    for candidate in (re.sub(r"[^A-Z]+", "_", text).strip("_"),
                      re.sub(r"[^A-Z_]", "", text)):
        hits = [(candidate.find(c), c) for c in CATEGORIES if c in candidate]
        if hits:
            hits.sort(key=lambda h: (h[0], -len(h[1])))
            return hits[0][1]
    return None


def classify_one(client, model, sentence, before="", after="", max_retry=4):
    delay, last = 2.0, "no response"
    body = payload(sentence, before, after)
    for attempt in range(max_retry):
        try:
            resp = client.messages.create(
                model=model,
                max_tokens=64,
                temperature=0 if attempt == 0 else 0.3,
                system=prompt(),
                messages=[{"role": "user",
                           "content": body if attempt == 0
                           else body + NUDGE}],
            )
            raw = "".join(b.text for b in resp.content
                          if b.type == "text").strip()
            stop = getattr(resp, "stop_reason", None)
            got = extract_code(raw)
            if got:
                return got
            if stop == "max_tokens":
                last = f"truncated at 64 tokens: {raw[:60]!r}"
            elif not raw:
                last = f"empty output (stop_reason={stop})"
            else:
                last = f"no code found in {raw[:70]!r}"
        except Exception as exc:
            last = f"{type(exc).__name__}: {exc}"
        if attempt == max_retry - 1:
            break
        time.sleep(delay * (0.5 + random.random()))
        delay = min(delay * 2, 20)
    return f"__ERROR__ {last}"


def every_sentence(data):
    out = []
    for aid, entries in data.items():
        for entry in entries:
            text = str(entry.get("sentence") or "").strip()
            if len(text) >= 40:
                out.append((aid, entry.get("page"), text,
                            str(entry.get("before") or "").strip(),
                            str(entry.get("after") or "").strip()))
    return out


def compare(old_path, new_path):
    from collections import Counter
    old = load(old_path, None)
    new = load(new_path, None)
    if old is None or new is None:
        sys.exit(f"Need both {old_path} and {new_path}.")
    shared = [k for k in new if k in old
              and not str(old[k]).startswith("__")
              and not str(new[k]).startswith("__")]
    print(f"{old_path}: {len(old):,}")
    print(f"{new_path}: {len(new):,}")
    print(f"in both   : {len(shared):,}\n")
    if not shared:
        return
    moved = [k for k in shared if old[k] != new[k]]
    print(f"changed label: {len(moved):,} of {len(shared):,} "
          f"({len(moved) / len(shared):.0%})\n")
    was_unstated = [k for k in shared if old[k] == "REASON_NOT_STATED"]
    rescued = [k for k in was_unstated if new[k] != "REASON_NOT_STATED"]
    print(f"was REASON_NOT_STATED : {len(was_unstated):,}")
    if was_unstated:
        print(f"  now carries a cause : {len(rescued):,} "
              f"({len(rescued) / len(was_unstated):.0%})")
        for code, n in Counter(new[k] for k in rescued).most_common(8):
            print(f"      {n:>4}  -> {code}")
    gained = [k for k in shared if old[k] != "REASON_NOT_STATED"
              and new[k] == "REASON_NOT_STATED"]
    if gained:
        print(f"\nmoved INTO REASON_NOT_STATED: {len(gained):,} "
              f"-- context should not do this; inspect if large")
    print("\nThis says the labels moved. It does NOT say they moved to the "
          "right answer.\nThat needs a fresh blind benchmark on the new "
          "file.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--reasons", default=REASONS)
    ap.add_argument("--model", default="claude-haiku-4-5-20251001")
    ap.add_argument("--sample", type=int)
    ap.add_argument("--limit", type=int)
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--seed", type=int, default=20261113)
    ap.add_argument("--audit", action="store_true")
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--compare")
    ap.add_argument("--why", action="store_true")
    args = ap.parse_args()

    if args.why:
        from collections import Counter
        bad = load(FAILLOG, {})
        if not bad:
            sys.exit(f"{FAILLOG} is empty — nothing has failed.")
        print(f"{len(bad):,} unresolved failure(s)\n")
        kinds = Counter(str(v.get("reason", ""))[:60] for v in bad.values())
        for text, n in kinds.most_common(15):
            print(f"  {n:>5}  {text}")
        print("\nexamples:")
        for key, info in list(bad.items())[:6]:
            print(f"\n  {info.get('reason', '')[:140]}")
            print(f"    {' '.join(str(info.get('sentence', ''))[:170].split())}")
        return

    if args.compare:
        compare(args.compare, args.out)
        return

    if args.audit and not os.path.exists(args.out):
        sys.exit(f"{args.out} does not exist. Nothing has been classified "
                 f"into it yet.")
    data = load(args.reasons, None)
    if data is None:
        sys.exit(f"{args.reasons} not found — run extract_rejection_reasons "
                 "first.")
    sentences = every_sentence(data)
    done = load(args.out, {})
    withctx = sum(1 for _a, _p, _t, b, a in sentences if b or a)
    print(f"{len(sentences):,} sentences from {len(data):,} appraisals")
    print(f"carrying context  : {withctx:,}")
    print(f"writing to        : {args.out}")
    print(f"already classified: {len(done):,}")

    if args.audit:
        from collections import Counter
        counts = Counter(v for v in done.values()
                         if not str(v).startswith("__"))
        total = sum(counts.values())
        if not total:
            sys.exit("Nothing classified yet.")
        print(f"\n{total:,} classified\n")
        for code, n in counts.most_common():
            rwe = CATEGORIES.get(code, ("", "?"))[1]
            print(f"  {n:>5}  {n / total:>5.1%}  {code:<32} RWE: {rwe}")
        nar = counts.get("NOT_A_REASON", 0)
        unstated = counts.get("REASON_NOT_STATED", 0)
        addressable = sum(n for code, n in counts.items()
                          if CATEGORIES.get(code, ("", "no"))[1] == "yes")
        wide = total - nar
        narrow = total - nar - unstated
        print(f"\n  not a reason at all          : {nar:,}")
        print(f"  state no cause               : {unstated:,}")
        print(f"  state a readable cause       : {narrow:,}")
        print(f"\n  real-world evidence could address {addressable:,} of these.")
        if narrow and wide:
            lo, hi = addressable / wide, addressable / narrow
            print(f"  As a share that is {lo:.0%} to {hi:.0%}, not one number:")
            print(f"    {lo:.0%}  counting the {unstated:,} unstated as "
                  f"not addressable")
            print(f"    {hi:.0%}  excluding them, i.e. of causes we can read")
            print(f"  Quote the range. The {unstated:,} unstated are unread, "
                  f"not negative.")
        print("\nUNIT: one distinct sentence. Identical sentences across "
              "appraisals are\ncounted once, so this is the spread of "
              "phrasings, NOT the share of\nappraisals. Do not read a "
              "per-appraisal rate off this table.")
        print("\nThe RWE column is a fixed property of each category, set in "
              "this file,\nnot something the model decided per sentence.")
        return

    try:
        import anthropic
    except ImportError:
        sys.exit("pip install anthropic")
    client = anthropic.Anthropic(api_key=api_key())

    if args.sample:
        random.seed(args.seed)
        picked = random.sample(sentences, min(args.sample, len(sentences)))
        print(f"\nSAMPLE of {len(picked)} — nothing will be written\n")
        for i, (aid, page, text, before, after) in enumerate(picked, 1):
            got = classify_one(client, args.model, text, before, after)
            print(f"--- {i}  {aid} p{page}")
            if before:
                print(f"    [before] ...{' '.join(before[-110:].split())}")
            print(f"    {' '.join(text[:230].split())}")
            if after:
                print(f"    [after ] {' '.join(after[:110].split())}...")
            print(f"    -> {got}\n")
        print("Read these. A wrong label here is a claim about why a "
              "regulator\nrejected evidence, which is worse than a wrong "
              "table.")
        return

    todo, seen = [], set()
    for a, p, t, b, x in sentences:
        k = key_of(t)
        if k in done or k in seen:
            continue
        seen.add(k)
        todo.append((a, p, t, b, x))
    if args.limit:
        todo = todo[:args.limit]
    print(f"to do now         : {len(todo):,}\n")
    if not todo:
        print("Nothing to do. Run --audit to see the distribution.")
        return

    started, errors, failures = time.time(), 0, load(FAILLOG, {})
    with futures.ThreadPoolExecutor(max_workers=args.workers) as pool:
        pending = {pool.submit(classify_one, client, args.model, t, b, x): t
                   for _a, _p, t, b, x in todo}
        for i, fut in enumerate(futures.as_completed(pending), 1):
            text = pending[fut]
            try:
                got = fut.result()
            except Exception as exc:
                got = f"__ERROR__ {type(exc).__name__}: {exc}"
            if str(got).startswith("__ERROR__"):
                errors += 1
                failures[key_of(text)] = {
                    "reason": str(got)[10:].strip(),
                    "sentence": text[:300]}
            else:
                done[key_of(text)] = got
                failures.pop(key_of(text), None)
            if i % 50 == 0 or i == len(todo):
                save(args.out, done)
                save(FAILLOG, failures)
                rate = (time.time() - started) / i
                print(f"  [{i:>5}/{len(todo)}] {len(done):,} done, "
                      f"{errors} error(s), ~{(len(todo)-i)*rate/60:.0f} min left")
    save(args.out, done)
    save(FAILLOG, failures)
    print(f"\n{len(done):,} classified -> {args.out}")
    if errors:
        print(f"{errors} failed -> {FAILLOG}")
        print("Run --why to see what they said. Re-running retries them.")
    print("\nRun --audit for the distribution, then benchmark it before "
          "putting it on screen.")


if __name__ == "__main__":
    main()
