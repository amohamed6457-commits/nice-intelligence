import argparse
import concurrent.futures as futures
import hashlib
import json
import os
import random
import re
import sys
import time

REASONS = "reasons.json"
OUT = "reason_categories.json"

CATEGORIES = {
    "COMPARATOR_NOT_IMPLEMENTED": (
        "The study did not implement the comparator therapy the G-BA "
        "specified (zweckmäßige Vergleichstherapie).", "no"),
    "INADEQUATE_THERAPY_CONTINUED": (
        "The control arm continued an already-inadequate therapy instead of "
        "escalating or optimising it, which is not German practice.", "no"),
    "SINGLE_COMPARATOR_INSUFFICIENT": (
        "The comparator was 'therapy at the physician's discretion' and a "
        "trial with one fixed comparator cannot implement that.", "partly"),
    "NO_COMPARATIVE_DATA": (
        "Single-arm study, or no direct or indirect comparison was "
        "presented at all.", "yes"),
    "INSUFFICIENT_FOLLOW_UP": (
        "Follow-up or study duration was too short to say whether effects "
        "persist, or long-term data are missing.", "yes"),
    "UNSUITABLE_STATISTICAL_METHOD": (
        "The statistical approach used was not appropriate -- the wrong "
        "model, an unsuitable regression, or an inadequate adjustment.",
        "no"),
    "INDIRECT_COMPARISON_UNUSABLE": (
        "An indirect comparison or network meta-analysis was presented but "
        "is methodologically unsound or uninterpretable.", "partly"),
    "ENDPOINT_NOT_PATIENT_RELEVANT": (
        "The endpoint was not operationalised in a patient-relevant way, or "
        "cannot support a statement about added benefit.", "no"),
    "SEARCH_INCOMPLETE": (
        "The literature search or study pool is incomplete or cannot be "
        "shown to be complete.", "no"),
    "POPULATION_MISMATCH": (
        "The population studied does not match the appraised indication, or "
        "populations being compared are not sufficiently similar.", "yes"),
    "RISK_OF_BIAS": (
        "The effect is not large enough to rule out that it arises from "
        "systematic bias.", "no"),
    "NOTHING_SUBMITTED": (
        "No evidence was submitted for the assessment.", "yes"),
    "REASON_NOT_STATED": (
        "The sentence says evidence was unsuitable or could not be used but "
        "gives no cause for it.", "unknown"),
    "NOT_A_REASON": (
        "Not a rejection reason: a heading, a table caption, a scope "
        "statement, a description of the drug, a reported result, or a "
        "statement of agreement.", "n/a"),
}

SYSTEM = """You classify single sentences from German HTA dossier assessments
(IQWiG Dossierbewertungen) by WHY the submitted evidence could not be used.

You are given one sentence. Return exactly one category code from the list,
and nothing else -- no explanation, no punctuation, no quotes.

RULES
1. Classify what the sentence SAYS, not what you infer about the appraisal.
2. If the sentence does not state a reason evidence was unusable, return
   NOT_A_REASON. Headings, table captions, scope statements, descriptions of
   what a drug is, and statements that the company AGREED with the G-BA are
   all NOT_A_REASON.
3. A sentence REPORTING A RESULT is NOT_A_REASON, whichever way the result
   went. "A statistically significant effect in favour of X was shown" is a
   finding, not a reason evidence could not be used.
4. If the sentence says evidence was unsuitable but gives NO cause -- "study
   Y is not suitable for the assessment", with nothing further -- return
   REASON_NOT_STATED. Do not supply a cause the sentence does not give.
5. COMPARATOR_NOT_IMPLEMENTED vs NO_COMPARATIVE_DATA: if a comparative study
   exists but its comparator arm is not the one the G-BA specified, that is
   COMPARATOR_NOT_IMPLEMENTED. If there is no comparison at all -- a
   single-arm study, or nothing comparative submitted -- that is
   NO_COMPARATIVE_DATA.
6. A statement of what the law requires or exempts is NOT_A_REASON.
   "Nachweise ... muessen nicht vorgelegt werden (Paragraph 35a ...)" says
   evidence need NOT be submitted, by statute -- usually the orphan-drug
   exemption. That is not a failure to submit evidence and must never be
   classified as NOTHING_SUBMITTED. Use NOTHING_SUBMITTED only where the
   company was required to submit evidence and did not.
7. If a sentence gives more than one reason, return the one it leads with.
8. Do not guess. NOT_A_REASON is the correct answer for anything unclear.

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


def classify_one(client, model, sentence, max_retry=4):
    delay, last = 2.0, "no response"
    for attempt in range(max_retry):
        try:
            resp = client.messages.create(
                model=model,
                max_tokens=24,
                temperature=0,
                system=prompt(),
                messages=[{"role": "user", "content": sentence}],
            )
            out = "".join(b.text for b in resp.content
                          if b.type == "text").strip().upper()
            out = re.sub(r"[^A-Z_]", "", out)
            if out in CATEGORIES:
                return out
            last = f"unrecognised answer {out[:40]!r}"
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
                out.append((aid, entry.get("page"), text))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--reasons", default=REASONS)
    ap.add_argument("--model", default="claude-haiku-4-5-20251001")
    ap.add_argument("--sample", type=int)
    ap.add_argument("--limit", type=int)
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--seed", type=int, default=20261113)
    ap.add_argument("--audit", action="store_true")
    args = ap.parse_args()

    data = load(args.reasons, None)
    if data is None:
        sys.exit(f"{args.reasons} not found — run extract_rejection_reasons "
                 "first.")
    sentences = every_sentence(data)
    done = load(OUT, {})
    print(f"{len(sentences):,} sentences from {len(data):,} appraisals")
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
        real = total - counts.get("NOT_A_REASON", 0)
        addressable = sum(n for code, n in counts.items()
                          if CATEGORIES.get(code, ("", "no"))[1] == "yes")
        print(f"\n  actual reasons (excluding NOT_A_REASON): {real:,}")
        if real:
            print(f"  of those, real-world evidence could address: "
                  f"{addressable:,} ({addressable / real:.0%})")
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
        for i, (aid, page, text) in enumerate(picked, 1):
            got = classify_one(client, args.model, text)
            print(f"--- {i}  {aid} p{page}")
            print(f"    {' '.join(text[:230].split())}")
            print(f"    -> {got}\n")
        print("Read these. A wrong label here is a claim about why a "
              "regulator\nrejected evidence, which is worse than a wrong "
              "table.")
        return

    todo = [(a, p, t) for a, p, t in sentences if key_of(t) not in done]
    if args.limit:
        todo = todo[:args.limit]
    print(f"to do now         : {len(todo):,}\n")
    if not todo:
        print("Nothing to do. Run --audit to see the distribution.")
        return

    started, errors = time.time(), 0
    with futures.ThreadPoolExecutor(max_workers=args.workers) as pool:
        pending = {pool.submit(classify_one, client, args.model, t): t
                   for _a, _p, t in todo}
        for i, fut in enumerate(futures.as_completed(pending), 1):
            text = pending[fut]
            try:
                got = fut.result()
            except Exception as exc:
                got = f"__ERROR__ {type(exc).__name__}: {exc}"
            if str(got).startswith("__ERROR__"):
                errors += 1
            else:
                done[key_of(text)] = got
            if i % 50 == 0 or i == len(todo):
                save(OUT, done)
                rate = (time.time() - started) / i
                print(f"  [{i:>5}/{len(todo)}] {len(done):,} done, "
                      f"{errors} error(s), ~{(len(todo)-i)*rate/60:.0f} min left")
    save(OUT, done)
    print(f"\n{len(done):,} classified -> {OUT}")
    if errors:
        print(f"{errors} failed — re-run to retry them.")
    print("\nRun --audit for the distribution, then benchmark it before "
          "putting it on screen.")


if __name__ == "__main__":
    main()
