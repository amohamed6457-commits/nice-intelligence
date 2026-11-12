import argparse
import collections
import json
import os
import re
import sys

INDEX = "chapter_index.json"
SECTIONS = "chapter_sections.json"


def load(path):
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (ValueError, OSError) as exc:
        print(f"  ! {path} could not be read: {exc}")
        return None


# The real layout, confirmed against the project's own index: two buckets at
# the top level, each keyed by TA number. Detecting this matters — read as a
# flat mapping it looks like a file with exactly two records called 'done'
# and 'failed'.
_BUCKETS = {"done": "ok", "ok": "ok", "success": "ok", "succeeded": "ok",
            "failed": "failed", "errors": "failed", "failures": "failed",
            "skipped": "skipped", "pending": "pending"}


def buckets(blob):
    """
    (bucket_name, {id: value}) for a bucketed index, or None if it is not
    one. A dict counts as bucketed when every top-level key is a known
    bucket name and holds a dict or list.
    """
    if not isinstance(blob, dict) or not blob:
        return None
    if not all(str(k).lower() in _BUCKETS and isinstance(v, (dict, list))
               for k, v in blob.items()):
        return None
    out = []
    for name, payload in blob.items():
        if isinstance(payload, dict):
            entries = {str(k): v for k, v in payload.items()}
        else:
            entries = {str(item): None for item in payload}
        out.append((str(name), entries))
    return out


def records(blob):
    """
    Yield (key, value) for whatever shape the file turns out to be: a dict
    keyed by appraisal, a list of dicts, or a dict with the real payload
    one level down under a key like 'records' or 'appraisals'.
    """
    if isinstance(blob, dict):
        inner = next((v for k, v in blob.items()
                      if k.lower() in {"records", "appraisals", "index", "items", "data"}
                      and isinstance(v, (dict, list))), None)
        if inner is not None and len(blob) <= 4:
            yield from records(inner)
            return
        for k, v in blob.items():
            yield str(k), v
    elif isinstance(blob, list):
        for item in blob:
            if isinstance(item, dict):
                key = next((str(item[k]) for k in
                            ("appraisal_id", "id", "ta", "ta_id", "slug")
                            if k in item), None)
                yield key or "?", item
            else:
                yield str(item), {}


def describe(value):
    """A short reason string for one record, whatever keys it happens to use."""
    if isinstance(value, str):
        # In the failed bucket the value IS the reason. Treating a non-empty
        # string as merely "ok" threw away exactly the text we are here to
        # read.
        text = re.sub(r"\s+", " ", value).strip()
        if not text:
            return "empty value"
        return "ok" if text.lower() in {"ok", "done", "success", "fetched"} \
            else text[:90]
    if not isinstance(value, dict):
        return "ok" if value else "empty value"
    for key in ("error", "reason", "failure", "why", "message", "status"):
        if key in value and value[key]:
            text = str(value[key]).strip()
            if text.lower() in {"ok", "success", "done", "fetched", "200"}:
                return "ok"
            return re.sub(r"\s+", " ", text)[:90]
    for key in ("chapters", "sections", "urls", "links"):
        if key in value:
            n = value[key]
            if isinstance(n, (list, dict)):
                return "ok" if len(n) else "no chapter links found"
            if isinstance(n, int):
                return "ok" if n else "no chapter links found"
    return "no status recorded"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--index", default=INDEX)
    ap.add_argument("--sections", default=SECTIONS)
    ap.add_argument("--list", type=int, default=10, help="failing ids to show")
    ap.add_argument("--dump", help="print everything held for one appraisal id")
    args = ap.parse_args()

    index = load(args.index)
    if index is None:
        sys.exit(f"{args.index} not found here. Run this from the folder that holds it.")

    grouped = buckets(index)
    if grouped:
        rows, origin = {}, {}
        for name, entries in grouped:
            print(f"{args.index} · {name}: {len(entries):,} appraisals")
            for key, value in entries.items():
                rows[key] = value
                origin[key] = str(name).lower()
        print()
    else:
        rows = dict(records(index))
        origin = {}
        print(f"{args.index}: {len(rows):,} entries\n")

    if args.dump:
        want = args.dump.strip().lower()
        hit = next((f"{k}: {json.dumps(v, indent=2, ensure_ascii=False)[:2500]}"
                    for k, v in rows.items() if want in k.lower()), None)
        print(hit or f"no entry matching {args.dump!r}")
        return

    # Shape report first: what the entries actually look like.
    sample = next(iter(rows.values()), None)
    if isinstance(sample, dict):
        print("keys present on a record:", ", ".join(sorted(sample)[:12]))
    else:
        print("records are not dicts; sample value:", repr(sample)[:120])
    print()

    sections = load(args.sections)
    have = set()
    if sections is not None:
        for key, value in records(sections):
            if value:
                have.add(str(key).strip().lower())
        print(f"{args.sections}: {len(have):,} appraisals with parsed sections\n")

    reasons = collections.Counter()
    failing = []
    for key, value in rows.items():
        norm = str(key).strip().lower()
        bucket = _BUCKETS.get(origin.get(key, ""), None)
        if have and (norm in have or f"ta{norm}" in have
                     or norm.lstrip("ta0") in have):
            reasons["ok — sections parsed"] += 1
            continue
        if bucket == "ok":
            reasons["listed as done, but no parsed sections"] += 1
            failing.append((key, "done but no sections — parse step dropped it"))
            continue
        why = describe(value)
        if bucket == "failed" and why in {"no status recorded", "ok"}:
            # The failed bucket often stores just a timestamp or True; the
            # bucket itself is the status in that case.
            why = "in failed bucket, no reason recorded"
        reasons[why] += 1
        if why != "ok":
            failing.append((key, why))

    print("why each appraisal does or does not have chapter text:")
    width = max(len(r) for r in reasons) if reasons else 10
    for why, n in reasons.most_common():
        print(f"  {n:5d}  {why[:width]}")

    if failing:
        print(f"\nfirst {min(args.list, len(failing))} without chapter text:")
        for key, why in failing[:args.list]:
            print(f"  {key:<12} {why}")
        print("\nOpen two or three of those on nice.org.uk. If they are terminated "
              "appraisals,\nnon-submissions or fast-track approvals, there is no "
              "committee discussion to\nfetch and they should be excluded from the "
              "coverage denominator rather than\ncounted as failures.")


if __name__ == "__main__":
    main()
