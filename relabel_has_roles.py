import argparse
import json
import os
import sys
from collections import Counter

from extract_has_sections import role_of

SECTIONS = "has_clinical_sections.json"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sections", default=SECTIONS)
    ap.add_argument("--write", action="store_true")
    args = ap.parse_args()
    if not os.path.exists(args.sections):
        sys.exit(f"{args.sections} not found.")
    with open(args.sections, encoding="utf-8") as fh:
        data = json.load(fh)

    moves, before, after = Counter(), Counter(), Counter()
    for doc in data.values():
        for s in doc.get("sections", []):
            old = s["role"]
            new = role_of(s["title"])
            before[old] += 1
            after[new] += 1
            if new != old:
                moves[(old, new)] += 1
            s["role"] = new

    full = {a: d for a, d in data.items()
            if not d.get("synthese") and d.get("sections")}
    print(f"{len(data):,} appraisals, {len(full):,} full avis\n")
    print("reassigned sections:")
    for (old, new), n in moves.most_common(14):
        print(f"  {n:>6}  {old} -> {new}")
    print(f"\n{sum(moves.values()):,} sections changed role\n")
    print(f"{'role':<20} {'before':>8} {'after':>8}   docs with it")
    for role in sorted(set(before) | set(after)):
        docs = sum(1 for d in full.values()
                   if any(s["role"] == role for s in d["sections"]))
        print(f"  {role:<18} {before.get(role, 0):>8,} "
              f"{after.get(role, 0):>8,}   {docs:>5} "
              f"({docs / max(1, len(full)):>4.0%})")

    for group, label in ((("clinical_evidence",), "clinical evidence"),
                         (("clinical_evidence", "discussion"),
                          "clinical evidence or discussion")):
        docs = sum(1 for d in full.values()
                   if any(s["role"] in group for s in d["sections"]))
        print(f"\nfull avis with {label}: {docs:,} of {len(full):,} "
              f"({docs / max(1, len(full)):.0%})")

    if args.write:
        tmp = args.sections + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False)
        os.replace(tmp, args.sections)
        print(f"\n-> {args.sections} rewritten")
    else:
        print("\nNothing written. Add --write to apply.")


if __name__ == "__main__":
    main()
