import argparse
import json
import os
import random
import re
from collections import defaultdict

import pandas as pd

BASE = "https://www.has-sante.fr"

GUIDE = [
    ("What this sheet measures",
     "Whether the endpoints the model pulled out of each avis are really "
     "there, stated correctly, and on the page it claims."),
    ("", ""),
    ("Sheet 1 - Endpoints",
     "One row per extracted endpoint. Open the Link, go to the page range "
     "in Pages, and check that row against the PDF."),
    ("  Endpoint_OK",
     "Y  = this endpoint is reported in that section. "
     "N  = it is not there at all. "
     "PARTLY = it is there but named loosely or merged with another."),
    ("  Result_OK",
     "Y  = the number in Result matches the PDF. "
     "N  = the number is wrong, or invented. "
     "BLANK = Result column was empty, nothing to check."),
    ("  Direction_OK",
     "Y  = favours_drug / favours_comparator / no_difference / not_stated "
     "matches what the avis concludes for that endpoint. N = it does not."),
    ("  Page_OK",
     "Y  = the endpoint appears somewhere in the Pages range. "
     "N  = it is in the document but outside that range."),
    ("  Notes", "Where you actually found it, and what the avis says."),
    ("", ""),
    ("Sheet 2 - Missing",
     "One row per appraisal, naming the sections to read. List any "
     "endpoint those sections report that the model did NOT extract. "
     "This is the recall half - precision alone cannot show what was "
     "skipped. Sheet 1 shows a sample of what was extracted, so compare "
     "against Extracted_total, not against the rows you see."),
    ("", ""),
    ("Sheet 3 - Empty_sections",
     "Sections where the model returned nothing. Confirm that is right. "
     "If an avis section does report endpoints and the model returned "
     "none, write them in - a false 'no endpoints' is the worst failure "
     "mode because it is silent."),
    ("", ""),
    ("Do not", "Do not correct the model from memory or from the abstract. "
               "Only the section named in the Section column counts."),
]


def newest(pattern, folder="."):
    hits = [f for f in os.listdir(folder) if re.fullmatch(pattern, f)]
    if not hits:
        return None
    return max(hits, key=lambda f: os.path.getmtime(os.path.join(folder, f)))


def next_free(stem, ext=".xlsx"):
    if not os.path.exists(stem + ext):
        return stem + ext
    n = 2
    while os.path.exists(f"{stem} {n}{ext}"):
        n += 1
    return f"{stem} {n}{ext}"


def norm_id(x):
    return re.sub(r"[^A-Za-z0-9]+", "", str(x or "")).upper()


def norm_num(x):
    t = str(x or "").strip()
    parts = [p.lstrip("0") or "0" for p in t.split(".")]
    while len(parts) > 1 and parts[-1] == "0":
        parts.pop()
    return ".".join(parts)


def link_of(doc):
    url = doc.get("url") or ""
    if url.startswith("http"):
        return url
    if url:
        return BASE + ("" if url.startswith("/") else "/") + url
    return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--endpoints", default=None)
    ap.add_argument("--sections", default="has_sections.json")
    ap.add_argument("--sample", type=int, default=8)
    ap.add_argument("--empty", type=int, default=4)
    ap.add_argument("--per-section", type=int, default=3)
    ap.add_argument("--seed", type=int, default=20261117)
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    src = args.endpoints or newest(r"has_endpoints.*\.json")
    if not src:
        raise SystemExit("no has_endpoints*.json found - pass --endpoints")
    done = json.load(open(src, encoding="utf-8"))
    print(f"reading {src}  ({len(done):,} sections)")

    meta = {}
    if os.path.exists(args.sections):
        for aid, doc in json.load(open(args.sections, encoding="utf-8")).items():
            meta[norm_id(aid)] = doc
    else:
        print(f"warning: {args.sections} not found - no links or drug names")

    by_aid = defaultdict(dict)
    for key, val in done.items():
        if not isinstance(val, list):
            continue
        aid, _, number = key.partition("#")
        by_aid[aid][number or "?"] = val

    filled = sorted(a for a, secs in by_aid.items()
                    if any(v for v in secs.values()))
    blanks = sorted((a, n) for a, secs in by_aid.items()
                    for n, v in secs.items() if not v)
    if not filled:
        raise SystemExit("no appraisal has any endpoint - nothing to score")

    random.seed(args.seed)
    pick = sorted(random.sample(filled, min(args.sample, len(filled))))
    pick_empty = random.sample(blanks, min(args.empty, len(blanks)))

    def doc_of(aid):
        return meta.get(norm_id(aid), {})

    def section_meta(aid, number):
        want = norm_num(number)
        for s in doc_of(aid).get("sections", []):
            if norm_num(s.get("number")) == want:
                return s
        return {}

    matched = sum(1 for a in pick if doc_of(a))
    print(f"section metadata matched for {matched}/{len(pick)} appraisals")
    if matched < len(pick):
        print("  (unmatched appraisals get no title, role, page span or link)")

    rows, missing, empties = [], [], []
    shown, total = defaultdict(int), defaultdict(int)
    seen = defaultdict(list)
    for aid in pick:
        doc = doc_of(aid)
        for number, eps in sorted(by_aid[aid].items()):
            if not eps:
                continue
            sm = section_meta(aid, number)
            if len(eps) <= args.per_section:
                keep = list(eps)
            else:
                idx = sorted(random.sample(range(len(eps)),
                                           args.per_section))
                keep = [eps[i] for i in idx]
            shown[aid] += len(keep)
            total[aid] += len(eps)
            seen[aid].append(f"{number} {sm.get('title', '')}".strip())
            for e in keep:
                rows.append({
                    "Appraisal": aid,
                    "Drug": doc.get("drug", ""),
                    "Year": doc.get("year", ""),
                    "Section": f"{number} {sm.get('title', '')}".strip(),
                    "Role": sm.get("role", ""),
                    "Pages": f"{e.get('page', '')}-{e.get('page_end', '')}",
                    "Link": link_of(doc),
                    "Endpoint": e.get("endpoint_en") or e.get("endpoint", ""),
                    "Endpoint_FR": e.get("endpoint", ""),
                    "Category": e.get("category", ""),
                    "Study": e.get("study", ""),
                    "Comparator": e.get("comparator", ""),
                    "Result": e.get("result", ""),
                    "Direction": e.get("direction", ""),
                    "Primary": e.get("primary", ""),
                    "Endpoint_OK": "",
                    "Result_OK": "",
                    "Direction_OK": "",
                    "Page_OK": "",
                    "Notes": "",
                })
        missing.append({
            "Appraisal": aid,
            "Drug": doc.get("drug", ""),
            "Year": doc.get("year", ""),
            "Sections_to_read": "; ".join(seen[aid]),
            "Link": link_of(doc),
            "Extracted_total": total[aid],
            "Shown_on_sheet_1": shown[aid],
            "Endpoints_the_model_missed": "",
            "How_many_missed": "",
            "Notes": "",
        })

    for aid, number in pick_empty:
        doc = doc_of(aid)
        sm = section_meta(aid, number)
        empties.append({
            "Appraisal": aid,
            "Drug": doc.get("drug", ""),
            "Year": doc.get("year", ""),
            "Section": f"{number} {sm.get('title', '')}".strip(),
            "Role": sm.get("role", ""),
            "Pages": f"{sm.get('page', '')}-{sm.get('page_end', '')}",
            "Chars": sm.get("chars", ""),
            "Link": link_of(doc),
            "Model_said": "no endpoints",
            "Correct": "",
            "Endpoints_actually_there": "",
            "Notes": "",
        })

    out = args.out or next_free("has_endpoint_benchmark")
    with pd.ExcelWriter(out, engine="openpyxl") as xl:
        pd.DataFrame(GUIDE, columns=["Field", "Meaning"]).to_excel(
            xl, sheet_name="Instructions", index=False)
        pd.DataFrame(rows).to_excel(xl, sheet_name="Endpoints", index=False)
        pd.DataFrame(missing).to_excel(xl, sheet_name="Missing", index=False)
        pd.DataFrame(empties).to_excel(
            xl, sheet_name="Empty_sections", index=False)
        widths = {"Instructions": [34, 100],
                  "Endpoints": [12, 20, 7, 34, 18, 10, 46, 40, 34, 14,
                                22, 22, 34, 20, 9, 13, 11, 14, 9, 52],
                  "Missing": [12, 20, 7, 40, 46, 16, 17, 54, 15, 40],
                  "Empty_sections": [12, 20, 7, 34, 18, 10, 9, 46,
                                     14, 10, 54, 40]}
        for name, cols in widths.items():
            ws = xl.sheets[name]
            for i, w in enumerate(cols, start=1):
                ws.column_dimensions[
                    ws.cell(row=1, column=i).column_letter].width = w
            ws.freeze_panes = "A2"

    print(f"\n{len(rows)} endpoint rows across {len(pick)} appraisals")
    print(f"{len(missing)} sections to check for misses")
    print(f"{len(empties)} empty sections to confirm")
    print(f"-> {out}")
    print("\nSend the filled sheet back and the scorer gets written to")
    print("match how you actually filled it, not the other way round.")


if __name__ == "__main__":
    main()
