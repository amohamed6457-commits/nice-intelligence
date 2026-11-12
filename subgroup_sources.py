import argparse
import os
import re
import sys

WORKBOOKS = ["NICE_HAS_GBA_updated_2026-11-07.xlsx",
             "NICE_v14_updated_2026-09-29.xlsx"]

# Both the German resolutions and the English courtesy translations.
CATEGORY_PATTERNS = [
    ("Mortality",        r"Mortalit[äa]t|\bMortality\b"),
    ("Morbidity",        r"Morbidit[äa]t|\bMorbidity\b"),
    ("Quality of life",  r"[Gg]esundheitsbezogene\s+Lebensqualit[äa]t"
                         r"(?:\s+of\s+life)?|Health-?related\s+quality"
                         r"(?:\s+of\s+life)?|Lebensqualit[äa]t"),
    ("Adverse events",   r"Nebenwirkungen|Side\s+effects|Adverse\s+events"),
]

# The direction glyph carries the verdict; the words confirm it.
SYMBOL_VERDICT = {
    "∅": "NO DATA",        # empty set: no data available
    "↑": "ADVANTAGE",
    "↑↑": "ADVANTAGE",
    "↓": "DISADVANTAGE",
    "↓↓": "DISADVANTAGE",
    "↔": "NO DIFFERENCE",
    "n.a.": "NO DATA",          # not assessable
}

TABLE_ANCHOR = re.compile(
    r"Endpunktkategorie|Endpoint\s+category", re.IGNORECASE)
# Everything after this is the legend, not data.
LEGEND = re.compile(r"Erl[äa]uterung|Explanation|Begr[üu]ndung:", re.I)


def find_workbook(explicit=None):
    if explicit:
        if os.path.exists(explicit):
            return explicit
        sys.exit(f"Workbook not found: {explicit}")
    for w in WORKBOOKS:
        if os.path.exists(w):
            return w
    sys.exit("No workbook found. Pass --workbook.")


def _tidy(text):
    """Undo the two-column wrap that splits a heading across lines, so
    'Health-related quality ... of life' reads as one label."""
    t = " ".join(str(text).split())
    t = re.sub(r"(Health-?related\s+quality)\s+(.*?)\s+of\s+life", r"\1 of life \2", t)
    t = re.sub(r"(Gesundheitsbezogene)\s+(.*?)\s+Lebensqualit", r"\1 Lebensqualit \2", t)
    return t


def parse_table(text):
    """{category: (symbol, summary)} from one passage, or {} if no table."""
    t = _tidy(text)
    anchor = TABLE_ANCHOR.search(t)
    if not anchor:
        return {}
    body = t[anchor.end():]
    stop = LEGEND.search(body)
    if stop:
        body = body[:stop.start()]
    if not body.strip():
        return {}

    # Where each category label starts, in reading order.
    marks = []
    for name, pattern in CATEGORY_PATTERNS:
        m = re.search(pattern, body)
        if m:
            marks.append((m.start(), m.end(), name))
    marks.sort()

    out = {}
    for i, (_start, end, name) in enumerate(marks):
        chunk = body[end:marks[i + 1][0] if i + 1 < len(marks) else len(body)]
        chunk = chunk.strip(" :–-")
        sym = re.match(r"(↑↑|↓↓|↑|↓|↔|∅|n\.\s?a\.)",
                       chunk)
        symbol = None
        if sym:
            symbol = sym.group(1).replace(" ", "")
            chunk = chunk[sym.end():].strip(" :–-")
        out[name] = (symbol, " ".join(chunk.split())[:400])

    # Reject a numeric endpoint table that merely sits near the anchor. The
    # real summary table states a verdict per category in words, with a
    # direction glyph; a results table is mostly digits and patient counts.
    if len(out) < 3:
        return {}
    glyphs = sum(1 for sym, _ in out.values() if sym)
    digits = sum(sum(c.isdigit() for c in text) / max(len(text), 1)
                 for _sym, text in out.values()) / len(out)
    if glyphs < 2 and digits > 0.08:
        return {}
    return out


def verdict_from_table(symbol, summary):
    """The verdict the resolution states, from its own glyph and wording."""
    text = str(summary or "")
    if symbol in SYMBOL_VERDICT:
        base = SYMBOL_VERDICT[symbol]
        # "<-> No deaths occurred" is not the same claim as "no difference".
        if base == "NO DIFFERENCE" and re.search(
                r"[Nn]o deaths (occurred|were)|[Tt]here were no deaths|"
                r"keine\s+Todesf[äa]lle|traten\s+keine", text):
            return "NO EVENTS"
        return base
    if re.search(r"[Nn]o assessable data|nicht\s+bewertbar|"
                 r"keine\s+bewertbaren\s+Daten", text):
        return "NO DATA"
    if re.search(r"[Nn]o data (are |is )?available|liegen\s+keine\s+Daten", text):
        return "NO DATA"
    if re.search(r"[Nn]o relevant difference|kein[e]?\s+.{0,30}relevante", text):
        return "NO DIFFERENCE"
    if re.search(r"\bAdvantage|\bVorteil", text):
        return "ADVANTAGE"
    if re.search(r"\bDisadvantage|\bNachteil", text):
        return "DISADVANTAGE"
    return None


def build_index(workbook):
    """appraisal_id -> {page, url, table} for every locatable summary table."""
    import pandas as pd
    pas = pd.read_excel(workbook, sheet_name="GBA_Decision_Passages")
    pas["appraisal_id"] = pas["appraisal_id"].astype(str).str.strip()
    index = {}
    for r in pas.itertuples():
        text = str(getattr(r, "source_text", "") or "")
        if not TABLE_ANCHOR.search(text):
            continue
        table = parse_table(text)
        if not table:
            continue
        page = getattr(r, "document_page", None)
        prev = index.get(r.appraisal_id)
        # Prefer the table that resolved the most categories.
        if prev and len(prev["table"]) >= len(table):
            continue
        index[r.appraisal_id] = {
            "page": int(page) if page == page and page is not None else None,
            "url": getattr(r, "source_url", None),
            "table": table,
        }
    return index


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook")
    ap.add_argument("--appraisal")
    ap.add_argument("--disagreements", type=int, default=0)
    args = ap.parse_args()

    import pandas as pd
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from endpoint_summary import verdict as ais_verdict, CATEGORIES

    wb = find_workbook(args.workbook)
    print(f"workbook: {wb}")
    index = build_index(wb)
    subs = pd.read_excel(wb, sheet_name="GBA_Subgroups")
    subs["appraisal_id"] = subs["appraisal_id"].astype(str).str.strip()
    print(f"appraisals with a locatable endpoint table: {len(index):,}\n")

    if args.appraisal:
        got = index.get(args.appraisal.strip())
        if not got:
            sys.exit(f"No endpoint-category table found for {args.appraisal}.")
        print(f"=== {args.appraisal} === page {got['page']}")
        print(f"    {str(got['url'])[:100]}\n")
        rows = subs[subs["appraisal_id"] == args.appraisal.strip()]
        for name, (symbol, summary) in got["table"].items():
            from_pdf = verdict_from_table(symbol, summary)
            field = {c[1]: c[0] for c in CATEGORIES}.get(name)
            ais = None
            if field and len(rows):
                ais = ais_verdict(rows.iloc[0][field])
            agree = "agree" if ais == from_pdf else f"DIFFER (AIS: {ais})"
            print(f"  {name:<17} {str(symbol):<3} {str(from_pdf):<14} {agree}")
            print(f"       \"{summary[:95]}\"")
        return

    # An appraisal with several subgroups prints a summary table PER
    # SUBGROUP. Comparing one parsed table against every subgroup row would
    # manufacture disagreements, so agreement is measured only where the
    # mapping is unambiguous: appraisals with exactly one subgroup.
    counts = subs.groupby("appraisal_id").size()
    single = set(counts[counts == 1].index)
    scored = subs[subs["appraisal_id"].isin(single)]
    print(f"single-subgroup appraisals: {len(single):,} of {len(counts):,} "
          f"- agreement is measured on these only,\nbecause a multi-subgroup "
          f"resolution prints one summary table per subgroup.\n")

    agree = differ = missing = 0
    examples = []
    for r in scored.itertuples():
        got = index.get(r.appraisal_id)
        if not got:
            continue
        for field, name in CATEGORIES:
            if name not in got["table"]:
                continue
            symbol, summary = got["table"][name]
            pdf = verdict_from_table(symbol, summary)
            ais = ais_verdict(getattr(r, field, None))
            if pdf is None or ais is None:
                missing += 1
            elif pdf == ais:
                agree += 1
            else:
                differ += 1
                if len(examples) < args.disagreements:
                    examples.append((r.appraisal_id, name, ais, pdf,
                                     summary[:70]))
    total = agree + differ
    print("AIS verdict vs the resolution's own summary table")
    print(f"  agree        : {agree:,}")
    print(f"  differ       : {differ:,}")
    print(f"  not scorable : {missing:,}")
    if total:
        print(f"\n  agreement: {agree / total:.1%}  (n={total:,})")
        print("\nThis is the first external check on the decision data: the "
              "AIS record\nagainst the resolution PDF, which are produced "
              "separately.")
    for aid, name, ais, pdf, summary in examples:
        print(f"\n  {aid} {name}\n    AIS: {ais}   table: {pdf}\n    \"{summary}\"")


if __name__ == "__main__":
    main()
