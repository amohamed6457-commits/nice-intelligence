"""
Pilot: recover the tables in a G-BA resolution as real tables.

Why this exists
---------------
The passages in GBA_Decision_Passages kept the documents' LINE order but lost
their COLUMN positions. One logical row of the KEYNOTE-091 endpoints table
arrives as three separate lines:

    Overall survival 506 n.r. 504 n.r. 0.79
    113 (22.3) 138 (27.4) [0.62; 1.01]
    0.064

Rebuilding a table from that means inferring which number belongs to which
column. For a tool whose value is traceability, a hazard ratio silently
placed in the wrong column is worse than no table at all — so this does not
attempt it. It goes back to the PDF, where the column geometry still exists.

What the first pilot run taught us (D-1058, 11 pages)
----------------------------------------------------
1. RULED BORDERS WORK. Page 8's adverse-event table came out perfectly
   columned: 'Endocrine disorders | 496 | 10 (2.0) | 499 | 0 (0) | 21.13
   [1.24; 359.55] 0.002'.

2. THE TEXT-ALIGNMENT FALLBACK PRODUCED NONSENSE, and has been removed.
   On pages 2 and 9, which are prose, it sliced sentences into fake columns:
   'New therapeutic indicati | on (according to | the marketing authorisat'.
   A page with no ruled table should report no table, not invent one.

3. THE CITED PAGES ARE THE WRONG PAGES. The workbook records the page a
   SECTION STARTS on (2, 8, 9 for this resolution), but the endpoints tables
   run across pages 3-7. So this now scans every page by default.

Even with ruling lines, a bordered text box is not a data table, so each
candidate has to earn its place — see looks_like_data().

    pip install pdfplumber requests
    python pilot_gba_tables.py --appraisal D-1058 --csv
    python pilot_gba_tables.py --appraisal D-1058 --page 4 --keep-all
"""

import argparse
import os
import re
import statistics
import sys

WORKBOOK_CANDIDATES = [
    "NICE_v14_updated_2026-09-29.xlsx",
    "NICE_v14_updated_2026-09-28.xlsx",
]
CACHE = "gba_pdfs"

# Ruled borders only. The text strategy is deliberately absent: see note 2.
SETTINGS = {"vertical_strategy": "lines", "horizontal_strategy": "lines"}

# A lone decimal fragment: "1.", "12.", ".6" — half a number left by a
# PDF whose text layer positions each glyph separately.
_FRAGMENT = re.compile(r"^\d{0,3}\.$|^\.\d+$")

_NUMERIC = re.compile(r"^[<>≥≤~±]?\s*-?[\d.,]+\s*%?$|^\d+\s*\([\d.,]+\)$|"
                      r"^\[[-\d.,;\s]+\]$|^n\.?[rda]\.?$|^-+$")


def clean(cell):
    return "" if cell is None else " ".join(str(cell).split())


def _numeric_count(row):
    return sum(1 for c in row if c and _NUMERIC.match(c))


def tidy(rows):
    """
    Repair the four things the pilot run showed pdfplumber doing to these
    resolutions. Every one is a layout artefact, not a data problem, so each
    repair moves cells around and never invents or reorders a value.

      1. PHANTOM COLUMNS. Page 4's 6-column table came out 12 columns wide,
         every other one empty, because the ruling detector sees the cell
         padding as a border. Columns that are empty in every row are
         dropped.

      2. SPLIT HEADERS. A multi-line column title ('N' / 'Median time to' /
         'event in months' / '[95% CI]') becomes four separate rows. The
         leading rows, up to the first row carrying real data, are joined
         per column into one header.

      3. LABEL/DATA SPLIT. On page 7 the row label sits alone on one row and
         its numbers on the next ('Serious adverse events (SAE)' then '496 |
         127 (25.6) | ...'). A label-only row immediately followed by a row
         with an empty first cell is merged into one.

      4. FOOTNOTES. Page 8's table was REJECTED as prose because the
         footnote block sits in a 2,820-character cell at the bottom. A
         trailing row whose only content is one very long cell is lifted out
         and returned separately, so it no longer poisons the prose test.

    Returns (rows, notes).
    """
    rows = [list(r) for r in rows if any(r)]
    if not rows:
        return [], []

    # 4 — pull trailing footnote rows out before anything else measures them.
    notes = []
    while rows:
        filled = [c for c in rows[-1] if c]
        if len(filled) == 1 and len(filled[0]) > 150:
            notes.insert(0, filled[0])
            rows.pop()
        else:
            break
    if not rows:
        return [], notes

    # 1 — drop columns empty in every row.
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    keep = [i for i in range(width) if any(r[i] for r in rows)]
    rows = [[r[i] for i in keep] for r in rows]
    if not rows or not rows[0]:
        return [], notes

    # 2 — join the leading title rows into one header.
    #
    # The header ends at the first row that is either real data (two or more
    # numeric cells) or a SECTION HEADING — a row with text in the first
    # column and nothing anywhere else, like 'Mortality' or 'Total adverse
    # events (presented additionally)'. Without the second test the heading
    # was being merged into the header ('Endpoint Mortality') and, worse, on
    # page 7 it stole the label off the row beneath it, leaving a line of
    # numbers with no name.
    def _is_heading(r):
        return bool(r[0]) and not any(r[1:])

    first_stop = next((i for i, r in enumerate(rows)
                       if _numeric_count(r) >= 2 or _is_heading(r)), None)
    if first_stop and first_stop > 1:
        head = rows[:first_stop]
        merged = [" ".join(p for p in (r[c] for r in head) if p).strip()
                  for c in range(len(rows[0]))]
        rows = [merged] + rows[first_stop:]

    # 3 — merge a label-only row into the data row that follows it.
    out = []
    i = 0
    while i < len(rows):
        row = rows[i]
        nxt = rows[i + 1] if i + 1 < len(rows) else None
        if (nxt is not None and row[0] and not any(row[1:])
                and not nxt[0] and _numeric_count(nxt) >= 2):
            out.append([row[0]] + list(nxt[1:]))
            i += 2
            continue
        out.append(row)
        i += 1
    return out, notes


def merge_split_columns(rows):
    """
    Rejoin a logical column that was cut in two.

    The commonest failure across all three pilot resolutions is not damage,
    it is a clean split. D-1313's overall-survival table came out like this:

        col 3   header (empty)                data '28,7 [21,98; n. b.]'
        col 4   header 'Mediane Überlebens…'  data (empty)

    Every value sits one column left of its own heading, and the two columns
    are mutually exclusive — one carries the heading and no data, its
    neighbour carries data and no heading. Merging the pair restores the
    original column with nothing moved anywhere ambiguous. The same shape
    appears in D-1058's English mortality table and in D-935's clinical
    tables, which is why so few survived.

    Only an unambiguous pair is merged: one side must be header-only across
    EVERY row, the other data-only. Anything else is left alone for
    alignment_ok() to judge, because a column that holds both a heading and
    data is a real column and must not be absorbed into its neighbour.

    A trailing heading with no data and no partner — 'Imlunestrant vs.
    Fulvestrant/Exemestan', the spanning title over the comparison columns —
    has its text folded into the previous heading and the empty column
    dropped. No data moves.
    """
    if len(rows) < 2:
        return rows
    header, body = rows[0], rows[1:]
    if _numeric_count(header) >= 2 or (header[0] and not any(header[1:])):
        return rows                      # no header row to align against

    width = len(header)

    def head_only(c):
        return bool(header[c]) and not any(c < len(r) and r[c] for r in body)

    def data_only(c):
        return not header[c] and any(c < len(r) and r[c] for r in body)

    merged_into = {}
    c = 0
    while c < width - 1:
        if data_only(c) and head_only(c + 1):
            merged_into[c] = c + 1       # heading sits to the right
            c += 2
        elif head_only(c) and data_only(c + 1):
            merged_into[c + 1] = c       # heading sits to the left
            c += 2
        else:
            c += 1
    if not merged_into:
        return rows

    out_header, out_body, used = [], [[] for _ in body], set()
    for c in range(width):
        if c in used:
            continue
        partner = merged_into.get(c)
        if partner is not None and partner not in used:
            used.update({c, partner})
            out_header.append(header[c] or header[partner])
            for i, r in enumerate(body):
                a = r[c] if c < len(r) else ""
                b = r[partner] if partner < len(r) else ""
                out_body[i].append(a or b)
        else:
            used.add(c)
            out_header.append(header[c])
            for i, r in enumerate(body):
                out_body[i].append(r[c] if c < len(r) else "")

    # Fold a leftover heading-only column into the heading before it.
    i = len(out_header) - 1
    while i > 0:
        if out_header[i] and not any(r[i] for r in out_body):
            out_header[i - 1] = (out_header[i - 1] + " " + out_header[i]).strip()
            for r in out_body:
                del r[i]
            del out_header[i]
        i -= 1
    return [out_header] + out_body


def alignment_ok(rows):
    """
    Do the data cells actually sit under the headings above them?

    Page 4 of D-1058 comes out of pdfplumber like this:

        header | Placebo N | (blank)         | Median time to event |
        data   | 504       | n.r. 138 (27.4) | (blank)              |

    The placebo median is one column left of its heading. Everything is
    present and nothing is corrupted, but read as a table it says the wrong
    thing — and a hazard ratio under the wrong heading is precisely the
    failure this whole approach exists to avoid. No tidying can fix it: the
    column boundaries were already wrong when the cells were found.

    So it has to be detected and the table refused, falling back to the
    passage text and the page link, which are never wrong about anything.

    The signal is a column that is blank in the header but carries data, or
    the reverse. Tables with no header row at all (a continuation from the
    previous page) are not checked, as there is nothing to misalign against.
    """
    if len(rows) < 2:
        return True, ""
    header, body = rows[0], rows[1:]
    if _numeric_count(header) >= 2:
        return True, "no header row to check"
    # A SECTION HEADING is not a header. Page 4's recurrences table opens
    # with 'Recurrences' alone, and page 6's with 'EORTC QLQ-C30 h'. Reading
    # those as the header makes every populated column look unmatched, and
    # the first version of this check threw away two tables that were
    # verifiably correct. A table whose top row is a heading simply has no
    # header to align against — the same case as a continuation table.
    if header[0] and not any(header[1:]):
        return True, "section heading, no header row to check"

    width = len(header)
    bad = []
    for c in range(width):
        head_filled = bool(header[c]) if c < len(header) else False
        data_filled = any(c < len(r) and r[c] for r in body)
        if head_filled == data_filled:
            continue
        fill = (sum(1 for r in body if c < len(r) and r[c]) / len(body)
                if body else 0)
        # A SECTION-HEADING COLUMN carries no header and is blank in almost
        # every row, holding only the labels that span the table —
        # 'Morbidität', 'Nebenwirkungen', 'Symptomatik (EORTC QLQ-MY20)'.
        # A genuinely shifted column is the opposite: blank header, but a
        # value on nearly every row. So the fill rate separates them, where
        # the mere presence of data could not.
        if not head_filled and fill < 0.5:
            continue
        # A SECOND LABEL COLUMN is densely filled but holds row descriptors,
        # not data: 'Nombre d'événements, n (%)', 'Médiane de SG, mois',
        # 'IC 95%' down the side of a HAS survival table, with the arms
        # across the top. It has no header of its own because the heading
        # above it names the endpoint ('SG'). Density alone reads it as a
        # shifted data column; what separates them is that a shifted column
        # carries VALUES. 'n.r.' and '-' count as values, so a genuine shift
        # is still caught.
        if not head_filled and not any(
                _NUMERIC.match(r[c]) for r in body if c < len(r) and r[c]):
            continue
        # A SPANNING TITLE sits over several columns with no data of its own
        # ('Intervention vs. Kontrolle' above the HR and p-value columns).
        # merge_split_columns folds most of these into their neighbour; one
        # that survives is harmless, and refusing the table over it costs
        # every endpoint in it.
        if head_filled and fill == 0:
            continue
        # A ROW-LABEL COLUMN under a blank header is normal, not a shift.
        # Every G-BA endpoint table is laid out this way:
        #
        #   header | (blank)          | Endpunktkategorie | Ruxolitinib N | …
        #   data   | Gesamtmortalität | (blank)           | 110           | …
        #
        # The first column names the endpoint and has no heading of its own,
        # because the heading sits in the next column over the category.
        # merge_split_columns cannot join them, since that second column is
        # not header-only — it carries 'Morbidität', 'Nebenwirkungen' and
        # the other section headings on its own rows. Reading column 1 as a
        # misalignment threw away three perfectly extracted tables from
        # D-161, 25 endpoints in total, and the same shape across the older
        # German resolutions. A real shift displaces data in the BODY of the
        # table, which the checks below still catch.
        if c == 0 and not head_filled and data_filled:
            continue
        bad.append(c + 1)
    if bad:
        return False, ("columns " + ", ".join(map(str, bad)) +
                       " do not line up with their headings")
    return True, ""


def looks_like_data(rows):
    """
    Is this a data table, or a bordered box of prose?

    Page 2 of D-1058 is a bordered box containing the therapeutic indication
    in sentences. It has ruling lines, so pdfplumber finds it, but it is not
    a table anyone wants rendered as one. Three tests separate them:

      - at least 2 columns and 2 rows
      - cells are short; prose boxes have long ones
      - some cell somewhere is a number, a percentage, a count like
        '113 (22.3)' or an interval like '[0.62; 1.01]'

    Returns (ok, reason) so a rejection can be printed rather than silent.
    """
    if len(rows) < 2:
        return False, "fewer than 2 rows"
    width = max(len(r) for r in rows)
    if width < 2:
        return False, "single column"

    cells = [c for r in rows for c in r if c]
    if not cells:
        return False, "empty"
    # Measure the BODY, not the header: a column title can legitimately run
    # long ('Hazard ratio [95% CI] p value Absolute difference (AD)') without
    # making the table prose. Footnote rows are already removed by tidy().
    body = [c for r in rows[1:] for c in r if c] or cells
    longest = max(len(c) for c in body)
    median = statistics.median(len(c) for c in body)
    if median > 45 or longest > 300:
        return False, f"cells read as prose (median {int(median)}, longest {longest})"

    numeric = sum(1 for c in cells if _NUMERIC.match(c))
    if numeric < 2:
        return False, "no numeric cells"

    # A FRAGMENTED TABLE is the dangerous case, because it looks numeric.
    # Some PDFs position every glyph separately, so cell detection chops
    # numbers apart: '31.4' arrives as '3' | '1.' | '4', and '12.6' as
    # '12' | '.6'. Each piece counts as a numeric cell, so the checks above
    # pass and the table renders as data that means nothing. Found in HAS
    # committee opinions, where it would have shipped at scale.
    #
    # Two signals separate it from a real table, measured on the same
    # examples: a real one is 71-100% filled with cells like '0,17 [0,02;
    # 1,37]' and '66 (60)'; a fragmented one is a quarter filled and EVERY
    # cell is one or two characters.
    filled = len(cells)
    total_cells = sum(len(r) for r in rows)
    fill = filled / total_cells if total_cells else 0
    tiny = sum(1 for c in cells if len(c) <= 2) / filled
    pieces = sum(1 for c in cells if _FRAGMENT.match(c)) / filled
    if pieces >= 0.20:
        return False, (f"{pieces:.0%} of cells are split decimals "
                       "('1.', '.6') — the text layer is fragmented")
    if fill < 0.35 and tiny >= 0.6:
        return False, (f"only {fill:.0%} of cells filled and {tiny:.0%} are "
                       "1-2 characters — scattered text, not a table")
    return True, f"{len(rows)}x{width}, {numeric} numeric cells"


def find_workbook(explicit=None):
    if explicit:
        if os.path.exists(explicit):
            return explicit
        sys.exit(f"Workbook not found: {explicit}")
    for name in WORKBOOK_CANDIDATES:
        if os.path.exists(name):
            return name
    here = [f for f in os.listdir(".") if f.lower().endswith(".xlsx")]
    sys.exit("No workbook found. Pass --workbook. Here: " + (", ".join(here) or "none"))


def source_rows(workbook, appraisal):
    import pandas as pd
    passages = pd.read_excel(workbook, sheet_name="GBA_Decision_Passages")
    rows = passages[passages["appraisal_id"].astype(str).str.strip() == appraisal]
    if not len(rows):
        ids = passages["appraisal_id"].astype(str).unique()[:5]
        sys.exit(f"No passages for {appraisal}. Examples: {', '.join(ids)}")
    return rows


def fetch(url, appraisal):
    os.makedirs(CACHE, exist_ok=True)
    path = os.path.join(CACHE, re.sub(r"[^A-Za-z0-9_.-]", "_", f"{appraisal}.pdf"))
    if os.path.exists(path) and os.path.getsize(path) > 1000:
        print(f"  cached: {path}")
        return path
    try:
        import requests
    except ImportError:
        sys.exit("requests is not installed: pip install requests")
    print(f"  downloading {url}")
    resp = requests.get(url, timeout=120,
                        headers={"User-Agent": "hta-intelligence-pilot"})
    resp.raise_for_status()
    if not resp.content.startswith(b"%PDF"):
        sys.exit(f"  that URL did not return a PDF ({len(resp.content)} bytes)")
    with open(path, "wb") as fh:
        fh.write(resp.content)
    print(f"  saved {len(resp.content):,} bytes to {path}")
    return path


def show(rows, index, page_no, to_csv, stem, note):
    width = max(len(r) for r in rows)
    print(f"\n  ── page {page_no}, table {index} ── {note}")
    for r in rows[:12]:
        cells = [(c[:24] + "…") if len(c) > 25 else c for c in r]
        cells += [""] * (width - len(cells))
        print("    | " + " | ".join(f"{c:<25}" for c in cells))
    if len(rows) > 12:
        print(f"    … {len(rows) - 12} more rows")
    if to_csv:
        import csv
        out = f"{stem}_p{page_no}_t{index}.csv"
        with open(out, "w", newline="", encoding="utf-8-sig") as fh:
            csv.writer(fh).writerows(rows)
        print(f"    -> {out}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--appraisal", required=True, help="e.g. D-1058")
    ap.add_argument("--page", type=int, action="append",
                    help="limit to these pages; default is every page")
    ap.add_argument("--workbook")
    ap.add_argument("--csv", action="store_true")
    ap.add_argument("--keep-all", action="store_true",
                    help="also show candidates the data-table test rejected")
    args = ap.parse_args()

    try:
        import pdfplumber
    except ImportError:
        sys.exit("pdfplumber is not installed: pip install pdfplumber")

    workbook = find_workbook(args.workbook)
    print(f"workbook: {workbook}")
    rows = source_rows(workbook, args.appraisal)
    urls = [u for u in rows["source_url"].dropna().unique() if str(u).strip()]
    if not urls:
        sys.exit(f"No source_url recorded for {args.appraisal}.")
    path = fetch(str(urls[0]), args.appraisal)

    stem = os.path.splitext(path)[0]
    kept = rejected = 0
    with pdfplumber.open(path) as pdf:
        pages = args.page or range(1, len(pdf.pages) + 1)
        print(f"PDF has {len(pdf.pages)} pages; scanning "
              f"{'all' if not args.page else ', '.join(map(str, args.page))}\n")
        for page_no in pages:
            if page_no < 1 or page_no > len(pdf.pages):
                print(f"page {page_no}: out of range")
                continue
            try:
                tables = pdf.pages[page_no - 1].extract_tables(SETTINGS)
            except Exception as exc:                   # noqa: BLE001
                print(f"page {page_no}: {type(exc).__name__}: {exc}")
                continue
            for i, raw in enumerate(tables, 1):
                body = [[clean(c) for c in r] for r in raw]
                body = [r for r in body if any(r)]
                if not body:
                    continue
                body, notes = tidy(body)
                if not body:
                    continue
                body = merge_split_columns(body)
                ok, reason = looks_like_data(body)
                if ok:
                    aligned, why = alignment_ok(body)
                    if not aligned:
                        ok, reason = False, why
                if ok:
                    kept += 1
                    show(body, i, page_no, args.csv, stem, reason)
                else:
                    rejected += 1
                    if args.keep_all:
                        show(body, i, page_no, False, stem, f"REJECTED: {reason}")
                    else:
                        print(f"  page {page_no}, table {i}: skipped — {reason}")

    print("\n" + "-" * 62)
    print(f"{kept} data table(s) kept, {rejected} candidate(s) rejected.")
    if kept:
        print("Check the HEADER rows against the real pages — multi-line column\n"
              "titles are where this breaks. The data rows are usually fine.")
    else:
        print("Nothing survived. Re-run with --keep-all to see what was found\n"
              "and rejected, and send me that output.")


if __name__ == "__main__":
    main()
