"""
Wire source_citations.py into app3.py.

Adds page-level source citation to the G-BA and HAS record views: every
indexed passage shown inline with a link that opens the official PDF scrolled
to the page the text came from.

Three edits:

  1. An optional-import block for source_citations, placed BEFORE the HTA body
     switcher. This matters: render_agency_view() is called at module level
     and followed by st.stop(), so anything imported further down the file
     never executes on a G-BA or HAS page. A module imported after that line
     raises NameError the moment a non-NICE record renders.

  2. attach() in render_agency_view(), handing the module the three sheets
     through the app's existing load_agency_sheet() — which already returns an
     empty frame for a missing sheet, so a short workbook degrades quietly.

  3. render_panel() in render_agency_record(), directly after the NICE
     attributes panel. Silent for NICE and for records with no passages.

The patcher also CHECKS (but does not touch) the import position of
nice_attributes, which has the same ordering hazard and its own patcher in
patch_app3_imports.py.

Idempotent. Backs up to app3.py.bak5. Compiles before writing.

  python patch_app3_citations.py --dry-run
  python patch_app3_citations.py
"""

import argparse
import os
import re
import shutil
import sys

MARK = "source_citations"

IMPORT_ANCHOR = 'DATA_FILE = '

IMPORT_BLOCK = '''# --- page-level source citation for G-BA / HAS (optional) ---------------
# Imported HERE, above the HTA body switcher, and not with the other optional
# modules further down: render_agency_view() runs at module level and is
# followed by st.stop(), so a name bound after that point does not exist when
# a G-BA or HAS record renders. A module whose name matches a data file or
# directory can also import as an empty namespace package, which succeeds and
# then fails on first use, so check for a known attribute rather than
# trusting the import.
try:
    import source_citations
    if not hasattr(source_citations, "render_panel"):
        raise ImportError("source_citations resolved to a namespace package")
except (ImportError, AttributeError):
    class _NoCitations:
        @staticmethod
        def attach(*a, **k):
            return None

        @staticmethod
        def is_attached():
            return False

        @staticmethod
        def has_sources(*a, **k):
            return False

        @staticmethod
        def coverage():
            return {}

        @staticmethod
        def render_panel(*a, **k):
            return False

    source_citations = _NoCitations()


'''

ATTACH_ANCHOR = "    doc_index = load_document_index(version)\n"

ATTACH_BLOCK = '''    # Index the three page-cited passage sheets for this run. Each is absent
    # from the 6-sheet build of the workbook, in which case load_agency_sheet
    # returns an empty frame and the citation panel renders nothing.
    source_citations.attach(
        load_agency_sheet("GBA_Decision_Passages", version),
        load_agency_sheet("HAS_Economic_Pages", version),
        load_agency_sheet("Economic_Findings", version),
    )
'''

RENDER_ANCHOR = '    nice_attributes.render_panel(row.get("appraisal_id"))\n'

RENDER_BLOCK = '''
    # Cited source passages, with a link that opens the official PDF at the
    # page the text came from. Silent for NICE and for any record with no
    # indexed passages, so no guard is needed here.
    source_citations.render_panel(agency, row.get("appraisal_id"))
'''


def _fail(msg):
    print(msg)
    sys.exit(1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--app", default="app3.py")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    if not os.path.exists(args.app):
        _fail(f"Not found: {args.app}")
    if not os.path.exists("source_citations.py"):
        print("warning: source_citations.py is not in this folder. The patch "
              "will apply, but the panel stays silent until you copy the "
              "module in beside app3.py.\n")

    with open(args.app, "rb") as f:
        raw = f.read()
    crlf = b"\r\n" in raw
    text = raw.decode("utf-8")
    if crlf:
        text = text.replace("\r\n", "\n")

    if MARK in text:
        print(f"Already applied — {args.app} already wires in source_citations.")
        _check_ordering(text)
        return

    # --- edit 1: the import block, before the switcher -------------------
    idx = text.find("\n" + IMPORT_ANCHOR)
    if idx == -1:
        _fail(f"Could not find the '{IMPORT_ANCHOR.strip()}' line in {args.app}.\n"
              "Send me the lines around it and I'll adjust the anchor.")
    idx += 1
    text = text[:idx] + IMPORT_BLOCK + text[idx:]

    # --- edit 2: attach() in render_agency_view --------------------------
    if text.count(ATTACH_ANCHOR) != 1:
        _fail(f"Expected exactly one '{ATTACH_ANCHOR.strip()}' line, found "
              f"{text.count(ATTACH_ANCHOR)}. Not patching.")
    text = text.replace(ATTACH_ANCHOR, ATTACH_ANCHOR + ATTACH_BLOCK, 1)

    # --- edit 3: render_panel in render_agency_record --------------------
    if text.count(RENDER_ANCHOR) != 1:
        _fail(f"Expected exactly one nice_attributes.render_panel line, found "
              f"{text.count(RENDER_ANCHOR)}. Not patching.")
    text = text.replace(RENDER_ANCHOR, RENDER_ANCHOR + RENDER_BLOCK, 1)

    # compile() rather than ast.parse(): only compile catches a name used
    # prior to its global declaration, which ast accepts happily.
    try:
        compile(text, args.app, "exec")
    except SyntaxError as e:
        _fail(f"Patched source does not compile ({e}). {args.app} is untouched.")

    if args.dry_run:
        print("--dry-run: all three edits apply and the result compiles. "
              "Nothing written.")
        _check_ordering(text)
        return

    shutil.copy2(args.app, args.app + ".bak5")
    out = text.replace("\n", "\r\n") if crlf else text
    with open(args.app, "wb") as f:
        f.write(out.encode("utf-8"))
    print(f"Wrote {args.app}  (backup: {args.app}.bak5)")
    print("  + source_citations import, above the HTA body switcher")
    print("  + attach() for the three passage sheets in render_agency_view")
    print("  + render_panel() in render_agency_record")
    _check_ordering(text)
    print("\nCopy source_citations.py in beside app3.py, then restart Streamlit.")


def _check_ordering(text):
    """
    Warn about the SAME hazard for nice_attributes: a name bound after the
    'if HTA_BODY != "NICE"' block does not exist when a G-BA or HAS record
    renders, because st.stop() ends the run before the import line.
    """
    lines = text.split("\n")
    def line_of(pattern):
        for i, line in enumerate(lines, 1):
            if re.match(pattern, line):
                return i
        return None

    switcher = line_of(r'\s*HTA_BODY\s*=\s*render_agency_switcher\(\)')
    attrs = line_of(r'\s*import nice_attributes\b')
    if switcher and attrs and attrs > switcher:
        print(f"\n  ⚠️  nice_attributes is imported at line {attrs}, AFTER the HTA "
              f"body switcher at line {switcher}.")
        print("      A G-BA or HAS record will raise NameError on that module "
              "before it reaches the citation panel.")
        print("      Run patch_app3_imports.py to move it, then re-check.")
    elif switcher and attrs:
        print(f"\n  ✓ nice_attributes imports at line {attrs}, before the "
              f"switcher at line {switcher}.")


if __name__ == "__main__":
    main()
