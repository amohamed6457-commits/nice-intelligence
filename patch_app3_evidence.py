"""
Add the "Evidence available" filter to the Explorer, for G-BA and HAS.

Five small edits to app3.py, each anchored on a single line that occurs
exactly once:

  1. a guarded import of evidence_filter;
  2. "evidence" added to EXPLORER_FIELDS so the form collects it;
  3. the attach() calls, beside the other frames;
  4. an st.pills chip row under Procedure type;
  5. the filter applied in run_agency_search.

    python patch_app3_evidence.py --check
    python patch_app3_evidence.py
    python patch_app3_evidence.py --revert
"""

import argparse
import os
import shutil
import sys

TARGET = "app3.py"
BACKUP = "app3.py.bak_evidence"
SENTINEL = "evidence_filter.ids_for("

HUNKS = [
    (
        "    source_citations = _NoCitations()",
        '''    source_citations = _NoCitations()


# --- filter by whether anything was assessable (optional) ---------------
try:
    import evidence_filter
    if not hasattr(evidence_filter, "ids_for"):
        raise ImportError("evidence_filter resolved to a namespace package")
except (ImportError, AttributeError):
    class _NoEvidenceFilter:
        @staticmethod
        def attach_gba(*a, **k):
            return None

        @staticmethod
        def attach_has(*a, **k):
            return None

        @staticmethod
        def options_for(*a, **k):
            return []

        @staticmethod
        def counts_for(*a, **k):
            return {}

        @staticmethod
        def ids_for(*a, **k):
            return None

        @staticmethod
        def help_for(*a, **k):
            return ""

    evidence_filter = _NoEvidenceFilter()''',
        "guarded import",
    ),
    (
        'EXPLORER_FIELDS = ("drug", "condition", "areas", "outcomes", '
        '"procedures", "years", "orphan")',
        'EXPLORER_FIELDS = ("drug", "condition", "areas", "outcomes", '
        '"procedures", "years", "orphan", "evidence", "evidence_dir")',
        "EXPLORER_FIELDS",
    ),
    (
        '    docs = doc_index[doc_index["Institution"] == agency]',
        '''    # Which appraisals had assessable evidence at all. G-BA reads this
    # from the subgroup verdicts; HAS from whether an SMR grade exists.
    evidence_filter.attach_gba(load_agency_sheet("GBA_Subgroups", version)
                               if agency == "G-BA" else None)
    evidence_filter.attach_has(load_agency_sheet("HAS_Ratings", version)
                               if agency == "HAS" else None)
    docs = doc_index[doc_index["Institution"] == agency]''',
        "attach() calls",
    ),
    (
        '            key=_explorer_key(slug, "procedures"),',
        '''            key=_explorer_key(slug, "procedures"),''',
        "procedures key (anchor only)",
    ),
    (
        '        c3, c4 = st.columns([3, 2])',
        '''        _evidence_options = evidence_filter.options_for(agency)
        if _evidence_options:
            _dir_key = _explorer_key(slug, "evidence_dir")
            st.session_state.setdefault(_dir_key,
                                        evidence_filter.DIRECTIONS[0])
            _dir = st.session_state.get(_dir_key)
            st.pills(
                evidence_filter.label_for(agency),
                _evidence_options, selection_mode="multi",
                format_func=_counted(
                    evidence_filter.counts_for(agency, _dir)),
                key=_explorer_key(slug, "evidence"),
                help=evidence_filter.help_for(agency))
            if agency == "G-BA":
                st.radio(evidence_filter.direction_label(),
                         evidence_filter.DIRECTIONS, horizontal=True,
                         key=_dir_key)
        c3, c4 = st.columns([3, 2])''',
        "chip row",
    ),
    (
        '    if query.get("procedures"):\n'
        '        result = result[result["procedure_type"].isin(query["procedures"])]',
        '''    if query.get("procedures"):
        result = result[result["procedure_type"].isin(query["procedures"])]
    _evidence_ids = evidence_filter.ids_for(agency, query.get("evidence"),
                                            query.get("evidence_dir"))
    if _evidence_ids is not None:
        result = result[result["appraisal_id"].astype(str).str.strip()
                        .isin(_evidence_ids)]''',
        "filter application",
    ),
    (
        '    for field, label in (("areas", "areas"), ("outcomes", "outcomes"),',
        '    for field, label in (("areas", "areas"), ("outcomes", "outcomes"),\n'
        '                         ("evidence", "evidence"),',
        "search summary",
    ),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", default=TARGET)
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--revert", action="store_true")
    args = ap.parse_args()

    if args.revert:
        if not os.path.exists(BACKUP):
            sys.exit(f"No backup at {BACKUP}.")
        shutil.copy(BACKUP, args.file)
        print(f"Restored {args.file} from {BACKUP}.")
        return

    if not os.path.exists(args.file):
        sys.exit(f"{args.file} not found. Run this from the folder holding it.")
    if not os.path.exists("evidence_filter.py"):
        sys.exit("evidence_filter.py must sit beside app3.py before patching.")

    with open(args.file, encoding="utf-8", newline="") as fh:
        src = fh.read()
    crlf = src.count("\r\n") > 100
    print(f"{args.file}: {len(src):,} bytes, {'CRLF' if crlf else 'LF'}")

    if SENTINEL in src:
        print("Already patched - nothing to do.")
        return

    patched = src
    for find, replace, label in HUNKS:
        f = find.replace("\n", "\r\n") if crlf else find
        r = replace.replace("\n", "\r\n") if crlf else replace
        n = patched.count(f)
        if n != 1:
            sys.exit(f"Expected '{label}' exactly once, found {n}.\n"
                     "Send me your app3.py and I will redo the patch "
                     "against the real file.")
        if f == r:
            print(f"  anchor present: {label}")
            continue
        patched = patched.replace(f, r, 1)
        print(f"  will add: {label}")

    try:
        compile(patched, args.file, "exec")
    except SyntaxError as exc:
        sys.exit(f"Patched file would not compile: line {exc.lineno}: {exc.msg}")
    for token, want in (("evidence_filter.ids_for(", 1),
                        ("import evidence_filter", 1),
                        ('"orphan", "evidence", "evidence_dir")', 1)):
        if patched.count(token) != want:
            sys.exit(f"Sanity check failed: {token!r} appears "
                     f"{patched.count(token)} times, expected {want}.")

    if args.check:
        print("\nPatched file compiles. Re-run without --check to apply.")
        return

    shutil.copy(args.file, BACKUP)
    tmp = args.file + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="") as fh:
        fh.write(patched)
    os.replace(tmp, args.file)
    print(f"\nBacked up to {BACKUP}")
    print(f"Patched {args.file} - restart Streamlit.")


if __name__ == "__main__":
    main()
