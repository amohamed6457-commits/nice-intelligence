"""
Wire nice_attributes.py into app3.py.

Four edits. Idempotent - run it twice and the second run reports "already
applied" and changes nothing. Writes app3.py.bak first.

It patches YOUR app3.py in place, so the committee-sections work already in
that file is untouched; it never replaces the file wholesale.

Usage:
  python patch_app3_attributes.py
  python patch_app3_attributes.py --app app3.py --dry-run
"""

import argparse
import os
import shutil
import sys

MARK = "nice_attributes"


# Regex-anchored edits, tried in order until one matches. Used where the
# surrounding code varies between versions of app3.py: a literal anchor there
# fails on a single changed space, and guessing a replacement is worse than
# saying so.
#
# Each entry: (description, [candidate patterns], line to insert BEFORE the
# match, indented to match the matched line).
REGEX_EDITS = [
    (
        "attribute panel in the record view",
        [
            r"^(?P<i>[ \t]*)record_docs\s*=\s*docs\[",
            # `if links:` sits at function level; the st.markdown inside it
            # does not, so matching that would nest the panel in the branch
            # and only show it for records that have source documents.
            r"^(?P<i>[ \t]*)if links:",
            r"^(?P<i>[ \t]*)st\.markdown\(\s*[\"\']\*\*Source documents\*\*",
        ],
        [
            "# Silent for anything without coded attributes: every non-NICE",
            "# record, and the NICE appraisals with no committee discussion",
            "# indexed.",
            'nice_attributes.render_panel(row.get("appraisal_id"))',
            "",
        ],
    ),
]


def apply_regex_edits(text):
    """Returns (text, applied, failures). Inserts before the first match."""
    import re as _re
    applied, failures = [], []
    for desc, patterns, lines in REGEX_EDITS:
        for pat in patterns:
            m = _re.search(pat, text, _re.MULTILINE)
            if not m:
                continue
            indent = m.group("i")
            block = "".join(f"{indent}{ln}\n" if ln else "\n" for ln in lines)
            text = text[: m.start()] + block + text[m.start():]
            applied.append(f"{desc}  (matched /{pat[:38]}.../)")
            break
        else:
            failures.append(desc)
    return text, applied, failures


# (description, anchor that must appear exactly once, replacement)
def edits():
    return [
        (
            "import block with namespace-package guard",
            "st.sidebar.title(\"🔍 Filters\")",
            '''# --- evidence attributes (optional) ------------------------------------
# A module whose name matches a data file or directory can import as an empty
# namespace package, which succeeds and then fails on first use. So check for
# a known attribute rather than trusting the import.
# NB: must not be written as a bare expression - Streamlit's magic would
# render it to the page.
try:
    import nice_attributes
    if not hasattr(nice_attributes, "FIELDS"):
        raise ImportError("nice_attributes resolved to a namespace package")
except (ImportError, AttributeError):
    class _NoAttributes:
        FIELDS = []

        @staticmethod
        def available():
            return False

        @staticmethod
        def attach(frame, *a, **k):
            return frame, 0

        @staticmethod
        def sidebar_filters(*a, **k):
            return {}, {}

        @staticmethod
        def apply_filters(frame, *a, **k):
            return frame

        @staticmethod
        def render_panel(*a, **k):
            return None

        @staticmethod
        def render_status_caption(*a, **k):
            return None

        @staticmethod
        def coverage(ids):
            return 0, len(list(ids))

    nice_attributes = _NoAttributes()

st.sidebar.title("🔍 Filters")''',
        ),
        (
            "sidebar filters + attach, applied after the existing filters",
            '''filtered_df = filtered_df[
    filtered_df["year_start"].between(year_range[0], year_range[1])
]''',
            '''filtered_df = filtered_df[
    filtered_df["year_start"].between(year_range[0], year_range[1])
]

# Evidence attributes: attach first so the filters can offer only values that
# actually occur in the current result set, then narrow. Both are no-ops when
# evidence_attributes.json is absent.
filtered_df, _ev_matched = nice_attributes.attach(filtered_df)
_ev_selections, _ev_options = nice_attributes.sidebar_filters(
    filtered_df["appraisal_id"] if "appraisal_id" in filtered_df.columns else []
)
_ev_before = len(filtered_df)
filtered_df = nice_attributes.apply_filters(
    filtered_df, _ev_selections, **_ev_options
)''',
        ),
        (
            "coverage caption under the metrics",
            '''st.divider()

export_cols = [c for c in filtered_df.columns if c != "search_blob"]''',
            '''if nice_attributes.available():
    _n_coded, _n_total = nice_attributes.coverage(filtered_df["appraisal_id"])
    _note = (f"Evidence attributes coded for {_n_coded} of these {_n_total} "
             f"appraisals.")
    if _ev_selections:
        _note += (f" Attribute filters narrowed {_ev_before} → "
                  f"{len(filtered_df)}.")
    _note += (" Values marked ~ are where two independent extraction passes "
              "disagreed — about 11% of cells, mostly the stance fields.")
    st.caption(_note)

st.divider()

export_cols = [c for c in filtered_df.columns if c != "search_blob"]''',
        ),
    ]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--app", default="app3.py")
    ap.add_argument("--dry-run", action="store_true",
                    help="report what would change, write nothing")
    args = ap.parse_args()

    if not os.path.exists(args.app):
        raise SystemExit(f"Not found: {args.app}  (run this beside app3.py)")

    # Binary read so the file's own line endings survive. app3.py is CRLF and
    # rewriting it as LF makes every line of a diff look changed.
    with open(args.app, "rb") as f:
        raw = f.read()
    crlf = b"\r\n" in raw
    text = raw.decode("utf-8")
    if crlf:
        text = text.replace("\r\n", "\n")

    if MARK in text:
        print(f"Already applied — {args.app} already references {MARK}.")
        print("Nothing to do. Delete app3.py.bak if you want a fresh backup.")
        return

    todo = edits()
    problems = []
    for desc, anchor, _ in todo:
        n = text.count(anchor)
        if n != 1:
            problems.append(f"  {desc}: anchor found {n} times, expected 1")
    if problems:
        print("Refusing to patch — the file does not look as expected:")
        print("\n".join(problems))
        print("\nYour app3.py has probably moved on. Paste me the surrounding")
        print("lines and I will adjust the anchors rather than guess.")
        sys.exit(1)

    for desc, anchor, replacement in todo:
        text = text.replace(anchor, replacement, 1)
        print(f"  applied: {desc}")

    text, applied, failures = apply_regex_edits(text)
    for line in applied:
        print(f"  applied: {line}")
    if failures:
        print("\nCould not place these, and will not guess:")
        for f in failures:
            print(f"  {f}")
        print("\nRun this and paste the output:")
        print("  Select-String -Path app3.py -Pattern 'record_docs|Source "
              "documents' -Context 3,3")
        sys.exit(1)

    # Compile before writing. A syntax error reaching app3.py means a blank
    # page in Streamlit with the traceback buried in the server log.
    try:
        compile(text, args.app, "exec")
    except SyntaxError as e:
        raise SystemExit(f"Patched source does not compile ({e}). "
                         f"{args.app} is untouched.")

    if args.dry_run:
        print("\n--dry-run: nothing written. Patched source compiles.")
        return

    shutil.copy2(args.app, args.app + ".bak")
    out = text.replace("\n", "\r\n") if crlf else text
    with open(args.app, "wb") as f:
        f.write(out.encode("utf-8"))

    print(f"\nWrote {args.app}  (backup: {args.app}.bak, "
          f"{'CRLF' if crlf else 'LF'} preserved)")
    print("Restart Streamlit. If evidence_attributes.json is missing the app")
    print("behaves exactly as before — every call degrades to a no-op.")


if __name__ == "__main__":
    main()
