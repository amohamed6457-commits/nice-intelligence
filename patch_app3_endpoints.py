import argparse
import os
import shutil
import sys

TARGET = "app3.py"
BACKUP = "app3.py.bak_endpoints"
SENTINEL = "endpoint_panel.render("

# Anchors are SINGLE lines on purpose. A multi-line anchor breaks the moment
# anything nearby changes, and app3.py is edited often; each of these four
# lines occurs exactly once and is unlikely to be reworded.
IMPORT_ANCHOR = """    source_citations = _NoCitations()"""

IMPORT_NEW = '''    source_citations = _NoCitations()


# --- endpoint verdicts, with the resolution page behind them (optional) ---
# Same guard as above: imported before the HTA body switcher, and checked
# with hasattr because a namespace package imports fine and fails on use.
try:
    import endpoint_panel
    if not hasattr(endpoint_panel, "render"):
        raise ImportError("endpoint_panel resolved to a namespace package")
except (ImportError, AttributeError):
    class _NoEndpoints:
        @staticmethod
        def attach(*a, **k):
            return None

        @staticmethod
        def is_attached():
            return False

        @staticmethod
        def render(*a, **k):
            return None

    endpoint_panel = _NoEndpoints()'''

ATTACH_ANCHOR = '    docs = doc_index[doc_index["Institution"] == agency]'

ATTACH_NEW = '''    # The endpoint panel needs the AIS subgroup rows and the same passages,
    # which it parses for each resolution's endpoint-category summary table.
    endpoint_panel.attach(
        load_agency_sheet("GBA_Subgroups", version),
        load_agency_sheet("GBA_Decision_Passages", version),
        frame,
    )
    docs = doc_index[doc_index["Institution"] == agency]'''

RENDER_ANCHOR = '    source_citations.render_panel(agency, row.get("appraisal_id"))'

RENDER_NEW = '''    # Verdicts first, passages underneath: a reader wants the conclusion and
    # its page before the wall of source text.
    endpoint_panel.render(agency, row.get("appraisal_id"))

    source_citations.render_panel(agency, row.get("appraisal_id"))'''

HUNKS = [(IMPORT_ANCHOR, IMPORT_NEW, "guarded import"),
         (ATTACH_ANCHOR, ATTACH_NEW, "attach() call"),
         (RENDER_ANCHOR, RENDER_NEW, "render() call above the passages")]


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
    for need in ("endpoint_panel.py", "subgroup_sources.py",
                 "endpoint_summary.py"):
        if not os.path.exists(need):
            sys.exit(f"{need} must sit beside app3.py before patching.")

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
                     "Your app3.py differs from the copy this was built "
                     "against; send it over and I will redo the patch.")
        patched = patched.replace(f, r, 1)
        print(f"  will add: {label}")

    try:
        compile(patched, args.file, "exec")
    except SyntaxError as exc:
        sys.exit(f"Patched file would not compile: line {exc.lineno}: {exc.msg}")
    # The stub class defines `def render` / `def attach`, which do NOT match
    # the qualified "endpoint_panel.render(" form, so each call site is one.
    for token, want in (("endpoint_panel.render(", 1),
                        ("endpoint_panel.attach(", 1),
                        ("import endpoint_panel", 1)):
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
