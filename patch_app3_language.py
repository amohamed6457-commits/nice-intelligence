"""
Put the Original / English switch in the sidebar, beside the HTA body radio.

Why a patcher and not a replacement file: app3.py is 269 KB and may hold
changes this session has not seen. This edits one spot in your copy and
leaves everything else exactly as it was.

The switch has to live in app3's own sidebar rather than in the citation
panel, because the panel only runs once a record is open -- so the control
was invisible on the search page, which is where someone looks for it.

    python patch_app3_language.py            # patch, with a .bak
    python patch_app3_language.py --check    # say what it would do, change nothing
    python patch_app3_language.py --revert   # restore the .bak

Safe to run twice: it detects its own work and stops.
"""

import argparse
import os
import shutil
import sys

TARGET = "app3.py"
BACKUP = "app3.py.bak_language"
SENTINEL = "w_src_lang_mode"

# CRLF is added at write time; keep this block in plain LF.
ANCHOR = "    st.sidebar.divider()\n    return choice"

INSERT = '''    # Passage language. Drawn beside the HTA body switch, not inside the
    # citation panel, so it is visible before a record is opened. Only shown
    # when a translations.json actually exists, and never for NICE, which is
    # already English. source_citations reads this through _show_english();
    # the "w_" prefix keeps the choice when switching HTA body.
    if choice in ("G-BA", "HAS") and getattr(
            source_citations, "translations_loaded", lambda: False)():
        st.sidebar.markdown("**\\U0001F310 Passage language**")
        st.sidebar.radio(
            "Passage language",
            ["Original", "English"],
            horizontal=True,
            key="w_src_lang_mode",
            label_visibility="collapsed",
        )
        st.sidebar.caption(
            ("German" if choice == "G-BA" else "French") + " is the official "
            "record. English is a machine translation shown for reading only "
            "\\u2014 quotes and page links always use the original.")
    st.sidebar.divider()
    return choice'''


def read(path):
    with open(path, encoding="utf-8", newline="") as fh:
        return fh.read()


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

    src = read(args.file)
    crlf = src.count("\r\n") > 100
    print(f"{args.file}: {len(src):,} bytes, line endings "
          f"{'CRLF' if crlf else 'LF'}")

    if SENTINEL in src:
        print("Already patched - nothing to do.")
        return

    anchor = ANCHOR.replace("\n", "\r\n") if crlf else ANCHOR
    found = src.count(anchor)
    if found != 1:
        sys.exit(f"Expected the sidebar anchor exactly once, found {found}.\n"
                 "Your app3.py differs from the one this patch was built "
                 "against; send it over and I will redo the patch.")

    insert = INSERT.replace("\n", "\r\n") if crlf else INSERT
    patched = src.replace(anchor, insert, 1)

    # Must compile, and must still reference the sidebar radio once.
    try:
        compile(patched, args.file, "exec")
    except SyntaxError as exc:
        sys.exit(f"Patched file would not compile: line {exc.lineno}: {exc.msg}")
    if patched.count(SENTINEL) != 1:
        sys.exit("Sanity check failed: the key should appear exactly once.")

    if args.check:
        print("Would insert the language switch before st.sidebar.divider() "
              "in agency_sidebar().")
        print("Patched file compiles. Re-run without --check to apply.")
        return

    shutil.copy(args.file, BACKUP)
    tmp = args.file + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="") as fh:
        fh.write(patched)
    os.replace(tmp, args.file)
    print(f"Backed up to {BACKUP}")
    print(f"Patched {args.file} - restart Streamlit to see it.")
    print("\nThe switch appears under HTA body for G-BA and HAS, and only "
          "when\ntranslations.json is present beside app3.py.")


if __name__ == "__main__":
    main()
