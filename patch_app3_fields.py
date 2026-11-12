import argparse
import os
import shutil
import sys

TARGET = "app3.py"
BACKUP = "app3.py.bak_fields"
# Must be distinctive: a bare "display_text" also matches Streamlit's
# LinkColumn(display_text=...), which app3 already uses twice, and the
# patcher would wrongly report itself as already applied.
SENTINEL = "source_citations.display_text("

HUNKS = [
    # (what to find, what to replace it with, description)
    ('        st.markdown(f"**Indication (as recorded):** '
     '{_md_safe(_shorten(indication, 700))}")',
     '        st.markdown(f"**Indication (as recorded):** '
     '{_md_safe(_shorten(source_citations.display_text(indication), 700))}")',
     "Indication (as recorded)"),

    ('    for label, col in cfg["detail_fields"]:\n'
     '        value = _text(row.get(col))\n'
     '        if value:\n'
     '            st.markdown(f"**{label}:** {_md_safe(value)}")',
     '    for label, col in cfg["detail_fields"]:\n'
     '        value = _text(row.get(col))\n'
     '        if value:\n'
     '            st.markdown(f"**{label}:** "\n'
     '                        f"{_md_safe(source_citations.display_text(value))}")',
     "comparator / patient population / added benefit"),
]


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
    print(f"{args.file}: {len(src):,} bytes, {'CRLF' if crlf else 'LF'}")

    if SENTINEL in src:
        print("Already patched - nothing to do.")
        return
    if "w_src_lang_mode" not in src:
        sys.exit("Run patch_app3_language.py first - there is no sidebar "
                 "switch to obey yet.")

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
        print(f"  will wrap: {label}")

    try:
        compile(patched, args.file, "exec")
    except SyntaxError as exc:
        sys.exit(f"Patched file would not compile: line {exc.lineno}: {exc.msg}")
    if patched.count("source_citations.display_text(") != 2:
        sys.exit("Sanity check failed: expected exactly two wrapped fields.")

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
    print("\nThe record body will only change for values that are IN "
          "translations.json.\nRe-run translate_passages.py first: the "
          "record fields were never in it.")


if __name__ == "__main__":
    main()
