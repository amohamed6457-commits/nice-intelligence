import argparse
import os
import re
import shutil
import sys

# Where the blocks should end up: immediately after the nice_sections
# fallback, which is already in the right place.
ANCHOR = "    nice_sections = _NoChapters()\n"

# Each block runs from its banner comment to the line that binds the name.
BLOCKS = [
    ("evidence attributes",
     "# --- evidence attributes (optional)",
     "    nice_attributes = _NoAttributes()\n"),
    ("committee quotes",
     "# --- committee quotes (optional)",
     "    committee_quotes = _NoQuotes()\n"),
]


def cut(text, banner, tail):
    """Remove a block and return (text_without_it, the_block)."""
    start = text.find(banner)
    if start == -1:
        return text, None
    end = text.find(tail, start)
    if end == -1:
        return text, None
    end += len(tail)
    block = text[start:end]
    # Swallow one trailing blank line so removing the block doesn't leave a gap.
    after = end
    if text[after:after + 1] == "\n":
        after += 1
    return text[:start] + text[after:], block


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--app", default="app3.py")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    if not os.path.exists(args.app):
        raise SystemExit(f"Not found: {args.app}")

    with open(args.app, "rb") as f:
        raw = f.read()
    crlf = b"\r\n" in raw
    text = raw.decode("utf-8")
    if crlf:
        text = text.replace("\r\n", "\n")

    if ANCHOR not in text:
        raise SystemExit(
            "Could not find the nice_sections block to anchor against.\n"
            "Expected a line reading:    nice_sections = _NoChapters()")

    anchor_at = text.index(ANCHOR)
    moved, already = [], []

    for name, banner, tail in BLOCKS:
        pos = text.find(banner)
        if pos == -1:
            print(f"  [skip] {name}: block not found — was it ever applied?")
            continue
        if pos < anchor_at + len(ANCHOR) + 200:
            already.append(name)
            print(f"  [ok]   {name}: already near the top")
            continue
        text, block = cut(text, banner, tail)
        if block is None:
            raise SystemExit(f"Found the start of the {name} block but not "
                             f"its end ({tail.strip()!r}). Not touching the file.")
        # Re-find the anchor: cutting shifted everything after it.
        at = text.index(ANCHOR) + len(ANCHOR)
        text = text[:at] + "\n" + block + text[at:]
        moved.append(name)
        print(f"  [move] {name}: -> top of file")

    # While we are here: a bare attribute reference at module level is an
    # expression statement, and Streamlit's magic renders those to the page.
    # The hasattr form is identical in behaviour and cannot be rendered.
    bare = "    nice_sections.CITATION_RULES\n"
    if bare in text:
        text = text.replace(
            bare,
            '    if not hasattr(nice_sections, "CITATION_RULES"):\n'
            '        raise ImportError("nice_sections resolved to a '
            'namespace package")\n', 1)
        moved.append("nice_sections bare-expression guard")
        print("  [fix]  nice_sections.CITATION_RULES -> hasattr check "
              "(bare expressions can be rendered by Streamlit's magic)")

    if not moved:
        print("\nNothing to change." if already else
              "\nNothing matched — check the file is the patched app3.py.")
        return

    # Sanity: every name must be bound before its first use.
    #
    # Two things this must NOT trip over, both of which it used to:
    #
    #   "nice_sections.py not deployed"   - the module name inside a message
    #       string. '\bmod\.\w+' matches 'nice_sections.py' happily, and that
    #       string sits above the fallback binding, so the check reported a
    #       use-before-bind that does not exist. Hence the (?!py\b) guard.
    #
    #   the probe inside the try block - 'nice_sections.CITATION_RULES' runs
    #       only when the import SUCCEEDED, while the '= _No...()' binding runs
    #       only in the except branch. They are alternative paths, not
    #       sequential ones, so a textual comparison between them is
    #       meaningless. Only uses AFTER the whole try/except block matter,
    #       so the search starts from the end of the binding's own line.
    for mod in ("nice_attributes", "committee_quotes", "nice_sections"):
        importer = re.search(rf"^\s*import {mod}\b", text, re.M)
        if not importer:
            continue
        # Everything above the import statement. A use down here runs before
        # the module exists on ANY path, which is the real failure. Uses
        # inside or below the try/except are fine: the try branch only runs
        # after a successful import, the except branch binds the stand-in.
        before = text[:importer.start()]
        stray = re.search(rf"\b{mod}\.(?!py\b)\w+", before)
        if stray:
            line = before[:stray.start()].count("\n") + 1
            raise SystemExit(
                f"After patching, {mod} is used at line {line}, above its "
                f"import. Not writing.")

    try:
        compile(text, args.app, "exec")
    except SyntaxError as e:
        raise SystemExit(f"Patched source does not compile ({e}). "
                         f"{args.app} is untouched.")

    if args.dry_run:
        print("\n--dry-run: nothing written. Patched source compiles and "
              "every module is bound before first use.")
        return

    shutil.copy2(args.app, args.app + ".bak3")
    out = text.replace("\n", "\r\n") if crlf else text
    with open(args.app, "wb") as f:
        f.write(out.encode("utf-8"))
    print(f"\nWrote {args.app}  (backup: {args.app}.bak3, "
          f"{'CRLF' if crlf else 'LF'} preserved)")
    print("Restart Streamlit and open the G-BA view.")


if __name__ == "__main__":
    main()
