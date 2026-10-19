"""
Move the optional-module imports to the top of app3.py.

The bug: both blocks were inserted next to `st.sidebar.title("Filters")`,
around line 4270. But `render_agency_view()` runs at line ~4200, and the
G-BA and HAS views render an appraisal record from there — which calls
nice_attributes.render_panel() before the import statement has executed.

    NameError: name 'nice_attributes' is not defined

NICE happened to work because its path reaches that code later. Python
resolves a module-level name at call time, so the fix is simply to bind the
name earlier: move both blocks up beside the nice_sections block that was
already correctly placed at the top.

Idempotent. Backs up to app3.py.bak3. Compiles before writing.

  python patch_app3_imports.py --dry-run
  python patch_app3_imports.py
"""

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


def first_trigger(text):
    """Offset of the earliest module-level statement that can reach these
    modules. Everything must be bound before this point."""
    # The real call is indented - it sits inside an `if` at module level -
    # so anchoring at column zero misses it entirely, which is how the first
    # version of this patcher decided the blocks were already fine.
    hits = [m.start() for m in
            re.finditer(r"^[ \t]*render_agency_view\(", text, re.M)]
    hits += [m.start() for m in
             re.finditer(r"^[ \t]*st\.sidebar\.title\(", text, re.M)]
    return min(hits) if hits else len(text)


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
        # "Near the top" is not a character count - it is "before the first
        # module-level statement that could use it". A fixed window broke
        # idempotency, because the second block legitimately sits a few
        # thousand characters past the anchor.
        if pos < first_trigger(text):
            already.append(name)
            print(f"  [ok]   {name}: already above the first top-level use")
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

    # Sanity: the binding must happen before the module-level statements that
    # can reach it. Textual order is only meaningful for code that runs at
    # import time - a name used inside a function body is resolved when the
    # function is called, which is why the original check was wrong.
    #
    # The statement that triggered this bug is the top-level call into the
    # agency views, so that is the line to be in front of.
    earliest_trigger = first_trigger(text)

    for mod, binder in (("nice_attributes", "_NoAttributes"),
                        ("committee_quotes", "_NoQuotes")):
        m = re.search(rf"^\s*{mod} = {binder}\(\)$", text, re.M)
        if not m:
            continue
        if m.start() > earliest_trigger:
            raise SystemExit(
                f"After patching, {mod} is still bound after the first "
                f"module-level code that uses it. Not writing.")
        print(f"  [check] {mod} bound at char {m.start():,}, "
              f"first top-level trigger at {earliest_trigger:,}")

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
