"""
Stop preferring a column whose stored text is already cut off.

The bug: original_nice_comment is stored pre-truncated at 80 characters with
the ellipsis baked into the value --

    'At the time of appraisal, the technology was not considered to be an appropri...'

-- and it sits first in the fallback chain, so it wins every time while 700+
characters of usable text sit unread in detailed_reasoning. Nothing in the app
was truncating anything; it was faithfully showing a column that had been
truncated when the workbook was built.

The fix keeps the preference order (the committee's own comment is better than
a generated summary WHEN IT IS COMPLETE) but skips any candidate that has
obviously been cut, falling back to the longest remaining option.

Also teaches committee_quotes.is_stub to recognise an ellipsis-terminated
value, so the degraded path doesn't reintroduce it.

Idempotent. Backs up to app3.py.bak4. Compiles before writing.

  python patch_app3_truncated.py --dry-run
  python patch_app3_truncated.py
"""

import argparse
import os
import shutil
import sys

MARK = "_looks_truncated"

OLD = '''                            # Order matters: the committee's own comment first, the
                            # workbook's summary last. Each is only a fallback - the
                            # indexed discussion is preferred over all of them.
                            for col in ("original_nice_comment", "detailed_reasoning",
                                        "rejection_reasoning"):
                                value = source_row.get(col)
                                if pd.notna(value) and not committee_quotes.is_stub(value):
                                    fallback = str(value).strip()
                                    break
'''

NEW = '''                            # Preference order still holds - the committee's own
                            # comment beats a generated summary - but only when
                            # it is COMPLETE. original_nice_comment is stored
                            # pre-truncated at 80 chars with the ellipsis baked
                            # into the value, so taking the first non-stub
                            # candidate showed 80 characters while 700 sat
                            # unused in the next column.
                            def _looks_truncated(text):
                                t = str(text).rstrip()
                                return t.endswith("...") or t.endswith("\\u2026")

                            _cands = []
                            for col in ("original_nice_comment", "detailed_reasoning",
                                        "rejection_reasoning"):
                                value = source_row.get(col)
                                if pd.notna(value) and not committee_quotes.is_stub(value):
                                    _cands.append(str(value).strip())
                            _whole = [c for c in _cands if not _looks_truncated(c)]
                            if _whole:
                                fallback = _whole[0]
                            elif _cands:
                                # Everything available is cut off; show the most
                                # of it rather than the first of it.
                                fallback = max(_cands, key=len)
'''


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--app", default="app3.py")
    ap.add_argument("--module", default="committee_quotes.py")
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

    if MARK in text:
        print(f"Already applied — {args.app} already handles truncated values.")
    elif OLD not in text:
        print("Could not find the candidate-selection block in "
              f"{args.app}.\nExpected the loop over original_nice_comment / "
              "detailed_reasoning / rejection_reasoning.")
        print("Send me the lines around 'original_nice_comment' and I'll adjust.")
        sys.exit(1)
    else:
        text = text.replace(OLD, NEW, 1)
        try:
            compile(text, args.app, "exec")
        except SyntaxError as e:
            raise SystemExit(f"Patched source does not compile ({e}). "
                             f"{args.app} is untouched.")
        if args.dry_run:
            print("--dry-run: app3.py change compiles, nothing written.")
        else:
            shutil.copy2(args.app, args.app + ".bak4")
            out = text.replace("\n", "\r\n") if crlf else text
            with open(args.app, "wb") as f:
                f.write(out.encode("utf-8"))
            print(f"Wrote {args.app}  (backup: {args.app}.bak4)")

    # --- teach is_stub about ellipsis-terminated stubs -------------------
    if not os.path.exists(args.module):
        print(f"\nnote: {args.module} not here, skipping its half of the fix.")
        return
    mod = open(args.module, encoding="utf-8").read()
    if "ends in an ellipsis" in mod:
        print(f"{args.module}: already handles ellipsis stubs.")
        return

    anchor = '''    if len(s) > 90:
        return False
    return bool(_STUB.match(s))'''
    newer = '''    # A value that ends in an ellipsis was cut off before it got here. On its
    # own that does not make it a stub - a long truncated passage is still
    # worth showing - but a SHORT one carries no information, like
    # "At the time of appraisal, the technology was not considered to be an
    # appropri...".
    if s.rstrip().endswith(("...", "\\u2026")) and len(s) <= 120:
        return True
    if len(s) > 90:
        return False
    return bool(_STUB.match(s))'''
    if anchor not in mod:
        print(f"\nnote: could not find is_stub's tail in {args.module}; "
              "the app3 fix alone is enough.")
        return
    mod = mod.replace(anchor, newer, 1)
    compile(mod, args.module, "exec")
    if args.dry_run:
        print(f"--dry-run: {args.module} change compiles, nothing written.")
        return
    shutil.copy2(args.module, args.module + ".bak")
    with open(args.module, "w", encoding="utf-8", newline="\n") as f:
        f.write(mod)
    print(f"Wrote {args.module}  (backup: {args.module}.bak)")
    print("\nRestart Streamlit.")


if __name__ == "__main__":
    main()
