"""
Why isn't the Original / English switch showing up?

Checks every link in the chain and names the one that is broken, instead of
guessing. Run it from the SAME folder you run streamlit from.

    python check_language_toggle.py
"""

import glob
import hashlib
import json
import os
import sys

OK, BAD, WARN = "  OK  ", " FAIL ", " WARN "


def say(state, label, detail=""):
    print(f"[{state}] {label}" + (f"\n         {detail}" if detail else ""))


def main():
    here = os.getcwd()
    print(f"running in: {here}\n")
    problems = []

    # 1. Are we even in the right folder?
    if not os.path.exists("app3.py"):
        say(BAD, "app3.py is not in this folder",
            "Streamlit must be started from the folder holding app3.py, and "
            "so must this check.")
        found = glob.glob(os.path.join(os.path.expanduser("~"), "**", "app3.py"),
                          recursive=True)[:3]
        if found:
            print("         app3.py looks like it lives in:")
            for f in found:
                print("           ", os.path.dirname(f))
        return
    say(OK, "app3.py found here")

    # 2. translations.json present, and actually full of translations?
    path = "translations.json"
    if not os.path.exists(path):
        say(BAD, "translations.json is NOT in this folder", "This is almost "
            "certainly the problem - the switch only appears when it exists.")
        stray = []
        for root in (os.path.expanduser("~\\Downloads"),
                     os.path.expanduser("~/Downloads"), here, ".."):
            p = os.path.join(root, "translations.json")
            if os.path.exists(p):
                stray.append(p)
        for p in glob.glob(os.path.join(here, "**", "translations.json"),
                           recursive=True):
            stray.append(p)
        if stray:
            print("         but a translations.json exists at:")
            for p in dict.fromkeys(stray):
                print(f"           {p}  ({os.path.getsize(p)/1e6:.1f} MB)")
            print("         Move it next to app3.py.")
        else:
            print("         None found anywhere obvious. Has the translation "
                  "run finished?\n         It writes the file only after the "
                  "first 25 passages.")
        problems.append("translations.json missing")
    else:
        size = os.path.getsize(path) / 1e6
        try:
            with open(path, encoding="utf-8") as fh:
                data = json.load(fh)
        except (ValueError, OSError) as exc:
            say(BAD, "translations.json cannot be read", str(exc))
            return
        if not isinstance(data, dict):
            say(BAD, "translations.json is not a dictionary")
            return
        errors = [v for v in data.values() if str(v).startswith("__ERROR__")]
        good = len(data) - len(errors)
        say(OK if good else BAD,
            f"translations.json: {len(data):,} entries, {good:,} usable, "
            f"{len(errors):,} errors, {size:.1f} MB")
        if errors:
            from collections import Counter
            kinds = Counter(str(v)[10:60] for v in errors)
            for kind, n in kinds.most_common(3):
                print(f"           {n:>5}  {kind}")
        if not good:
            problems.append("every translation is an error row")

    # 3. Is source_citations the new one?
    try:
        src = open("source_citations.py", encoding="utf-8").read()
    except OSError:
        say(BAD, "source_citations.py not found here")
        return
    for needed, label in [("def translations_loaded", "translations_loaded()"),
                          ("w_src_lang_mode", "reads w_src_lang_mode"),
                          ("def translation_for", "translation_for()")]:
        if needed in src:
            say(OK, f"source_citations.py has {label}")
        else:
            say(BAD, f"source_citations.py is the OLD file - missing {label}",
                "The download did not overwrite it.")
            problems.append("source_citations.py not updated")
    if "st.radio(" in src:
        say(WARN, "source_citations.py still draws its own radio",
            "Two widgets on one key will crash Streamlit. Use the newest file.")
        problems.append("duplicate widget key")

    # 4. Is app3 patched?
    app = open("app3.py", encoding="utf-8", newline="").read()
    if "w_src_lang_mode" in app:
        say(OK, "app3.py is patched (sidebar switch present)")
    else:
        say(BAD, "app3.py is NOT patched",
            "Run: python patch_app3_language.py")
        problems.append("app3.py not patched")

    # 5. Would the module actually report translations as loaded?
    sys.path.insert(0, here)
    try:
        import importlib.util
        import types
        st = types.ModuleType("streamlit")
        st.session_state = {}
        st.cache_data = lambda *a, **k: (lambda f: f) if not a else a[0]
        st.sidebar = types.SimpleNamespace(
            **{n: (lambda *a, **k: None) for n in
               ("radio", "markdown", "caption", "divider", "title")})
        for n in ("markdown", "caption", "info", "warning", "write", "radio"):
            setattr(st, n, lambda *a, **k: None)

        class _E:
            def __enter__(self): return None
            def __exit__(self, *a): return False
        st.expander = lambda *a, **k: _E()
        sys.modules["streamlit"] = st
        spec = importlib.util.spec_from_file_location("_sc", "source_citations.py")
        sc = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(sc)
        loaded = sc.translations_loaded()
        say(OK if loaded else BAD,
            f"source_citations.translations_loaded() -> {loaded}")
        if not loaded:
            problems.append("module cannot load the translations")
    except Exception as exc:                               # noqa: BLE001
        say(WARN, f"could not import source_citations: "
                  f"{type(exc).__name__}: {exc}")
        sc = None

    # 6. The silent killer: do the keys match real passages?
    if sc is not None and os.path.exists("translations.json"):
        books = [f for f in os.listdir(".") if f.endswith(".xlsx")
                 and "NICE" in f]
        if books:
            try:
                import pandas as pd
                book = sorted(books)[-1]
                df = pd.read_excel(book, sheet_name="HAS_Economic_Pages")
                texts = [str(t).strip() for t in
                         df["source_text"].dropna().head(400)]
                hits = sum(1 for t in texts
                           if sc.translation_for(t) is not None)
                rate = hits / max(1, len(texts))
                say(OK if rate > 0.3 else BAD,
                    f"key match on real passages: {hits}/{len(texts)} "
                    f"({rate:.0%}) from {book}")
                if rate <= 0.3:
                    print("         The file exists but its keys do not match "
                          "the workbook text.\n         Nothing would "
                          "translate even with the switch showing.")
                    problems.append("translation keys do not match")
            except Exception as exc:                       # noqa: BLE001
                say(WARN, f"could not sample the workbook: {exc}")

    print()
    if problems:
        print("VERDICT: " + "; ".join(problems))
    else:
        print("VERDICT: everything checks out. If the switch is still "
              "missing,\n         Streamlit was not restarted - Ctrl+C, then "
              "streamlit run app3.py.\n         It only shows on G-BA and HAS, "
              "never on NICE.")


if __name__ == "__main__":
    main()
