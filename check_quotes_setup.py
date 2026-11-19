"""
Why are the theme expanders still showing stubs?

Three things have to be true. This says which one isn't.
Run it in the folder you launch Streamlit from.
"""
import importlib.util, os, sys

print("working directory:", os.getcwd(), "\n")

ok = True

# 1. is the running app3.py actually patched?
if os.path.exists("app3.py"):
    src = open("app3.py", encoding="utf-8", errors="replace").read()
    n = src.count("committee_quotes")
    size = os.path.getsize("app3.py")
    if n:
        print(f"[ok]   app3.py is patched ({n} references, {size:,} bytes)")
    else:
        ok = False
        print(f"[FAIL] app3.py has no committee_quotes reference ({size:,} bytes)")
        print("       You are still running the old file. Replace it with the")
        print("       patched one, then restart Streamlit.")
else:
    ok = False
    print("[FAIL] no app3.py in this folder — you launch Streamlit elsewhere.")

# 2. is the module importable from here?
if importlib.util.find_spec("committee_quotes") is None:
    ok = False
    print("[FAIL] committee_quotes.py not importable from this folder.")
    print("       Without it the app falls back to the workbook text.")
else:
    import committee_quotes as Q
    if not hasattr(Q, "render_theme_evidence"):
        ok = False
        print("[FAIL] committee_quotes imported but has no render_theme_evidence")
        print("       — a directory of that name is shadowing the module.")
    else:
        print(f"[ok]   committee_quotes imports  ({Q.__file__})")

        # 3. is the section index there?
        if Q.available():
            print(f"[ok]   {Q.status()}")
            for ta, theme in (("TA1091", "Long-term extrapolation uncertainty"),
                              ("TA1108", "Evidential or modelling uncertainty"),
                              ("TA557", "Treatment effect waning / durability uncertainty")):
                q = Q.quote_for(ta, theme)
                if q:
                    print(f"       {ta} / {theme[:34]:<34} -> {q['citation']}")
                else:
                    print(f"       {ta} / {theme[:34]:<34} -> nothing indexed on "
                          f"this theme")
        else:
            ok = False
            print(f"[FAIL] {Q.status()}")
            print("       Expected hta_index/keyword_index.json.gz beside app3.py.")
            for p in Q.INDEX_PATHS:
                print(f"         tried: {p}")

print()
print("all three in place." if ok else
      "fix the [FAIL] lines above, then restart Streamlit.")
