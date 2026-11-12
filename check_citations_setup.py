

import importlib.util
import os
import re
import sys

SHEETS = ["GBA_Decision_Passages", "HAS_Economic_Pages", "Economic_Findings"]
ok = True


def good(msg):
    print(f"  [ok]   {msg}")


def bad(msg):
    global ok
    ok = False
    print(f"  [FAIL] {msg}")


print("\n1. The module")
if not os.path.exists("source_citations.py"):
    bad("source_citations.py is NOT in this folder. app3 will fall back to "
        "the silent stand-in and no citations will ever render.")
else:
    good("source_citations.py is here")
    spec = importlib.util.spec_from_file_location("_sc", "source_citations.py")
    mod = importlib.util.module_from_spec(spec)
    try:
        # streamlit must import for the module to load; if it doesn't, that is
        # itself the problem.
        spec.loader.exec_module(mod)
    except Exception as exc:                       # noqa: BLE001
        bad(f"it does not import: {type(exc).__name__}: {exc}")
        mod = None
    if mod is not None:
        if hasattr(mod, "render_panel"):
            good("render_panel() is present")
        else:
            bad("render_panel() missing — resolved to a namespace package?")

print("\n2. app3.py wiring")
if not os.path.exists("app3.py"):
    bad("app3.py not found — run this from the folder that holds it")
    data_file = None
else:
    src = open("app3.py", encoding="utf-8").read()
    for label, pattern in (
        ("import", r"^\s*import source_citations\b"),
        ("attach() call", r"source_citations\.attach\("),
        ("render_panel() call", r"source_citations\.render_panel\("),
    ):
        (good if re.search(pattern, src, re.M) else bad)(f"{label} present")

    imp = re.search(r"^\s*import source_citations\b", src, re.M)
    sw = re.search(r"^HTA_BODY\s*=\s*render_agency_switcher\(\)", src, re.M)
    if imp and sw:
        li = src[:imp.start()].count("\n") + 1
        ls = src[:sw.start()].count("\n") + 1
        if li < ls:
            good(f"imported at line {li}, before the switcher at line {ls}")
        else:
            bad(f"imported at line {li}, AFTER the switcher at line {ls} — "
                "a G-BA/HAS record will raise NameError")

    m = re.search(r'^DATA_FILE\s*=\s*["\'](.+?)["\']', src, re.M)
    data_file = m.group(1) if m else None

print("\n3. The workbook")
if not data_file:
    bad("could not read DATA_FILE from app3.py")
elif not os.path.exists(data_file):
    bad(f"app3.py points at '{data_file}', which is not in this folder")
    near = [f for f in os.listdir(".") if f.lower().endswith((".xlsx", ".xlsm"))]
    if near:
        print(f"         workbooks that ARE here: {', '.join(near)}")
else:
    good(f"DATA_FILE '{data_file}' exists")
    try:
        import pandas as pd
        names = pd.ExcelFile(data_file).sheet_names
        print(f"         {len(names)} sheets")
        for s in SHEETS:
            if s in names:
                n = len(pd.read_excel(data_file, sheet_name=s))
                good(f"{s}: {n:,} rows")
            else:
                bad(f"{s} is MISSING — this is the short build of the "
                    "workbook; the panel will render nothing")
    except Exception as exc:                       # noqa: BLE001
        bad(f"could not read it: {type(exc).__name__}: {exc}")

print("\n" + ("All checks passed. Restart Streamlit and open one of the demo "
              "records:\n  G-BA  D-935 (Cemiplimab)\n  HAS   CT-15634 "
              "(KEYTRUDA, NSCLC)" if ok else
              "Fix the [FAIL] lines above, then run this again."))
sys.exit(0 if ok else 1)
