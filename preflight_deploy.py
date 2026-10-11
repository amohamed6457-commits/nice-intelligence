"""
Check the repo is deployable before you push.

Streamlit Cloud fails quietly in exactly the ways this catches: a data file
that never got committed, so the app runs with every optional feature
degraded to a no-op; a dependency that works locally because you pip
installed it months ago; a secrets file that went up with everything else.

Run it in the repo folder. It changes nothing.

  python preflight_deploy.py
"""

import ast
import os
import re
import subprocess
import sys

# (path, why it matters, fatal if missing)
REQUIRED = [
    ("app3.py", "the app itself", True),
    ("requirements.txt", "Streamlit Cloud installs from this", True),
]

DATA = [
    ("evidence_attributes.json", "the 13 coded attributes — without it the "
     "sidebar filters vanish silently"),
    ("hta_index/keyword_index.json.gz", "committee sections — without it the "
     "theme expanders fall back to workbook summaries"),
    ("committee_quotes.py", "quote lookup"),
    ("nice_attributes.py", "attribute filters"),
    ("nice_sections.py", "chat grounding"),
]

MUST_IGNORE = [
    (".streamlit/secrets.toml", "your API key"),
    ("nice_chapters/", "504 HTML files"),
    ("__pycache__/", "compiled bytecode"),
]

# import name -> pip name, where they differ
PIP_NAME = {
    "bs4": "beautifulsoup4", "PIL": "pillow", "sklearn": "scikit-learn",
    "yaml": "pyyaml", "dateutil": "python-dateutil", "fitz": "pymupdf",
    "dotenv": "python-dotenv",
}
STDLIB = set(sys.stdlib_module_names)


def human(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f}{unit}" if unit == "B" else f"{n:.1f}{unit}"
        n /= 1024


def git(*args):
    try:
        out = subprocess.run(["git", *args], capture_output=True, text=True,
                             timeout=20)
        return out.stdout if out.returncode == 0 else None
    except (OSError, subprocess.SubprocessError):
        return None


def top_imports(path):
    """Top-level imports of a module, without executing it."""
    try:
        tree = ast.parse(open(path, encoding="utf-8", errors="replace").read())
    except (SyntaxError, OSError):
        return set()
    names = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            names |= {a.name.split(".")[0] for a in node.names}
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            names.add(node.module.split(".")[0])
    return names


def main():
    problems, warnings = [], []
    print(f"Checking {os.getcwd()}\n")

    # ---- files the app needs -------------------------------------------
    print("required")
    for path, why, fatal in REQUIRED:
        if os.path.exists(path):
            print(f"  [ok]   {path:<34} {human(os.path.getsize(path))}")
        else:
            print(f"  [FAIL] {path:<34} missing — {why}")
            problems.append(f"{path} missing")

    print("\ndata and modules (missing = feature silently disabled)")
    local_modules = set()
    for path, why in DATA:
        if os.path.exists(path):
            size = os.path.getsize(path)
            print(f"  [ok]   {path:<34} {human(size)}")
            if path.endswith(".py"):
                local_modules.add(os.path.basename(path)[:-3])
            if size > 50 * 1024 * 1024:
                warnings.append(f"{path} is {human(size)} — GitHub warns "
                                f"above 50MB and rejects above 100MB")
        else:
            print(f"  [warn] {path:<34} missing — {why}")
            warnings.append(f"{path} missing")

    # The workbook has a dated name, so find it rather than hardcode.
    books = [f for f in os.listdir(".") if f.lower().endswith((".xlsx", ".xlsm"))
             and not f.startswith("~$")]
    print()
    if books:
        for b in sorted(books):
            print(f"  [ok]   workbook: {b:<25} {human(os.path.getsize(b))}")
    else:
        print("  [FAIL] no .xlsx in this folder — the app cannot load its data")
        problems.append("no workbook")

    # ---- things that must NOT be committed ------------------------------
    print("\nmust not be committed")
    gitignore = ""
    if os.path.exists(".gitignore"):
        gitignore = open(".gitignore", encoding="utf-8", errors="replace").read()
    tracked = git("ls-files") or ""
    for pattern, what in MUST_IGNORE:
        stem = pattern.rstrip("/")
        ignored = any(line.strip().rstrip("/") == stem
                      for line in gitignore.splitlines())
        is_tracked = any(line.startswith(stem) for line in tracked.splitlines())
        if is_tracked:
            print(f"  [FAIL] {pattern:<34} IS TRACKED BY GIT — {what}")
            problems.append(f"{pattern} is tracked")
        elif ignored:
            print(f"  [ok]   {pattern:<34} gitignored")
        elif os.path.exists(stem):
            print(f"  [warn] {pattern:<34} exists and is not in .gitignore")
            warnings.append(f"{pattern} not gitignored")
        else:
            print(f"  [ok]   {pattern:<34} not present")

    if "secrets.toml" in tracked:
        print("\n  !! A secrets file is in git history. Rotate the API key at")
        print("     console.anthropic.com. Removing the file does not undo the")
        print("     exposure — the old key stays valid until you revoke it.")

    # ---- dependencies ---------------------------------------------------
    print("\ndependencies")
    reqs = ""
    if os.path.exists("requirements.txt"):
        reqs = open("requirements.txt", encoding="utf-8", errors="replace").read()
    listed = {re.split(r"[=<>!~\[ ]", line.strip())[0].lower()
              for line in reqs.splitlines()
              if line.strip() and not line.startswith("#")}

    needed = set()
    for f in ["app3.py"] + [p for p, _ in DATA if p.endswith(".py")]:
        if os.path.exists(f):
            needed |= top_imports(f)
    needed -= STDLIB | local_modules | {"app3"}

    for mod in sorted(needed):
        pip = PIP_NAME.get(mod, mod).lower()
        if pip in listed or mod.lower() in listed:
            print(f"  [ok]   {mod:<20} -> {pip}")
        else:
            print(f"  [FAIL] {mod:<20} -> {pip}  NOT in requirements.txt")
            problems.append(f"{pip} missing from requirements.txt")

    # pandas picks its Excel engine at call time, so an engine= argument is
    # a dependency that no import statement reveals.
    if os.path.exists("app3.py"):
        src = open("app3.py", encoding="utf-8", errors="replace").read()
        for engine, pkg in (("calamine", "python-calamine"),
                            ("openpyxl", "openpyxl"),
                            ("xlrd", "xlrd"), ("odf", "odfpy")):
            if f'engine="{engine}"' in src or f"engine='{engine}'" in src:
                if pkg.lower() in listed:
                    print(f"  [ok]   engine=\"{engine}\"{'':<10} -> {pkg}")
                else:
                    print(f"  [FAIL] engine=\"{engine}\" needs {pkg}, which is "
                          f"not in requirements.txt")
                    problems.append(f"{pkg} missing (used via engine=)")

    # ---- repo size ------------------------------------------------------
    total = 0
    for root, dirs, files in os.walk("."):
        dirs[:] = [d for d in dirs if d not in (".git", "__pycache__",
                                                "nice_chapters", ".venv", "venv")]
        for f in files:
            try:
                total += os.path.getsize(os.path.join(root, f))
            except OSError:
                pass
    print(f"\nrepo size excluding .git and nice_chapters: {human(total)}")
    if total > 500 * 1024 * 1024:
        warnings.append("repo is large; Streamlit Cloud clones it on every boot")

    # ---- verdict --------------------------------------------------------
    print("\n" + "=" * 58)
    if problems:
        print(f"{len(problems)} blocking:")
        for p in problems:
            print(f"  - {p}")
    if warnings:
        print(f"{len(warnings)} to think about:")
        for w in warnings:
            print(f"  - {w}")
    if not problems and not warnings:
        print("ready to push.")
    elif not problems:
        print("\nNo blockers. The warnings mean features degrade quietly "
              "rather than erroring,")
        print("so the app will look like it works while doing less.")


if __name__ == "__main__":
    main()
