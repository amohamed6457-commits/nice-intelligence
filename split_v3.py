"""
Move a group of definitions out of app3.py into a module of its own.

Why a tool and not an editor
----------------------------
app3.py is 5,720 lines. Cutting and pasting by hand across a file that size
is how a working app stops working, and the failure shows up as a NameError
on a page nobody opened during testing. This does the move mechanically and
then PROVES the result: both files have to compile, and every name the
remaining code still reads has to resolve to something.

What makes the split safe here
------------------------------
The groups below were chosen after checking what each function actually
reaches for. None of them touch the dataframe, the session state, or any
mutable module-level object -- they reach only constants, and the constants
move with them. That is why this is a mechanical move rather than a
redesign.

How it runs
-----------
One group at a time, verifying after each, because a failure after one move
is obvious and a failure after five is an afternoon.

    python split_app3.py --list
    python split_app3.py --group text --dry-run
    python split_app3.py --group text
    python split_app3.py --verify

A .bak is written before anything changes, and --revert puts it back.
"""

import argparse
import ast
import builtins
import os
import re
import shutil
import sys

TARGET = "app3.py"
BACKUP = "app3.py.bak_split"

# Each group: the module it becomes, and the definitions that move. Constants
# are found automatically -- a module-level name a moved function reads, that
# nothing left behind still needs, travels with it.
GROUPS = {
    # Six files, in dependency order. Each module only ever reads from ones
    # created before it, so no import cycle is possible. The grouping is by
    # FEATURE rather than by what was convenient to move, because the point
    # is to be able to open one file and review one thing.
    "data": ("hta_data.py", [
        "data_version", "_workbook", "_read_sheet", "load_data",
        "load_enrichment_log", "load_vocabulary", "vocab_options",
        "load_agency_sheet", "load_document_index",
    ]),
    "core": ("hta_core.py", [
        # reading a decision: text handling, G-BA, HAS, classification,
        # and the query expansion behind search
        "_text", "_substantive", "_shorten", "_md_safe", "_md_link",
        "_parse_date", "_to_datetime", "_date_span", "format_eur_icer",
        "_decision_text", "_clauses", "_extents_in", "_gba_extents",
        "gba_outcome", "gba_certainty", "gba_favourable",
        "gba_outcome_with_status", "gba_favourable_with_status",
        "gba_subtitle", "gba_headlines",
        "has_is_early_access", "asmr_level", "smr_level", "has_outcome",
        "has_outcome_labelled", "has_favourable", "has_ratings_text",
        "has_subtitle", "has_headlines",
        "family_order", "outcome_family", "_theme_mode", "_fold",
        "_fold_series", "_english_area", "therapy_areas",
        "has_procedure_group", "gba_procedure_group", "condition_label",
        "orphan_flag",
        "_is_stopword", "_term_body", "_term_pattern", "_any_term_regex",
        "expand_condition_query", "molecule_set", "_inn_stem",
        "_indication_tokens", "indication_overlap", "_clean_inn",
    ]),
    "nice": ("hta_nice.py", [
        "_guidance_status_category", "calculate_similarity_score",
        "has_tag_coverage", "resolve_keyword", "classify_historical_icer",
        "format_historical_icer", "generate_assessment_pdf",
    ]),
    "chat": ("hta_chat.py", [
        "build_chat_context", "ask_chat", "research_questions_from_themes",
        "extract_severity", "synthesise_themes", "structure_reasoning_card",
        "build_concern_frequency", "split_shared_unique",
        "_anthropic_api_key", "_matches_by_record",
        "build_agency_chat_context", "render_agency_chat",
    ]),
    "ui": ("hta_ui.py", [
        "prepare_agency_frame", "cross_agency_index", "agency_match_table",
        "find_same_molecule", "_cross_agency_lines",
        "render_nice_cross_agency",
        "outcome_order", "init_range_state", "render_agency_switcher",
        "agency_sidebar_guide", "_explorer_key", "_clear_agency_search",
        "_counted", "render_agency_explorer", "run_agency_search",
        "_query_key", "_search_ranking", "cached_agency_search",
        "describe_query", "render_agency_record", "_block_subset",
        "render_synthesis_block", "render_economic_block", "_coverage_line",
        "_integer_axis", "_style_bars", "outcome_breakdown_table",
        "render_outcome_table", "render_agency_charts", "_agency_footer",
        "_records_matched", "_agency_export_bytes",
        "_supports_deferred_download", "render_download", "_is_narrowed",
        "render_results_header", "_section_nav", "_results_signature",
        "render_results_decisions", "render_results_evidence",
        "render_results_other_bodies", "render_agency_view",
    ]),
}

HEADERS = {
    "hta_data.py": "Loading and caching the workbook, its sheets and the "
                   "vocabulary.",
    "hta_core.py": "Reading an HTA decision: text handling, G-BA and HAS "
                   "outcomes, classification, and query expansion.",
    "hta_nice.py": "Everything NICE-specific: similarity scoring, ICER "
                   "classification, the PDF export and the NICE page.",
    "hta_chat.py": "Both chat implementations, their prompts and the "
                   "context they are given.",
    "hta_ui.py": "The agency screens: explorer, record view, charts, "
                 "results, export and cross-agency matching.",
}


def read(path):
    with open(path, encoding="utf-8", newline="") as fh:
        return fh.read()


def write(path, text):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="") as fh:
        fh.write(text)
    os.replace(tmp, path)


def names_read_by(node, tree_funcs):
    """Module-level names a definition reads, ignoring its own locals."""
    local = set()
    for inner in ast.walk(node):
        if isinstance(inner, ast.arg):
            local.add(inner.arg)
        elif isinstance(inner, ast.Name) and isinstance(inner.ctx, ast.Store):
            local.add(inner.id)
    out = set()
    for inner in ast.walk(node):
        if isinstance(inner, ast.Name) and isinstance(inner.ctx, ast.Load):
            if inner.id not in local:
                out.add(inner.id)
    return out - tree_funcs


def block_of(node, lines):
    """Source lines of a definition, decorators included."""
    start = min([d.lineno for d in getattr(node, "decorator_list", [])]
                + [node.lineno]) - 1
    end = node.end_lineno
    return start, end


def plan_group(src, wanted):
    """What moves: the definitions, plus constants only they use."""
    tree = ast.parse(src)
    lines = src.split("\n")
    funcs = {n.name: n for n in tree.body
             if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef,
                               ast.ClassDef))}
    missing = [w for w in wanted if w not in funcs]
    if missing:
        sys.exit(f"Not found in {TARGET}: {', '.join(missing)}")

    # Only a bare "NAME = ..." counts. Walking the target of
    # df["guidance_status"] = ... also finds a Name node for df, which made
    # the dataframe itself look like a movable constant.
    consts = {}
    for node in tree.body:
        if isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = ([node.target] if isinstance(node, ast.AnnAssign)
                       else node.targets)
            for t in targets:
                if isinstance(t, ast.Name):
                    consts[t.id] = node
                elif isinstance(t, (ast.Tuple, ast.List)):
                    for element in t.elts:
                        if isinstance(element, ast.Name):
                            consts[element.id] = node

    all_funcs = set(funcs)
    needed = set()
    for name in wanted:
        needed |= names_read_by(funcs[name], all_funcs - set(wanted))

    # A constant travels only if nothing staying behind still reads it.
    staying = [n for n in tree.body
               if not (isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef,
                                      ast.ClassDef)) and n.name in wanted)]
    read_by_stayers = set()
    for node in staying:
        if isinstance(node, (ast.Assign, ast.AnnAssign)) and \
                any(n.id in needed for n in ast.walk(node)
                    if isinstance(n, ast.Name)):
            pass
        for n in ast.walk(node):
            if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Load):
                read_by_stayers.add(n.id)

    # A constant the moved functions need always travels with them, even
    # when code left behind also reads it: app3 star-imports the new module,
    # so the name is still in scope there. Importing it the other way --
    # the new module reading it back from app3 -- would be circular.
    # A value whose definition reads something that is not itself a constant
    # cannot travel: "TOTAL_ROWS = len(df)" looks like a constant because of
    # its case, but it depends on the dataframe and moving it drags df into
    # the new module.
    def is_portable(node):
        for n in ast.walk(node):
            if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Load):
                if n.id in all_funcs:
                    continue
                if n.id in consts and is_portable_cache.get(n.id, True):
                    continue
                if n.id in globals_non_const:
                    return False
        return True

    globals_non_const = set()
    for node in tree.body:
        if isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = ([node.target] if isinstance(node, ast.AnnAssign)
                       else node.targets)
            for t in targets:
                for n in ast.walk(t):
                    if isinstance(n, ast.Name) and not n.id.isupper():
                        globals_non_const.add(n.id)
    is_portable_cache = {}
    for name, node in consts.items():
        is_portable_cache[name] = is_portable(node)

    # Follow constants that other constants reference, to a fixpoint.
    pending, chosen = set(needed), set()
    while pending:
        name = pending.pop()
        if name in chosen:
            continue
        node = consts.get(name)
        if node is None:
            continue
        chosen.add(name)
        for n in ast.walk(node):
            if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Load):
                if n.id in consts and n.id not in chosen:
                    pending.add(n.id)

    moving_consts, shared_consts, stuck = [], [], []
    for name in sorted(chosen):
        node = consts[name]
        if not is_portable_cache.get(name, True):
            stuck.append(name)
            continue
        moving_consts.append((name, node))
        if name in read_by_stayers:
            shared_consts.append((name, node))
    if stuck:
        print(f"  staying (depend on live data): {', '.join(stuck)}")

    blocks = []
    for name in wanted:
        blocks.append(block_of(funcs[name], lines))
    for _name, node in moving_consts:
        blocks.append(block_of(node, lines))
    return blocks, moving_consts, shared_consts, needed, lines


def imports_for(text, source):
    """Imports the moved code needs.

    Copied from the ORIGINAL file's import block rather than guessed: the
    first version of this listed a handful of modules by name and missed
    warnings, collections and the whole of reportlab, which meant the new
    modules compiled and then failed on first call.
    """
    wanted, seen = [], set()
    tree = ast.parse(source)
    for node in tree.body:
        if not isinstance(node, (ast.Import, ast.ImportFrom)):
            continue
        for alias in node.names:
            bound = (alias.asname or alias.name).split(".")[0]
            if bound in seen or not re.search(rf"\b{re.escape(bound)}\b", text):
                continue
            seen.add(bound)
            if isinstance(node, ast.Import):
                wanted.append(f"import {alias.name}"
                              + (f" as {alias.asname}" if alias.asname else ""))
            else:
                wanted.append(f"from {node.module} import {alias.name}"
                              + (f" as {alias.asname}" if alias.asname else ""))
    return wanted


OPTIONAL_COMPANIONS = ["source_citations", "endpoint_panel",
                       "evidence_filter", "nice_attributes", "nice_sections",
                       "committee_quotes"]


def optional_imports(text):
    """Guarded imports for the companion modules app3 loads defensively.

    These are optional by design -- the app runs without them -- so a module
    that calls into one needs the same try/except fallback rather than a
    hard import that would break the app when a file is missing.
    """
    out = []
    for name in OPTIONAL_COMPANIONS:
        if re.search(rf"\b{name}\.", text):
            out.append(
                f"try:\n    import {name}\nexcept ImportError:  "
                f"# optional companion module\n    {name} = None")
    return out


def sibling_imports(text, exclude):
    """Star-imports from the hta_ modules already written, in creation order.

    A moved function looks its names up in ITS OWN module's globals, so
    hta_gba calling _text() needs _text imported here -- app3 having
    star-imported both is not enough. Order is creation order, so a module
    only ever reads from ones that already exist and no cycle is possible.
    """
    out = []
    for mod in [m[:-3] for _key, (m, _n) in GROUPS.items()]:
        if mod + ".py" == exclude or not os.path.exists(mod + ".py"):
            continue
        names = set()
        for node in ast.parse(read(mod + ".py").replace("\r\n", "\n")).body:
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef,
                                 ast.ClassDef)):
                names.add(node.name)
            elif isinstance(node, (ast.Assign, ast.AnnAssign)):
                for n in ast.walk(node):
                    if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Store):
                        names.add(n.id)
        if any(re.search(rf"\b{re.escape(n)}\b", text) for n in names):
            out.append(f"from {mod} import *  # noqa: F401,F403")
    return out


def move_nice_page(src, crlf):
    """Wrap the trailing NICE page in a function and move it to hta_nice.py.

    The last ~1,265 lines of app3.py are the original NICE page running as
    straight-line module code. Nothing after it reads the names it assigns,
    so wrapping it in a function makes those locals and changes nothing --
    and it only reads one live value from outside, the dataframe, which
    becomes its single argument.
    """
    lines = src.split("\n")
    tree = ast.parse(src)
    start = None
    for node in tree.body:
        chunk = "\n".join(lines[node.lineno - 1:node.end_lineno])
        if "render_agency_view(HTA_BODY)" in chunk:
            start = node.end_lineno
    if start is None:
        sys.exit("Could not find the agency dispatch; is this app3.py?")
    block = lines[start:]
    if not any(l.strip() for l in block):
        print("Nothing after the dispatch — already moved?")
        return None, None

    body = "\n".join(("    " + l) if l.strip() else "" for l in block)
    # TOTAL_ROWS is derived from the dataframe, so it is recomputed inside
    # rather than carried across; VOCAB is live data and becomes the second
    # argument.
    func = ('\n\ndef render_nice_page(df, VOCAB):\n'
            '    """The NICE page. Was straight-line module code in app3.py."""\n'
            '    TOTAL_ROWS = len(df)\n'
            + body + "\n")
    remaining = "\n".join(lines[:start]) + "\n\nrender_nice_page(df, VOCAB)\n"
    return func, remaining


def verify_one(path):
    """Names a file reads that nothing defines or imports for it."""
    src = read(path).replace("\r\n", "\n")
    tree = ast.parse(src)
    defined = set(dir(builtins)) | {"__name__", "__file__"}
    for mod in re.findall(r"^from\s+(hta_\w+)\s+import\s+\*", src, re.M):
        if not os.path.exists(mod + ".py"):
            print(f"  MISSING MODULE: {mod}.py")
            continue
        mod_tree = ast.parse(read(mod + ".py").replace("\r\n", "\n"))
        # What a star-import ACTUALLY gives you: __all__ when declared,
        # otherwise only the names that do not start with an underscore.
        declared = None
        for node in mod_tree.body:
            if isinstance(node, ast.Assign) and any(
                    isinstance(t, ast.Name) and t.id == "__all__"
                    for t in node.targets):
                declared = [e.value for e in node.value.elts
                            if isinstance(e, ast.Constant)]
        if declared is not None:
            defined |= set(declared)
            continue
        for node in mod_tree.body:
            names = []
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef,
                                 ast.ClassDef)):
                names = [node.name]
            elif isinstance(node, (ast.Assign, ast.AnnAssign)):
                names = [n.id for n in ast.walk(node)
                         if isinstance(n, ast.Name)
                         and isinstance(n.ctx, ast.Store)]
            elif isinstance(node, (ast.Import, ast.ImportFrom)):
                names = [(a.asname or a.name).split(".")[0]
                         for a in node.names]
            defined |= {n for n in names if not n.startswith("_")}
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef,
                             ast.ClassDef)):
            defined.add(node.name)
            if not isinstance(node, ast.ClassDef):
                for arg in ast.walk(node.args):
                    if isinstance(arg, ast.arg):
                        defined.add(arg.arg)
        elif isinstance(node, ast.Lambda):
            for arg in ast.walk(node.args):
                if isinstance(arg, ast.arg):
                    defined.add(arg.arg)
        elif isinstance(node, (ast.Import, ast.ImportFrom)):
            for alias in node.names:
                defined.add((alias.asname or alias.name).split(".")[0])
        elif isinstance(node, ast.Name) and isinstance(node.ctx, ast.Store):
            defined.add(node.id)
        elif isinstance(node, ast.ExceptHandler) and node.name:
            defined.add(node.name)
        elif isinstance(node, ast.comprehension):
            for n in ast.walk(node.target):
                if isinstance(n, ast.Name):
                    defined.add(n.id)
    used = {n.id for n in ast.walk(tree)
            if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Load)}
    missing = sorted(x for x in used - defined if not x.startswith("__"))
    compile(src, path, "exec")
    mark = "OK " if not missing else "BAD"
    print(f"  {mark} {path:<24} {len(src.splitlines()):>5} lines  "
          f"{len(missing)} unresolved"
          + (f"  {', '.join(missing[:6])}" if missing else ""))
    return len(missing)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", default=TARGET)
    ap.add_argument("--group", choices=sorted(GROUPS))
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--verify", action="store_true")
    ap.add_argument("--revert", action="store_true")
    ap.add_argument("--page", action="store_true",
                    help="move the trailing NICE page into hta_nice.py")
    args = ap.parse_args()

    if args.list:
        for key, (module, names) in GROUPS.items():
            print(f"  {key:<10} -> {module:<18} {len(names)} definitions")
        return
    if args.revert:
        if not os.path.exists(BACKUP):
            sys.exit(f"No backup at {BACKUP}.")
        shutil.copy(BACKUP, args.file)
        print(f"Restored {args.file} from {BACKUP}.")
        return
    if not os.path.exists(args.file):
        sys.exit(f"{args.file} not found.")

    src = read(args.file)
    crlf = src.count("\r\n") > 100
    flat = src.replace("\r\n", "\n")

    if args.verify:
        # Every file, not just app3. The first version of this checked only
        # the file being cut FROM, reported zero problems, and shipped ten
        # modules that could not resolve their own calls.
        targets = [args.file] + sorted(
            f for f in os.listdir(".")
            if f.startswith("hta_") and f.endswith(".py"))
        worst = 0
        for path in targets:
            worst += verify_one(path)
        print(f"\nTOTAL unresolved across {len(targets)} file(s): {worst}")
        if worst:
            print("Any name here will raise NameError at runtime, not at "
                  "import.")
        return

    if False:
        tree = ast.parse(flat)
        defined = set(dir(__builtins__)) | {"__name__", "__file__"}
        # Every binding anywhere in the module, not just at the top level:
        # app3 defines fallback classes inside try/except blocks, and a
        # shallower walk reports all of them as unresolved.
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef,
                                 ast.ClassDef)):
                defined.add(node.name)
                if not isinstance(node, ast.ClassDef):
                    for arg in ast.walk(node.args):
                        if isinstance(arg, ast.arg):
                            defined.add(arg.arg)
            elif isinstance(node, (ast.Import, ast.ImportFrom)):
                for alias in node.names:
                    defined.add((alias.asname or alias.name).split(".")[0])
            elif isinstance(node, ast.Name) and isinstance(node.ctx, ast.Store):
                defined.add(node.id)
            elif isinstance(node, ast.Lambda):
                for arg in ast.walk(node.args):
                    if isinstance(arg, ast.arg):
                        defined.add(arg.arg)
            elif isinstance(node, ast.ExceptHandler) and node.name:
                defined.add(node.name)
            elif isinstance(node, ast.comprehension):
                for n in ast.walk(node.target):
                    if isinstance(n, ast.Name):
                        defined.add(n.id)
        # Follow the star-imports FIRST and read what each module defines,
        # so the names we just moved out are known before anything is
        # judged missing.
        star = re.findall(r"^from\s+(hta_\w+)\s+import\s+\*", flat, re.M)
        for mod in star:
            path = mod + ".py"
            if not os.path.exists(path):
                print(f"  MISSING MODULE: {path}")
                continue
            for node in ast.parse(read(path).replace("\r\n", "\n")).body:
                if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef,
                                     ast.ClassDef)):
                    defined.add(node.name)
                elif isinstance(node, (ast.Assign, ast.AnnAssign)):
                    for n in ast.walk(node):
                        if isinstance(n, ast.Name) and \
                                isinstance(n.ctx, ast.Store):
                            defined.add(n.id)

        unresolved = set()
        for node in tree.body:
            for n in ast.walk(node):
                if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Load):
                    unresolved.add(n.id)
        missing = sorted(x for x in unresolved - defined
                         if not x.startswith("__"))
        compile(flat, args.file, "exec")
        print(f"{args.file}: compiles, "
              f"{len(flat.splitlines()):,} lines")
        if star:
            print(f"  star-imports resolved from: {', '.join(star)}")
        print(f"  unresolved names: {len(missing)}")
        for name in missing[:20]:
            print(f"    {name}")
        return

    if args.page:
        if not os.path.exists("hta_nice.py"):
            sys.exit("Run --group nice first; the page joins that module.")
        func, remaining = move_nice_page(flat, crlf)
        if func is None:
            return
        target = read("hta_nice.py").replace("\r\n", "\n")
        # The page reads from every other module and from the original
        # import block; give it what it needs before merging.
        head = []
        for line in imports_for(func, flat):
            if line not in target:
                head.append(line)
        for line in sibling_imports(func, "hta_nice.py"):
            if line not in target:
                head.append(line)
        if head:
            marker = target.index("\n\n", target.index('"""', 4) + 3)
            target = (target[:marker] + "\n" + "\n".join(head)
                      + target[marker:])
        for name in OPTIONAL_COMPANIONS:
            if re.search(rf"\b{name}\.", func) and \
                    f"import {name}" not in target:
                target += (f"\ntry:\n    import {name}\nexcept ImportError:"
                           f"  # optional companion\n    {name} = None\n")
        if "__all__" in target:
            target = target.replace("__all__ = [\n",
                                    '__all__ = [\n    "render_nice_page",\n',
                                    1)
        merged = target + func
        compile(merged, "hta_nice.py", "exec")
        compile(remaining, args.file, "exec")
        if not os.path.exists(BACKUP):
            shutil.copy(args.file, BACKUP)
        write("hta_nice.py", merged.replace("\n", "\r\n") if crlf else merged)
        write(args.file, remaining.replace("\n", "\r\n") if crlf else remaining)
        print(f"  NICE page moved into hta_nice.py as render_nice_page(df)")
        print(f"  {args.file} is now {len(remaining.splitlines()):,} lines")
        return

    if not args.group:
        sys.exit("Pass --group (see --list), or --verify, or --revert.")

    module, wanted = GROUPS[args.group]
    blocks, moving, shared, needed, lines = plan_group(flat, wanted)

    taken = set()
    for start, end in blocks:
        taken |= set(range(start, end))
    moved_text = "\n".join(lines[i] for i in sorted(taken))
    remaining = "\n".join(l for i, l in enumerate(lines) if i not in taken)

    print(f"group '{args.group}' -> {module}")
    print(f"  definitions moving : {len(wanted)}")
    print(f"  constants moving   : {len(moving)}"
          + (f"  ({', '.join(n for n, _ in moving[:8])})" if moving else ""))
    if shared:
        print(f"  also used by app3  : {len(shared)} — "
              f"{', '.join(n for n, _ in shared[:8])}")
        print("     these move too; app3 reads them back via the "
              "star-import")
    print(f"  lines moving       : {len(taken):,}")
    print(f"  app3.py after      : {len(remaining.splitlines()):,} lines")

    header = (f'"""\n{HEADERS.get(module, "Extracted from app3.py.")}\n\n'
              f'Moved out of app3.py by split_app3.py. These functions reach '
              f'only\nconstants, never the dataframe or session state, which '
              f'is what made\nthe move mechanical.\n"""\n\n')
    body = header + "\n".join(imports_for(moved_text, flat)) + "\n"
    siblings = sibling_imports(moved_text, module)
    if siblings:
        body += "\n".join(siblings) + "\n"
    optional = optional_imports(moved_text)
    if optional:
        body += "\n" + "\n".join(optional) + "\n"
    body += "\n"
    # "from X import *" skips names starting with an underscore unless the
    # module declares __all__. This codebase is full of them (_text,
    # _shorten, _fold), so without this they silently never cross a module
    # boundary and fail at runtime, not at import.
    exported = []
    for node in ast.parse(moved_text).body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef,
                             ast.ClassDef)):
            exported.append(node.name)
        elif isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = ([node.target] if isinstance(node, ast.AnnAssign)
                       else node.targets)
            for t in targets:
                if isinstance(t, ast.Name):
                    exported.append(t.id)
    body += "__all__ = [\n"
    for name in sorted(set(exported)):
        body += f'    "{name}",\n'
    body += "]\n\n"
    body += moved_text.strip() + "\n"

    import_line = f"from {module[:-3]} import *  # noqa: F401,F403\n"

    if args.dry_run:
        try:
            compile(body, module, "exec")
            print(f"\n  {module} would compile ({len(body.splitlines()):,} "
                  f"lines)")
        except SyntaxError as exc:
            print(f"\n  {module} would NOT compile: line {exc.lineno}: "
                  f"{exc.msg}")
        print("\n(dry run — nothing written)")
        return

    compile(body, module, "exec")
    # Put the import after the existing import block. Found with ast, not a
    # regex: a parenthesised "from x import (\n  A,\n  B,\n)" matches an
    # import regex on its FIRST line, and inserting there lands in the
    # middle of the bracket.
    out_lines = remaining.split("\n")
    insert_at = 0
    for node in ast.parse(remaining).body:
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            insert_at = max(insert_at, node.end_lineno)
    out_lines.insert(insert_at, import_line.rstrip("\n"))
    rebuilt = "\n".join(out_lines)
    compile(rebuilt, args.file, "exec")

    if not os.path.exists(BACKUP):
        shutil.copy(args.file, BACKUP)
        print(f"\n  backed up to {BACKUP}")
    write(module, body.replace("\n", "\r\n") if crlf else body)
    write(args.file, rebuilt.replace("\n", "\r\n") if crlf else rebuilt)
    print(f"  wrote {module}")
    print(f"  {args.file} is now {len(rebuilt.splitlines()):,} lines")
    print("\nRestart Streamlit and click through a record before moving the "
          "next group.")


if __name__ == "__main__":
    main()
