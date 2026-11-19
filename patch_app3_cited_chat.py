"""
Make the G-BA / HAS chat answer from page-cited passages and show them.

Today render_agency_chat() grounds its answers on each record's SUMMARY
fields and asks the model to cite by appraisal id. So an answer can say "no
added benefit was proven [D-1313]" and the reader still has to open a
20-page resolution to find where that was decided.

This patch adds, around the existing call:

  1. An import for cited_answers, placed above the HTA body switcher for the
     same reason as source_citations — render_agency_view() runs at module
     level and is followed by st.stop(), so a name bound after it does not
     exist when a G-BA or HAS record renders.

  2. attach() in render_agency_view(), building the passage pool from the
     three sheets through the app's own load_agency_sheet().

  3. In render_agency_chat(): retrieval of the passages most likely to
     answer the question, restricted to the appraisals currently on screen;
     those passages appended to the context; the citation rules appended to
     the system prompt; and the cited passages rendered under the answer
     with their page links.

The answer itself is unchanged in how it is produced — same model, same
endpoint, same turn limit. What changes is what it is given and what is
shown afterwards.

Nothing the model says about its sources is trusted: cited_answers.audit()
checks every marker against the passages actually supplied and every
quotation against the passage it is attributed to, and the render reports
any that fail.

Idempotent. Backs up to app3.py.bak6. Compiles before writing.

  python patch_app3_cited_chat.py --dry-run
  python patch_app3_cited_chat.py
"""

import argparse
import os
import re
import shutil
import sys

MARK = "cited_answers"

IMPORT_ANCHOR = "DATA_FILE = "

IMPORT_BLOCK = '''# --- answers that cite the passage they came from (optional) ------------
# Above the HTA body switcher deliberately: render_agency_view() runs at
# module level and is followed by st.stop(), so anything imported below it
# is undefined by the time a G-BA or HAS record renders.
try:
    import cited_answers
    if not hasattr(cited_answers, "select_passages"):
        raise ImportError("cited_answers resolved to a namespace package")
except (ImportError, AttributeError):
    class _NoCited:
        CITED_ANSWER_RULES = ""

        @staticmethod
        def attach(*a, **k):
            return None

        @staticmethod
        def is_attached():
            return False

        @staticmethod
        def select_passages(*a, **k):
            return []

        @staticmethod
        def build_context(*a, **k):
            return ""

        @staticmethod
        def render_sources(*a, **k):
            return {}

    cited_answers = _NoCited()


'''

ATTACH_ANCHOR = "    doc_index = load_document_index(version)\n"

ATTACH_BLOCK = '''    # Searchable pool of page-cited passages for the chat below.
    cited_answers.attach(
        load_agency_sheet("GBA_Decision_Passages", version),
        load_agency_sheet("HAS_Economic_Pages", version),
        load_agency_sheet("Economic_Findings", version),
    )
'''

OLD_CALL = '''        context = build_agency_chat_context(agency, frame, matches)
        with st.chat_message("assistant"):
            with st.spinner(f"Checking the selected {agency} {cfg['noun']}..."):
                answer, error = ask_chat(api_key, context, history[:-1], question,
                                         system_prompt=cfg["chat_prompt"])
            if error:
                st.error(error)
                history.pop()  # don't count a failed turn
            else:
                st.markdown(answer)
                history.append({"role": "assistant", "content": answer})
                counter.caption(f"{len(history) // 2}/{CHAT_MAX_TURNS} questions used for "
                                f"this selection.")
'''

NEW_CALL = '''        context = build_agency_chat_context(agency, frame, matches)
        # Page-cited passages for the records on screen, chosen for this
        # question. These let the answer point at where in a resolution a
        # finding sits, instead of naming the resolution and leaving the
        # reader to search 20 pages for it.
        passages = cited_answers.select_passages(
            question, appraisal_ids=frame["appraisal_id"].astype(str).tolist())
        prompt = cfg["chat_prompt"]
        if passages:
            context = context + "\\n\\n" + cited_answers.build_context(passages)
            prompt = prompt + cited_answers.CITED_ANSWER_RULES
        with st.chat_message("assistant"):
            with st.spinner(f"Checking the selected {agency} {cfg['noun']}..."):
                answer, error = ask_chat(api_key, context, history[:-1], question,
                                         system_prompt=prompt)
            if error:
                st.error(error)
                history.pop()  # don't count a failed turn
            else:
                st.markdown(answer)
                # Every marker is checked against the passages actually
                # supplied, and every quotation against the passage it is
                # attributed to, before any of it is shown as a source.
                if passages:
                    cited_answers.render_sources(answer, passages)
                history.append({"role": "assistant", "content": answer})
                counter.caption(f"{len(history) // 2}/{CHAT_MAX_TURNS} questions used for "
                                f"this selection.")
'''


def _fail(msg):
    print(msg)
    sys.exit(1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--app", default="app3.py")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    if not os.path.exists(args.app):
        _fail(f"Not found: {args.app}")
    if not os.path.exists("cited_answers.py"):
        print("warning: cited_answers.py is not in this folder. The patch applies, "
              "but the chat keeps its current behaviour until you copy it in.\n")

    with open(args.app, "rb") as f:
        raw = f.read()
    crlf = b"\r\n" in raw
    text = raw.decode("utf-8")
    if crlf:
        text = text.replace("\r\n", "\n")

    if MARK in text:
        print(f"Already applied — {args.app} already wires in cited_answers.")
        return

    idx = text.find("\n" + IMPORT_ANCHOR)
    if idx == -1:
        _fail(f"Could not find the '{IMPORT_ANCHOR.strip()}' line in {args.app}.")
    text = text[:idx + 1] + IMPORT_BLOCK + text[idx + 1:]

    if text.count(ATTACH_ANCHOR) != 1:
        _fail(f"Expected one '{ATTACH_ANCHOR.strip()}' line, found "
              f"{text.count(ATTACH_ANCHOR)}. Not patching.")
    text = text.replace(ATTACH_ANCHOR, ATTACH_ANCHOR + ATTACH_BLOCK, 1)

    if text.count(OLD_CALL) != 1:
        _fail("Could not find the chat call in render_agency_chat() "
              f"(found {text.count(OLD_CALL)} matches).\nIf you have edited that "
              "function, send me the current version and I'll re-anchor.")
    text = text.replace(OLD_CALL, NEW_CALL, 1)

    try:
        compile(text, args.app, "exec")
    except SyntaxError as e:
        _fail(f"Patched source does not compile ({e}). {args.app} is untouched.")

    if args.dry_run:
        print("--dry-run: all three edits apply and the result compiles.")
        _report(text)
        return

    shutil.copy2(args.app, args.app + ".bak6")
    out = text.replace("\n", "\r\n") if crlf else text
    with open(args.app, "wb") as f:
        f.write(out.encode("utf-8"))
    print(f"Wrote {args.app}  (backup: {args.app}.bak6)")
    print("  + cited_answers import, above the HTA body switcher")
    print("  + attach() for the passage pool in render_agency_view")
    print("  + retrieval, citation rules and source rendering in render_agency_chat")
    _report(text)
    print("\nCopy cited_answers.py in beside app3.py, then restart Streamlit.")


def _report(text):
    lines = text.split("\n")
    imp = next((i for i, l in enumerate(lines, 1)
                if re.match(r"\s*import cited_answers\b", l)), None)
    sw = next((i for i, l in enumerate(lines, 1)
               if re.match(r"\s*HTA_BODY\s*=\s*render_agency_switcher\(\)", l)), None)
    if imp and sw:
        ok = "✓" if imp < sw else "⚠️"
        print(f"\n  {ok} cited_answers imports at line {imp}; switcher at line {sw}.")


if __name__ == "__main__":
    main()
