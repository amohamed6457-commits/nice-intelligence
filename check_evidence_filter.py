"""
Why isn't the "No usable data in" chip row showing up?

Checks every link in the chain and names the broken one.
Run from the folder you start streamlit from.

    python check_evidence_filter.py
"""

import os
import sys

OK, BAD, WARN = "  OK  ", " FAIL ", " WARN "


def say(state, label, detail=""):
    print(f"[{state}] {label}" + (f"\n         {detail}" if detail else ""))


def main():
    print(f"running in: {os.getcwd()}\n")
    problems = []

    if not os.path.exists("app3.py"):
        say(BAD, "app3.py is not in this folder")
        return
    say(OK, "app3.py found here")

    if not os.path.exists("evidence_filter.py"):
        say(BAD, "evidence_filter.py is missing", "The download did not land.")
        return
    say(OK, "evidence_filter.py found here")

    src = open("evidence_filter.py", encoding="utf-8").read()
    if "def label_for" in src:
        say(OK, "evidence_filter.py is the per-category version")
    else:
        say(BAD, "evidence_filter.py is the OLD aggregate version",
            "Overwrite it with the newer download.")
        problems.append("old evidence_filter.py")

    app = open("app3.py", encoding="utf-8", newline="").read()
    tokens = {
        "import evidence_filter": "guarded import",
        '"orphan", "evidence")': "EXPLORER_FIELDS includes evidence",
        "evidence_filter.attach_gba(": "attach call",
        "evidence_filter.options_for(agency)": "chip row",
        "evidence_filter.ids_for(": "filter applied",
    }
    for token, label in tokens.items():
        if token in app:
            say(OK, f"app3.py has the {label}")
        else:
            say(BAD, f"app3.py is MISSING the {label}",
                "Run: python patch_app3_evidence.py")
            problems.append(label)

    if "evidence_filter.attach_gba(" in app and "options_for(agency)" in app:
        attach_at = app.index("evidence_filter.attach_gba(")
        form_at = app.index("evidence_filter.options_for(agency)")
        call_at = app.find("render_agency_explorer(agency, frame)",
                           app.index("def render_agency_explorer") + 40)
        if attach_at < call_at:
            say(OK, "attach() runs before the form is drawn")
        else:
            say(BAD, "attach() runs AFTER the form",
                "The chip row would be empty on first render.")
            problems.append("attach ordering")
        del form_at

    # Does the module actually produce options from the real workbook?
    try:
        import pandas as pd
        sys.path.insert(0, os.getcwd())
        import evidence_filter

        books = sorted(f for f in os.listdir(".")
                       if f.endswith(".xlsx") and "NICE" in f)
        if not books:
            say(WARN, "no workbook here to test against")
            return
        book = books[-1]
        try:
            subs = pd.read_excel(book, sheet_name="GBA_Subgroups")
        except Exception as exc:                           # noqa: BLE001
            say(BAD, f"cannot read GBA_Subgroups from {book}", str(exc))
            problems.append("GBA_Subgroups unreadable")
            subs = None
        if subs is not None:
            say(OK, f"GBA_Subgroups readable from {book}: {len(subs):,} rows")
            evidence_filter.attach_gba(subs)
            options = evidence_filter.options_for("G-BA")
            if options:
                say(OK, f"G-BA chips would be: {', '.join(options)}")
                print("         counts:", evidence_filter.counts_for("G-BA"))
            else:
                say(BAD, "attach_gba() produced no options",
                    "The verdict classifier matched nothing - is "
                    "endpoint_summary.py the current version?")
                problems.append("no options produced")
        try:
            rat = pd.read_excel(book, sheet_name="HAS_Ratings")
            evidence_filter.attach_has(rat)
            say(OK, "HAS chips would be: "
                    + ", ".join(evidence_filter.options_for("HAS")))
        except Exception as exc:                           # noqa: BLE001
            say(WARN, f"HAS_Ratings not testable: {exc}")
    except Exception as exc:                               # noqa: BLE001
        say(BAD, f"could not exercise evidence_filter: "
                 f"{type(exc).__name__}: {exc}")
        problems.append("module raised")

    print()
    if problems:
        print("VERDICT: " + "; ".join(problems))
    else:
        print("VERDICT: everything checks out. If the chip row is still "
              "missing,\n         Streamlit was not restarted - Ctrl+C, then "
              "streamlit run app3.py.\n         The row sits under 'Procedure "
              "type', above the year slider,\n         and shows only on G-BA "
              "and HAS - never on NICE.")


if __name__ == "__main__":
    main()
