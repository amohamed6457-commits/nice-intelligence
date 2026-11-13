

import inspect
import io
import os
import re
import warnings
import requests
from collections import Counter
from datetime import date, datetime

import pandas as pd
import plotly.express as px
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    HRFlowable,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from hta_data import *  # noqa: F401,F403
from hta_core import *  # noqa: F401,F403
from hta_nice import *  # noqa: F401,F403
from hta_chat import *  # noqa: F401,F403
from hta_ui import *  # noqa: F401,F403

# Committee discussion sections for the NICE chat. Entirely optional: if the
# module or its index is absent - a deployment that did not ship them, say -
# the app must still run exactly as it did before, so the import cannot be
# allowed to fail hard.
try:
    import nice_sections
    # A directory of the same name on sys.path imports as an EMPTY namespace
    # package rather than raising, so a successful import is not proof the
    # real module loaded. Touch an attribute to find out.
    if not hasattr(nice_sections, "CITATION_RULES"):
        raise ImportError("nice_sections resolved to a namespace package")
except (ImportError, AttributeError):
    class _NoChapters:
        """Stand-in so a missing nice_sections.py cannot take the app down."""
        CITATION_RULES = ""

        @staticmethod
        def augment_context(context, *_args, **_kwargs):
            return context

        @staticmethod
        def available():
            return False

        @staticmethod
        def status():
            return "Committee sections: nice_sections.py not deployed"

    nice_sections = _NoChapters()

# --- committee quotes (optional) ---------------------------------------
# Checked by attribute, not by a bare import: a module whose name collides
# with a directory imports as an empty namespace package, which succeeds and
# then fails on first use.
try:
    import committee_quotes
    if not hasattr(committee_quotes, "render_theme_evidence"):
        raise ImportError("committee_quotes resolved to a namespace package")
except (ImportError, AttributeError):
    class _NoQuotes:
        @staticmethod
        def available():
            return False

        # Same strictness as the real module. A weak test here would make
        # the degraded path reproduce the exact bug this patch removes:
        # "Negative recommendation." quoted as committee reasoning.
        _STUB_RE = re.compile(
            r"^\W*((negative|positive|favourable|unfavourable)\s+recommendation"
            r"|not?\s+recommend(ed|ation)?|recommend(ed|ation)?"
            r"|(terminated|withdrawn|discontinued)(\s+\w+)?"
            r"|rejected|approved|restricted|see\s+(full\s+)?guidance"
            r"|n/?a|none|unknown"
            r"|not\s+(applicable|specified|recorded|stated|available))\W*$",
            re.IGNORECASE)

        @staticmethod
        def is_stub(text):
            if text is None:
                return True
            s = str(text).strip()
            if not s or s.lower() == "nan":
                return True
            # Over ~90 chars it is a sentence, not a label.
            return len(s) <= 90 and bool(_NoQuotes._STUB_RE.match(s))

        @staticmethod
        def _trim(text, limit):
            # Never mid-word: a quote ending "were not acce" reads as a bug
            # even when the underlying text is complete.
            s = str(text).strip()
            if len(s) <= limit:
                return s
            w = s[:limit]
            end = max(w.rfind(". "), w.rfind("? "), w.rfind("! "))
            if end > limit * 0.6:
                return w[: end + 1]
            return w.rsplit(" ", 1)[0].rstrip(",;:") + " …"

        @staticmethod
        def render_theme_evidence(appraisal_id, theme_label,
                                  fallback_text=None, fallback_chars=1600,
                                  **kw):
            import streamlit as _st
            _st.markdown(f"**{appraisal_id}**")
            if fallback_text and not _NoQuotes.is_stub(fallback_text):
                _st.markdown(f"> {_NoQuotes._trim(fallback_text, fallback_chars)}")
                return True
            _st.caption("No committee text available for this theme.")
            return False

    committee_quotes = _NoQuotes()

# --- evidence attributes (optional) ------------------------------------
# A module whose name matches a data file or directory can import as an empty
# namespace package, which succeeds and then fails on first use. So check for
# a known attribute rather than trusting the import.
# NB: must not be written as a bare expression - Streamlit's magic would
# render it to the page.
try:
    import nice_attributes
    if not hasattr(nice_attributes, "FIELDS"):
        raise ImportError("nice_attributes resolved to a namespace package")
except (ImportError, AttributeError):
    class _NoAttributes:
        FIELDS = []

        @staticmethod
        def available():
            return False

        @staticmethod
        def attach(frame, *a, **k):
            return frame, 0

        @staticmethod
        def sidebar_filters(*a, **k):
            return {}, {}

        @staticmethod
        def apply_filters(frame, *a, **k):
            return frame

        @staticmethod
        def render_panel(*a, **k):
            return None

        @staticmethod
        def render_status_caption(*a, **k):
            return None

        @staticmethod
        def coverage(ids):
            return 0, len(list(ids))

    nice_attributes = _NoAttributes()

# --- page-level source citation for G-BA / HAS (optional) ---------------
# Imported HERE, above the HTA body switcher, and not with the other optional
# modules further down: render_agency_view() runs at module level and is
# followed by st.stop(), so a name bound after that point does not exist when
# a G-BA or HAS record renders. A module whose name matches a data file or
# directory can also import as an empty namespace package, which succeeds and
# then fails on first use, so check for a known attribute rather than
# trusting the import.
try:
    import source_citations
    if not hasattr(source_citations, "render_panel"):
        raise ImportError("source_citations resolved to a namespace package")
except (ImportError, AttributeError):
    class _NoCitations:
        @staticmethod
        def attach(*a, **k):
            return None

        @staticmethod
        def is_attached():
            return False

        @staticmethod
        def has_sources(*a, **k):
            return False

        @staticmethod
        def coverage():
            return {}

        @staticmethod
        def render_panel(*a, **k):
            return False

    source_citations = _NoCitations()


# --- filter by whether anything was assessable (optional) ---------------
try:
    import evidence_filter
    if not hasattr(evidence_filter, "ids_for"):
        raise ImportError("evidence_filter resolved to a namespace package")
except (ImportError, AttributeError):
    class _NoEvidenceFilter:
        @staticmethod
        def attach_gba(*a, **k):
            return None

        @staticmethod
        def attach_has(*a, **k):
            return None

        @staticmethod
        def options_for(*a, **k):
            return []

        @staticmethod
        def counts_for(*a, **k):
            return {}

        @staticmethod
        def ids_for(*a, **k):
            return None

        @staticmethod
        def help_for(*a, **k):
            return ""

    evidence_filter = _NoEvidenceFilter()


# --- endpoint verdicts, with the resolution page behind them (optional) ---
# Same guard as above: imported before the HTA body switcher, and checked
# with hasattr because a namespace package imports fine and fails on use.
try:
    import endpoint_panel
    if not hasattr(endpoint_panel, "render"):
        raise ImportError("endpoint_panel resolved to a namespace package")
except (ImportError, AttributeError):
    class _NoEndpoints:
        @staticmethod
        def attach(*a, **k):
            return None

        @staticmethod
        def is_attached():
            return False

        @staticmethod
        def render(*a, **k):
            return None

    endpoint_panel = _NoEndpoints()









st.set_page_config(
    page_title="NICE Intelligence Dashboard",
    page_icon="💊",
    layout="wide",
)

# Keep filters and Explorer inputs when switching HTA body. Streamlit deletes a
# widget's value at the end of any run in which that widget wasn't drawn, so
# NICE → G-BA → NICE would otherwise wipe every NICE input — and, because
# hta_query_ran survives, land on a bare "Please enter an indication" warning
# instead of the results that were on screen. Re-saving keyed values at the top
# of each run is Streamlit's documented way to keep them. Only value widgets use
# the "w_" key prefix: buttons and chat inputs can't be written this way. Their
# defaults are set through session state rather than value=/index= arguments,
# which is what lets this run without widget-state warnings.
for _state_key in [k for k in st.session_state.keys() if str(k).startswith("w_")]:
    st.session_state[_state_key] = st.session_state[_state_key]


# Data loading









VOCAB = load_vocabulary(data_version())




df["guidance_status"] = df["current_guidance_status"].apply(_guidance_status_category)


# Similarity scoring








# Keyword resolution





# Grounded chat over the retrieved comparator set



# Extends rule 4 (cite by TA ID) down to section level when committee
# discussion sections are present in the context.
CHAT_SYSTEM_PROMPT += nice_sections.CITATION_RULES










# Reasoning synthesis
















# PDF report



# ═════════════════════════════════════════════════════════════════════════════
# Multi-agency support — G-BA (Germany) and HAS (France)
#
# G-BA and HAS don't answer NICE's question. The G-BA rates added clinical
# benefit against a comparator it sets itself (AMNOG); HAS rates clinical
# benefit (SMR/ASMR) and decides early access. Neither produces an ICER in
# these records, so their views are built around their own outcome scales
# rather than squeezed into the NICE ICER-vs-threshold frame. Everything below
# is driven by AGENCY_CONFIG, so loading more rows into the G-BA_Recent /
# HAS_Recent sheets needs no code changes.
# ═════════════════════════════════════════════════════════════════════════════







# ── Text helpers ─────────────────────────────────────────────────────────────

# Notes that are useful on one record but say nothing across a set — English and
# French forms, since HAS rows may be entered either way.




















# ── G-BA outcome scale ───────────────────────────────────────────────────────

# Procedures that haven't produced an extent rating, named for what they are
# instead of all landing in 'Unclassified'. Keys are the G-BA's own
# procedure_status wording (Verfahrensstand), lower-cased.
# Reserve antibiotics (§35a(1c) SGB V): the added benefit is deemed proven by law
# and no extent is rated.
























# ── HAS outcome scale ────────────────────────────────────────────────────────





















# ── Outcome families (charts) ────────────────────────────────────────────────
# Each agency's full outcome scale has 12–13 classes — past the ~7 at which
# colour can carry meaning, so the full scale is shown as a table and the
# charts group it into a few families. Colours are validated for colour-vision
# separation against Streamlit's own light (#ffffff) and dark (#0e1117)
# backgrounds, with a separate dark-mode step for each hue; 'no rating' is a
# recessive grey rather than a hue.

# ASMR I–IV is an ordered scale: one green ramp, strongest step = ASMR I.








# ── Harmonised therapy areas ─────────────────────────────────────────────────
# One English vocabulary for all three bodies, close to NICE's own, so a
# checklist tick means the same thing on every page. G-BA files each
# assessment under a German Therapiegebiet; HAS under one or more French
# 'System/Specific' pairs, with cancers filed inside their body system (breast
# cancer under 'Sein') — so any HAS pair naming a cancer is read as Oncology
# first. HAS rows with no category (mostly pre-2011) fall back to the ATC code.







# English labels already in the workbook (early tester rows), by first word.
# First matching ATC prefix wins, so the specific groups come before their letter.






# ── Procedure types, condition labels, orphan status ─────────────────────────











# ── Search: English terms → the German / French the records are written in ───
# G-BA and HAS describe conditions in German and French, so an English search
# for 'breast cancer' found 3 of the ~50 G-BA breast-cancer assessments
# ('Mammakarzinom'). Each entry has triggers (what a user might type — any
# language, abbreviations included) and the terms searched for in each
# language. The longest trigger in a query wins and consumes its words, so
# 'chronic lymphocytic leukaemia' is the CLL entry, not CLL plus every
# leukaemia, and 'small cell lung cancer' isn't swallowed by 'lung'. Terms
# are accent-free (_fold) and match at a word start ('myelom' also finds
# 'Myeloms'); terms of four letters or fewer must be whole words; a term
# starting 're:' is a regular expression (used to keep 'non-Hodgkin' out of
# Hodgkin, or pulmonary hypertension out of hypertension). Add rows freely.

# Qualifiers that narrow an English search but don't name a condition; they
# can't be matched against German or French text, so they only rank results.
# Plurals are caught by _is_stopword.














# ── Methodology text & chat prompts ──────────────────────────────────────────








# ── Agency configuration ─────────────────────────────────────────────────────



# ── Same-molecule matching across agencies ───────────────────────────────────

# Words after which a drug-name part stops naming the drug ('Mepolizumab as an
# add-on…', 'Exenatide … for injection').










# ── Frames ───────────────────────────────────────────────────────────────────

# Columns every agency view reads. Missing ones are added empty, so renaming or
# dropping a column in the workbook degrades gracefully instead of crashing a
# view — including the NICE view, which reads these rows for "Also assessed by".
























# ── Sidebar ──────────────────────────────────────────────────────────────────







# ── HTA Evidence Explorer (search) ───────────────────────────────────────────





















# ── Record detail & synthesis ────────────────────────────────────────────────











# ── Charts ───────────────────────────────────────────────────────────────────













# ── Grounded chat over the current selection ─────────────────────────────────









# ── The agency page ──────────────────────────────────────────────────────────





























# ── HTA body switcher ────────────────────────────────────────────────────────
# Rendered first so it sits at the top of the sidebar. The NICE page below is
# the original script; the other agencies render their own page and stop the
# run here, so none of the NICE code executes for them.

HTA_BODY = render_agency_switcher()
if HTA_BODY != "NICE":
    render_agency_view(HTA_BODY)
    st.stop()

render_nice_page(df, VOCAB)
