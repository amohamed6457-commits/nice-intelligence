"""
Loading and caching the workbook, its sheets and the vocabulary.

Moved out of app3.py by split_app3.py. These functions reach only
constants, never the dataframe or session state, which is what made
the move mechanical.
"""

import os
import re
import pandas as pd
import streamlit as st
from hta_core import *  # noqa: F401,F403

DATA_FILE = "NICE_HAS_GBA_updated_2026-11-07.xlsx"
def data_version():
    """
    Cache key for the workbook: its path and last-modified time. Every loader
    below takes this as an argument, so replacing the workbook — or pointing
    DATA_FILE at a new one — refreshes every view on the next run instead of
    serving rows cached from the old file.
    """
    try:
        return (DATA_FILE, os.path.getmtime(DATA_FILE))
    except OSError:
        return (DATA_FILE, None)
@st.cache_resource
def _workbook(version):
    """
    Open the workbook once per (path, mtime) and hand back the same ExcelFile
    for every sheet. Each independent pd.read_excel(path, sheet_name=...) call
    re-parses the *entire* file — harmless at a few hundred rows, but at the
    full HAS/G-BA catalogue (12k+ and 1k+ rows, a 40k-row document index) six
    separate full-file parses is where a cold load stops feeling instant.
    Reading every sheet off one shared handle turns six parses of the whole
    file into one open plus six cheap per-sheet parses.

    python-calamine (if installed) is roughly an order of magnitude faster
    at opening a large .xlsx than the default openpyxl engine; if it isn't
    available, this falls back to the default engine and still gets the
    single-open benefit.
    """
    path = version[0]
    try:
        return pd.ExcelFile(path, engine="calamine")
    except (ImportError, ValueError):
        return pd.ExcelFile(path)
def _read_sheet(version, sheet_name):
    return _workbook(version).parse(sheet_name)
@st.cache_data
def load_data(version):
    return _read_sheet(version, "Sheet1")
@st.cache_data
def load_enrichment_log(version):
    try:
        return _read_sheet(version, "Enrichment_Log")
    except Exception:
        return None
@st.cache_data
def load_vocabulary(version):
    """
    Build dropdown options from the Tag_Vocabulary sheet.

    The previous build hardcoded these lists in the selectbox calls, which
    silently drifted from the data: 85 of 137 tagged appraisals carried a
    value no user could select. Because the similarity scorer treats a
    value mismatch as a MISS (not an exclusion), those rows were being
    actively penalised — e.g. every HER2-positive breast appraisal scored
    zero on biomarker because only 'HER2 mutation-positive' was offered.
    Driving the options off the vocabulary sheet makes that drift
    impossible by construction.
    """
    try:
        vocab = _read_sheet(version, "Tag_Vocabulary")
    except Exception:
        return {}

    options = {}
    for field, group in vocab.groupby("Field"):
        values = [
            str(v).strip()
            for v in group["Allowed value"].dropna().tolist()
            if str(v).strip() and str(v).strip().lower() != "nan"
        ]
        values = [v for v in dict.fromkeys(values) if v.lower() != "not specified"]
        options[field] = ["Not specified"] + sorted(values, key=str.lower)
    return options
def vocab_options(vocab, field, fallback):
    return vocab.get(field, fallback)
DOC_INDEX_COLUMNS = ["Institution", "Appraisal_ID", "Document_Type",
                     "Document_Title", "Date", "Source_URL", "Notes"]
@st.cache_data
def load_agency_sheet(sheet_name, version):
    """Raw rows for a non-NICE agency. A missing sheet yields an empty frame, so
    an older workbook still opens — the NICE view never depends on these."""
    try:
        return _read_sheet(version, sheet_name)
    except Exception:
        return pd.DataFrame()
@st.cache_data
def load_document_index(version):
    """Appraisal_Documents: one row per official source document, all agencies."""
    try:
        docs = _read_sheet(version, "Appraisal_Documents")
    except Exception:
        return pd.DataFrame(columns=DOC_INDEX_COLUMNS)
    for col in ("Institution", "Appraisal_ID"):
        docs[col] = docs[col].astype(str).str.strip()
    return docs
