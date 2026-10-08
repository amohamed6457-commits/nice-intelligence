"""
Evidence attributes for app3 — the 13 coded fields per NICE appraisal.

Drop this next to app3.py, alongside evidence_attributes.json (the output of
extract_attributes.py). It turns 485 appraisals x 13 attributes into sidebar
filters, table columns and a per-appraisal panel.

Deliberately named nice_attributes, not evidence_attributes: a module whose
name matches a data file or folder invites Python to import the wrong thing as
an empty namespace package, which is a bug that looks like a logic error.

Design notes:
  * Every value carries a confidence. The extractor ran twice per appraisal;
    where the two passes disagreed, the stored value is one of two readings,
    not a fact. The UI says so rather than hiding it - roughly 11% of cells
    are in that state, concentrated in the four stance fields.
  * Everything degrades quietly. No file, no matches, a corrupt file: the
    functions return empty results and the app behaves as it does today.
  * No pandas import at module level, so this file can be read and tested
    without Streamlit's environment.

Build the input first:
    python extract_attributes.py --runs 2
"""

import json
import os
import re

_HERE = os.path.dirname(os.path.abspath(__file__))

DATA_PATHS = [
    "evidence_attributes.json",
    os.path.join(_HERE, "evidence_attributes.json"),
    os.path.join("hta_index", "evidence_attributes.json"),
    os.path.join(_HERE, "hta_index", "evidence_attributes.json"),
    "evidence_attributes.json.gz",
    os.path.join(_HERE, "evidence_attributes.json.gz"),
]

# Order matters: this is the order they appear in the UI, grouped so the
# descriptive fields (what the evidence WAS) come before the stance fields
# (what the committee MADE of it). Mixing them confuses readers, because the
# two groups have different reliability and answer different questions.
DESCRIPTIVE = [
    ("trial_design", "Trial design"),
    ("direct_comparative_data", "Head-to-head data"),
    ("indirect_comparison", "Indirect comparison"),
    ("external_control", "External control arm"),
    ("os_maturity", "Overall survival maturity"),
    ("surrogate_primary", "Surrogate primary endpoint"),
    ("qol_source", "Utility source"),
    ("subgroup_central", "Subgroup-driven"),
    ("small_population", "Small population"),
]

STANCE = [
    ("qol_stance", "Quality of life"),
    ("extrapolation_stance", "Survival extrapolation"),
    ("generalisability_stance", "Generalisability to NHS"),
    ("comparator_stance", "Comparator choice"),
]

FIELDS = DESCRIPTIVE + STANCE
LABELS = dict(FIELDS)
STANCE_KEYS = {k for k, _ in STANCE}

# What a stance value means, for a tooltip. The committee's position, not the
# topic's presence - a distinction the codebook spends a paragraph on.
STANCE_HELP = {
    "not discussed": "the committee did not raise it",
    "accepted": "raised and settled without real objection",
    "contested": "the committee or EAG pushed back, but it was not decisive",
    "drove decision": "named as a reason for the outcome",
}

_CACHE = {"loaded": False, "by_ta": {}, "meta": {}, "error": None}


def _norm_ta(value):
    """'TA993', 'ta993', 993, 'TA993 (part)' -> '993'. None if not a TA id."""
    if value is None:
        return None
    s = str(value).strip().upper()
    m = re.search(r"TA[^0-9]{0,2}(\d+)", s) or re.fullmatch(r"\s*(\d+)\s*", s)
    if not m:
        return None
    return m.group(1).lstrip("0") or "0"


def load():
    """Load once. Returns {ta: record}; empty dict if unavailable."""
    if _CACHE["loaded"]:
        return _CACHE["by_ta"]
    _CACHE["loaded"] = True

    path = next((p for p in DATA_PATHS if os.path.exists(p)), None)
    if not path:
        _CACHE["error"] = "evidence_attributes.json not found"
        return {}

    try:
        if path.endswith(".gz"):
            import gzip
            with gzip.open(path, "rt", encoding="utf-8") as f:
                records = json.load(f)
        else:
            with open(path, encoding="utf-8") as f:
                records = json.load(f)
    except (json.JSONDecodeError, OSError, EOFError) as e:
        _CACHE["error"] = f"could not read {os.path.basename(path)}: {e}"
        return {}

    # The extractor writes a list. Tolerate a dict keyed by TA as well, since
    # that is the shape of its own progress file and someone will point this
    # at one by mistake.
    if isinstance(records, dict):
        records = list(records.values())
    if not isinstance(records, list):
        _CACHE["error"] = "unexpected file shape"
        return {}

    by_ta = {}
    for rec in records:
        if not isinstance(rec, dict):
            continue
        ta = _norm_ta(rec.get("ta"))
        if ta:
            by_ta[ta] = rec

    _CACHE["by_ta"] = by_ta
    # Provenance, if the run that produced this recorded any. Worth surfacing:
    # two files can look identical and have been produced under different
    # models, codebooks or run counts.
    meta_path = os.path.join(os.path.dirname(path) or ".", "extract_run_meta.json")
    if os.path.exists(meta_path):
        try:
            with open(meta_path, encoding="utf-8") as f:
                _CACHE["meta"] = json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return by_ta


def available():
    """True if any appraisal has coded attributes."""
    return bool(load())


def provenance():
    """What produced this data, as a short line, or '' if not recorded."""
    load()
    m = _CACHE.get("meta") or {}
    if not m:
        return ""
    bits = []
    if m.get("model"):
        bits.append(m["model"])
    if m.get("runs"):
        bits.append(f"{m['runs']} passes")
    if m.get("codebook_sha"):
        bits.append(f"codebook {m['codebook_sha']}")
    if m.get("finished"):
        bits.append(m["finished"][:10])
    line = " · ".join(bits)
    if m.get("mixed_configurations"):
        line += "  ⚠ mixed configurations"
    return line


def status():
    """One line for the UI: what attribute data is actually available."""
    idx = load()
    if not idx:
        return f"Evidence attributes: not available ({_CACHE['error'] or 'unknown'})"
    n_split = sum(1 for r in idx.values() if r.get("n_disputed"))
    line = (f"Evidence attributes: {len(idx)} appraisals × {len(FIELDS)} fields; "
            f"{n_split} appraisals have at least one split reading")
    prov = provenance()
    return f"{line}\n{prov}" if prov else line


def get(appraisal_id):
    """The record for one appraisal, or None."""
    ta = _norm_ta(appraisal_id)
    return load().get(ta) if ta else None


def value(appraisal_id, field):
    """The coded value, or None if absent."""
    rec = get(appraisal_id)
    return rec.get(field) if rec else None


def is_split(appraisal_id, field):
    """True if the two extraction passes disagreed on this cell."""
    rec = get(appraisal_id)
    return bool(rec and field in (rec.get("disputed") or {}))


def alternatives(appraisal_id, field):
    """Both readings for a split cell, best-supported first. [] if agreed."""
    rec = get(appraisal_id)
    if not rec:
        return []
    d = (rec.get("disputed") or {}).get(field)
    if not d:
        return []
    vals = [v for v in (d.get("values") or []) if v]
    chosen = d.get("chosen")
    # Chosen first, then the reading that lost, de-duplicated.
    ordered = ([chosen] if chosen else []) + [v for v in vals if v != chosen]
    return list(dict.fromkeys(ordered))


def confidence(appraisal_id):
    """(agreed_cells, total_cells) for one appraisal."""
    rec = get(appraisal_id)
    if not rec:
        return 0, 0
    n_disputed = len([f for f in (rec.get("disputed") or {}) if f in LABELS])
    return len(FIELDS) - n_disputed, len(FIELDS)


def coverage(appraisal_ids):
    """(n_with_attributes, n_total) for a result set — honest UI labelling."""
    idx = load()
    tas = [_norm_ta(a) for a in appraisal_ids]
    tas = [t for t in tas if t]
    if not idx:
        return 0, len(tas)
    return sum(1 for t in tas if t in idx), len(tas)


def options(field, appraisal_ids=None):
    """Values actually present for a field, so filters never offer a dead end.

    Restricted to `appraisal_ids` when given, which keeps the sidebar honest
    about the current result set rather than the whole corpus.
    """
    idx = load()
    if not idx:
        return []
    if appraisal_ids is None:
        recs = idx.values()
    else:
        tas = {t for t in (_norm_ta(a) for a in appraisal_ids) if t}
        recs = [idx[t] for t in tas if t in idx]
    seen = {r.get(field) for r in recs if r.get(field)}
    # Stance values read in escalating order; everything else alphabetical.
    if field in STANCE_KEYS:
        order = ["not discussed", "accepted", "contested", "drove decision"]
        return [v for v in order if v in seen] + sorted(seen - set(order))
    return sorted(seen)


def attach(frame, id_column="appraisal_id", prefix="ev_"):
    """Add one column per attribute to a DataFrame. Returns (frame, n_matched).

    Columns are prefixed so they cannot collide with the workbook's own, and
    a parallel <prefix><field>_split boolean marks the cells where the two
    passes disagreed. Unmatched appraisals get None, not a guessed value.
    """
    idx = load()
    if not idx or id_column not in frame.columns:
        return frame, 0

    out = frame.copy()
    tas = [_norm_ta(v) for v in out[id_column]]
    recs = [idx.get(t) if t else None for t in tas]

    for field, _ in FIELDS:
        out[prefix + field] = [r.get(field) if r else None for r in recs]
        out[prefix + field + "_split"] = [
            bool(r and field in (r.get("disputed") or {})) for r in recs
        ]
    out[prefix + "coded"] = [r is not None for r in recs]
    out[prefix + "n_split"] = [
        len([f for f in (r.get("disputed") or {}) if f in LABELS]) if r else 0
        for r in recs
    ]
    return out, sum(1 for r in recs if r is not None)


def apply_filters(frame, selections, prefix="ev_", exclude_split=False,
                  require_coded=True):
    """Filter a frame attached by attach().

    selections : {field: [allowed values]} — an empty or missing list means
                 "no constraint on this field".
    exclude_split : drop rows whose value for a CONSTRAINED field was split
                 between the two passes. Off by default, because silently
                 dropping uncertain rows makes a count look more precise than
                 it is; on when the user asks for only agreed readings.
    require_coded : drop rows with no attribute data at all, but only when at
                 least one filter is active. With no filters, an uncoded
                 appraisal is not being excluded by anything.
    """
    active = {f: list(v) for f, v in (selections or {}).items() if v}
    if not active:
        return frame
    out = frame
    if require_coded and prefix + "coded" in out.columns:
        out = out[out[prefix + "coded"]]
    for field, allowed in active.items():
        col = prefix + field
        if col not in out.columns:
            continue
        out = out[out[col].isin(allowed)]
        if exclude_split and col + "_split" in out.columns:
            out = out[~out[col + "_split"]]
    return out


def summarise(frame, field, prefix="ev_"):
    """[(value, n, n_split)] for one field, commonest first.

    n_split is how many of those rows were a split reading, so a chart or
    table can show the soft edge of its own bars.
    """
    col = prefix + field
    if col not in frame.columns:
        return []
    counts = {}
    split_col = col + "_split"
    has_split = split_col in frame.columns
    for i, v in enumerate(frame[col]):
        if v is None or (isinstance(v, float) and v != v):
            continue
        entry = counts.setdefault(v, [0, 0])
        entry[0] += 1
        if has_split and bool(frame[split_col].iloc[i]):
            entry[1] += 1
    # Commonest first, ties broken alphabetically so the order is stable
    # across reruns rather than following dict insertion.
    return sorted(((v, n, s) for v, (n, s) in counts.items()),
                  key=lambda r: (-r[1], r[0]))


# ---------------------------------------------------------------------------
# Streamlit rendering. Imported lazily so this module stays testable.
# ---------------------------------------------------------------------------

def sidebar_filters(appraisal_ids=None, key_prefix="w_nice_ev"):
    """Render the attribute filters in the sidebar. Returns (selections, opts).

    Returns ({}, {}) when no data is available, so the caller can call this
    unconditionally and get today's behaviour.
    """
    import streamlit as st

    if not available():
        return {}, {}

    selections = {}
    with st.sidebar.expander("🧬 Evidence attributes", expanded=False):
        n = len(load())
        st.caption(
            f"Coded from the committee's own discussion for {n} appraisals. "
            "Two independent passes; where they disagreed the value is one of "
            "two readings and is marked ~."
        )

        st.markdown("**What the evidence was**")
        for field, label in DESCRIPTIVE:
            opts = options(field, appraisal_ids)
            if not opts:
                continue
            picked = st.multiselect(label, opts, key=f"{key_prefix}_{field}",
                                    help="Leave empty to include all values.")
            if picked:
                selections[field] = picked

        st.markdown("**What the committee made of it**")
        st.caption("Stance fields are the least repeatable (~17% split). "
                   "Treat them as a starting point, not a verdict.")
        for field, label in STANCE:
            opts = options(field, appraisal_ids)
            if not opts:
                continue
            picked = st.multiselect(
                label, opts, key=f"{key_prefix}_{field}",
                help="; ".join(f"{v} — {STANCE_HELP[v]}"
                               for v in opts if v in STANCE_HELP) or None)
            if picked:
                selections[field] = picked

        exclude_split = st.checkbox(
            "Only cells both passes agreed on", key=f"{key_prefix}_agreed",
            help="Stricter and smaller. Off by default because dropping the "
                 "uncertain rows makes the remaining count look more precise "
                 "than the method supports.")

    return selections, {"exclude_split": exclude_split}


def render_panel(appraisal_id, compact=False):
    """Show one appraisal's attributes, with split cells marked."""
    import streamlit as st

    rec = get(appraisal_id)
    if not rec:
        # Silent. This is called for every record in every agency, most of
        # which are not NICE TAs, and "no attributes" on all of them is noise.
        return

    agreed, total = confidence(appraisal_id)
    st.caption(f"Evidence attributes — {agreed}/{total} cells agreed by both "
               f"extraction passes")

    for group, heading in ((DESCRIPTIVE, "Evidence"), (STANCE, "Committee stance")):
        st.markdown(f"**{heading}**")
        rows = []
        for field, label in group:
            v = rec.get(field)
            if not v:
                continue
            if field in (rec.get("disputed") or {}):
                alts = alternatives(appraisal_id, field)
                other = " / ".join(alts[1:]) or "—"
                rows.append((label, f"~ {v}", f"other pass read: {other}"))
            else:
                rows.append((label, v, ""))
        if not rows:
            st.caption("nothing coded")
            continue
        for label, val, note in rows:
            cols = st.columns([2, 2, 3] if not compact else [2, 3])
            cols[0].markdown(f"<small>{label}</small>", unsafe_allow_html=True)
            cols[1].markdown(f"**{val}**")
            if not compact and note:
                cols[2].markdown(f"<small><i>{note}</i></small>",
                                 unsafe_allow_html=True)

    if rec.get("primary_decision_driver"):
        st.markdown("**What drove it**")
        st.write(rec["primary_decision_driver"])


def render_status_caption():
    """A one-line provenance note, for the bottom of a page."""
    import streamlit as st
    if available():
        st.caption(status())
