"""
The endpoint-findings panel: four verdicts, each with its page and a check.

What a reader sees
------------------
For every subgroup, the four G-BA endpoint categories, each showing:

  * the verdict, as a coloured word rather than a sentence to read;
  * the agency's own wording, in German or English per the language switch;
  * a link to the page of the resolution where the summary table is printed;
  * whether that table AGREES with the AIS record it was read from.

Why the check matters
---------------------
The verdicts come from G-BA's AIS database and were never on a page. Every
resolution also prints its own endpoint-category summary table, which IS on
a page and is the source AIS was derived from. subgroup_sources parses that
table independently, so where the two agree a reader can trust the row and
open the page. Measured across 672 single-subgroup appraisals, they agree on
97.3% of 1,641 category verdicts.

Honest about what it cannot do
------------------------------
  * The table is findable for about 62% of appraisals. For the rest the row
    still shows the verdict, labelled as an AIS record with no page anchor.
  * A resolution with several subgroups prints one table PER subgroup. This
    module finds the appraisal's table but does not yet resolve WHICH
    subgroup it belongs to, so it offers the page without claiming
    agreement. Claiming it would be unearned.

app3 supplies the frames, the same way it does for source_citations:

    endpoint_panel.attach(subgroups_frame, passages_frame)
    endpoint_panel.render("G-BA", appraisal_id)
"""

import os

import streamlit as st

import endpoint_summary
import endpoint_tables
import source_citations
import subgroup_sources

_SUBGROUPS = None
_PASSAGES = None
_RECORDS = None
_INDEX = None          # built once, on first render that needs it
_FALLBACK = None       # appraisal_id -> (decision-section page, any page, url)

# How each verdict reads at a glance. Streamlit colour names only.
TONE_COLOUR = {
    "gap": "orange",
    "positive": "green",
    "negative": "red",
    "neutral": "gray",
    "mixed": "violet",
    "unknown": "gray",
}


def attach(subgroups, passages, records=None):
    """Hand over the frames. Any may be empty; the panel then renders
    nothing rather than failing. `records` is the G-BA_Recent sheet, used
    only to offer a document link for appraisals with no indexed page."""
    global _SUBGROUPS, _PASSAGES, _RECORDS, _INDEX, _FALLBACK
    _SUBGROUPS, _PASSAGES, _RECORDS = subgroups, passages, records
    _INDEX = _FALLBACK = None


def is_attached():
    return _SUBGROUPS is not None and len(_SUBGROUPS) > 0


def _index():
    """Parse every locatable endpoint-category table, once per process."""
    global _INDEX
    if _INDEX is not None:
        return _INDEX
    _INDEX = {}
    if _PASSAGES is None or not len(_PASSAGES):
        return _INDEX
    try:
        for r in _PASSAGES.itertuples():
            text = str(getattr(r, "source_text", "") or "")
            if not subgroup_sources.TABLE_ANCHOR.search(text):
                continue
            table = subgroup_sources.parse_table(text)
            if not table:
                continue
            key = str(getattr(r, "appraisal_id", "")).strip()
            page = getattr(r, "document_page", None)
            prev = _INDEX.get(key)
            if prev and len(prev["table"]) >= len(table):
                continue
            try:
                page = int(page)
            except (TypeError, ValueError):
                page = None
            _INDEX[key] = {"page": page,
                           "url": getattr(r, "source_url", None),
                           "table": table}
    except Exception:                                      # noqa: BLE001
        # A panel is never worth crashing the record view for.
        _INDEX = {}
    return _INDEX


def _fallbacks():
    """Where to send a reader when the summary table cannot be parsed.

    Two weaker claims, in order. The decision section's first page is where
    the resolution sets out its decision -- not where this verdict is
    printed, and the panel says so. Failing that, any indexed page of the
    resolution. These are NOT citations of the verdict and must never be
    labelled as though they were.
    """
    global _FALLBACK
    if _FALLBACK is not None:
        return _FALLBACK
    _FALLBACK = {}
    try:
        if _PASSAGES is not None and len(_PASSAGES):
            frame = _PASSAGES.dropna(subset=["document_page"])
            ids = frame["appraisal_id"].astype(str).str.strip()
            for key, group in frame.groupby(ids):
                decision = group[group["section"] == "Decision / subgroup findings"]
                pick = decision if len(decision) else group
                row = pick.loc[pick["document_page"].idxmin()]
                _FALLBACK[key] = {
                    "page": int(row["document_page"]),
                    "url": row.get("source_url"),
                    "kind": "decision" if len(decision) else "nearest",
                }
        if _RECORDS is not None and len(_RECORDS):
            ids = _RECORDS["appraisal_id"].astype(str).str.strip()
            for key, url in zip(ids, _RECORDS.get("official_source_url", [])):
                if key not in _FALLBACK and str(url or "").startswith("http"):
                    _FALLBACK[key] = {"page": None, "url": url, "kind": "document"}
    except Exception:                                      # noqa: BLE001
        pass
    return _FALLBACK


_TABLE_DATA = False     # False = not tried yet, None = unavailable


def _tables():
    """The recovered endpoint tables, loaded once. Absent file is not an
    error: the panel simply offers no expanders."""
    global _TABLE_DATA
    if _TABLE_DATA is not False:
        return _TABLE_DATA
    _TABLE_DATA = None
    try:
        import json
        if os.path.exists(endpoint_tables.TABLES_FILE):
            with open(endpoint_tables.TABLES_FILE, encoding="utf-8") as fh:
                loaded = json.load(fh)
            if isinstance(loaded, dict):
                _TABLE_DATA = loaded
    except (ValueError, OSError):
        _TABLE_DATA = None
    return _TABLE_DATA


def _render_submitted(appraisal_id, category, wording, page_url_for):
    """The expander: what was actually submitted for this category.

    Shown even when the verdict is "no usable data", because that is the
    common case and the interesting one -- G-BA frequently receives data and
    rules it nicht bewertbar, usually because the study was single-arm or
    the comparator arm was not the one it specified. The committee's own
    wording goes at the top so the tables and the verdict read as one
    statement rather than as a contradiction.
    """
    data = _tables()
    if not data:
        return
    try:
        blocks = endpoint_tables.for_category(data, appraisal_id, category)
    except Exception:                                      # noqa: BLE001
        return
    if not blocks:
        return

    rows_total = sum(len(b["rows"]) for b in blocks)
    with st.expander(f"Show the {rows_total} row(s) submitted for "
                     f"{category.lower()}", expanded=False):
        if wording:
            st.caption(f"The committee's conclusion on this evidence: "
                       f"“{str(wording)[:300]}”")
        import pandas as pd
        for block in blocks:
            header = [str(c or "") for c in (block.get("header") or [])]
            body = [[str(c or "") for c in row] for row in block["rows"]]
            width = max([len(r) for r in body] + [len(header)])
            body = [r + [""] * (width - len(r)) for r in body]
            if len(header) < width:
                header = header + [""] * (width - len(header))
            caption = []
            if block.get("heading"):
                caption.append(f"**{block['heading']}**")
            link = page_url_for(block.get("page"))
            if link:
                caption.append(f"[page {block['page']} ↗]({link})")
            if caption:
                st.markdown(" &nbsp;·&nbsp; ".join(caption))
            try:
                st.dataframe(pd.DataFrame(body, columns=header),
                             width="stretch", hide_index=True)
            except Exception:                              # noqa: BLE001
                st.table(body)
        st.caption("Recovered from the resolution PDF. A table here does not "
                   "mean the evidence was accepted — read the committee's "
                   "conclusion above.")


def _rows_for(appraisal_id):
    key = str(appraisal_id or "").strip()
    if not key or not is_attached():
        return []
    frame = _SUBGROUPS
    ids = frame["appraisal_id"].astype(str).str.strip()
    return list(frame[ids == key].itertuples())


def render(agency, appraisal_id):
    """The panel. Silent unless this is a G-BA record with subgroup data."""
    if agency != "G-BA" or not is_attached():
        return
    rows = _rows_for(appraisal_id)
    if not rows:
        return

    found = _index().get(str(appraisal_id).strip())
    multi = len(rows) > 1

    st.subheader("Endpoint findings")
    st.caption(
        "What the G-BA concluded for each endpoint category. "
        + ("Cross-checked against the resolution's own summary table."
           if found else
           "No endpoint summary table could be located in this resolution, "
           "so these carry no page anchor."))

    for row in rows:
        if multi:
            label = source_citations.display_text(
                getattr(row, "patient_subgroup", None)) or "Subgroup"
            st.markdown(f"**{str(label)[:170]}**")

        for field, name in endpoint_summary.CATEGORIES:
            raw = getattr(row, field, None)
            label = endpoint_summary.verdict(raw)
            if not label:
                continue
            colour = TONE_COLOUR.get(endpoint_summary.TONE.get(label, "unknown"),
                                     "gray")
            plain = endpoint_summary.PLAIN.get(label, label)
            wording = source_citations.display_text(raw)

            bits = []
            if found and found.get("page"):
                link = source_citations.page_url(found.get("url"),
                                                 found["page"])
                if link:
                    bits.append(f"[page {found['page']} ↗]({link})")
            else:
                # No parsable summary table. Offer the weaker link, clearly
                # labelled as a different claim: where the decision section
                # starts, not where this verdict is printed.
                back = _fallbacks().get(str(appraisal_id).strip())
                if back and back.get("page"):
                    link = source_citations.page_url(back.get("url"),
                                                     back["page"])
                    where = ("decision section" if back["kind"] == "decision"
                             else "resolution")
                    if link:
                        bits.append(f"[{where}, page {back['page']} ↗]({link})")
                elif back and back.get("url"):
                    bits.append(f"[open the resolution ↗]({back['url']})")

            if found and name in found["table"] and not multi:
                symbol, summary = found["table"][name]
                from_table = subgroup_sources.verdict_from_table(symbol, summary)
                if from_table == label:
                    bits.append(":green[✓ matches the resolution]")
                elif from_table:
                    bits.append(
                        f":orange[⚠ resolution's table reads "
                        f"{endpoint_summary.PLAIN.get(from_table, from_table)}]")
            elif found and multi:
                bits.append(f":gray[table on page {found['page']} covers "
                            f"{len(rows)} subgroups — check it matches this one]")
            else:
                bits.append(":gray[AIS record — this page is where the "
                            "resolution's decision begins, not where this "
                            "verdict is printed]")

            st.markdown(f"**{name}** &nbsp; :{colour}[{plain}]")
            if wording:
                st.caption(f"“{str(wording)[:300]}” &nbsp;·&nbsp; "
                           + " &nbsp;·&nbsp; ".join(bits))
            else:
                st.caption(" &nbsp;·&nbsp; ".join(bits))

            # What was actually submitted for this category, if anything was
            # recovered. Uses the resolution's own URL for the page links.
            base_url = (found or {}).get("url") or (
                _fallbacks().get(str(appraisal_id).strip()) or {}).get("url")
            _render_submitted(
                appraisal_id, name, wording,
                lambda page: source_citations.page_url(base_url, page)
                if page else None)

    if found:
        st.caption(
            "Verdicts are G-BA AIS records, checked against the endpoint "
            "summary table printed in the resolution. Across 672 "
            "single-subgroup appraisals the two agree on 97.3% of 1,641 "
            "category verdicts; where they differ, the row is flagged above "
            "rather than resolved.")
