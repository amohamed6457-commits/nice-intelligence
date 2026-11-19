import argparse
import hashlib
import json
import os
import re
from collections import Counter

import pandas as pd

ENDPOINTS_JSON = "has_endpoints_full.json"
DOC_SHEET = "Appraisal_Documents"

DOC_RANK = [(r"avisdef", 0), (r"avis\d", 1), (r"avis", 2), (r"\.pdf$", 3)]
GOOD_TYPES = ("committee opinion", "committee decision", "ceesp")

CATEGORY_LABEL = {
    "mortality": "Mortality",
    "morbidity": "Morbidity",
    "quality_of_life": "Quality of life",
    "safety": "Safety",
    "surrogate": "Surrogate",
    "pharmacokinetic": "Pharmacokinetic",
    "other": "Other",
}

DIRECTION_LABEL = {
    "favours_drug": "Favours drug",
    "favours_comparator": "Favours comparator",
    "no_difference": "No difference",
    "not_stated": "Not stated",
}

CAVEAT = ("Endpoints extracted from the committee's own text and cited to the "
          "page. Verified at 89% precision on a hand-checked sample "
          "(33/37, 95% CI 75-96%); rows without a reported result are hidden "
          "by default because precision falls to 60% on those. This is what "
          "was found, not a guarantee of everything the avis reports.")


def norm_id(x):
    return re.sub(r"[^A-Za-z0-9]+", "", str(x or "")).upper()


def rank_url(url, dtype):
    u, t = str(url or "").lower(), str(dtype or "").lower()
    base = 0 if any(g in t for g in GOOD_TYPES) else 10
    for pat, r in DOC_RANK:
        if re.search(pat, u):
            return base + r
    return base + 9


def canonical_names(names):
    counts = Counter(n.strip().lower() for n in names if str(n).strip())
    mapping = {}
    for name in counts:
        if name.endswith("s") and name[:-1] in counts:
            keep = name if counts[name] >= counts[name[:-1]] else name[:-1]
            mapping[name] = keep
            mapping[name[:-1]] = keep
    return mapping


GENERIC_STUDY = re.compile(
    r"^(?:[ée]tude\s+)?(?:de\s+)?phase\s*[i1v23]+[ab]?$|^essai|^study$|"
    r"^pivotal|^rct$|^n/?a$", re.IGNORECASE)

NUMBER = re.compile(r"\d+(?:[.,]\d+)?")


def numbers_in(text):
    return tuple(sorted(n.replace(",", ".") for n in NUMBER.findall(str(text or ""))))


def named_study(value):
    v = str(value or "").strip()
    return bool(v) and not GENERIC_STUDY.match(v)


def dedupe(df):
    if df is None or df.empty:
        return df, 0
    work = df.copy()
    work["_nums"] = work["result"].map(numbers_in)
    work["_named"] = work["study"].map(named_study)
    work["_rank"] = (~work["_named"]).astype(int) * 1000 + work["page"].fillna(
        9999).astype(int)
    before = len(work)
    drop = []
    for _, grp in work.groupby(["aid_key", "endpoint_key"], sort=False):
        if len(grp) < 2:
            continue
        kept = []
        for idx, row in grp.sort_values("_rank").iterrows():
            nums = set(row["_nums"])
            dup = False
            for other in kept:
                if nums == other:
                    dup = True
                elif len(nums) >= 2 and len(other) >= 2 and (
                        nums <= other or other <= nums):
                    dup = True
                if dup:
                    break
            if dup:
                drop.append(idx)
            else:
                kept.append(nums)
    work = work.drop(index=drop).drop(columns=["_nums", "_rank"])
    return work.sort_index(), before - len(work)


def read_quotes(path="has_endpoint_quotes.json"):
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def flatten(done):
    rows = []
    for key, val in done.items():
        if not isinstance(val, list):
            continue
        aid, _, number = str(key).partition("#")
        for pos, e in enumerate(val):
            rows.append({
                "qkey": f"{key}|{pos}",
                "appraisal_id": aid,
                "aid_key": norm_id(aid),
                "section": number or "",
                "endpoint": e.get("endpoint_en") or e.get("endpoint") or "",
                "endpoint_fr": e.get("endpoint") or "",
                "category": e.get("category") or "other",
                "study": e.get("study") or "",
                "comparator": e.get("comparator") or "",
                "result": (e.get("result") or "").strip(),
                "direction": e.get("direction") or "not_stated",
                "primary": bool(e.get("primary")),
                "page": e.get("page"),
                "page_end": e.get("page_end"),
            })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["has_result"] = df["result"].astype(str).str.strip().ne("")
    cmap = canonical_names(df["endpoint"])
    df["endpoint_key"] = (df["endpoint"].str.strip().str.lower()
                          .map(lambda n: cmap.get(n, n)))
    df, dropped = dedupe(df)
    df.attrs["duplicates_removed"] = dropped
    return df


def read_endpoints(path=ENDPOINTS_JSON):
    if not os.path.exists(path):
        return pd.DataFrame()
    with open(path, encoding="utf-8") as fh:
        return flatten(json.load(fh))


def read_links(frame):
    if frame is None or frame.empty:
        return {}
    need = {"Institution", "Appraisal_ID", "Source_URL"}
    if not need.issubset(frame.columns):
        return {}
    d = frame[frame["Institution"].astype(str).str.upper()
              .str.contains("HAS", na=False)]
    d = d[d["Source_URL"].notna()]
    types = (d["Document_Type"] if "Document_Type" in d.columns
             else pd.Series([""] * len(d), index=d.index))
    best = {}
    for aid, url, dtype in zip(d["Appraisal_ID"], d["Source_URL"], types):
        key = norm_id(aid)
        r = rank_url(url, dtype)
        if key not in best or r < best[key][0]:
            best[key] = (r, str(url).strip())
    return {k: v[1] for k, v in best.items()}


def page_link(url, page):
    if not url:
        return ""
    try:
        return f"{url}#page={int(page)}"
    except (TypeError, ValueError):
        return url


def endpoints_for(df, appraisal_id, confident_only=True, categories=None):
    if df is None or df.empty:
        return df
    out = df[df["aid_key"] == norm_id(appraisal_id)]
    if confident_only:
        out = out[out["has_result"]]
    if categories:
        out = out[out["category"].isin(categories)]
    order = {"mortality": 0, "morbidity": 1, "quality_of_life": 2,
             "safety": 3, "surrogate": 4, "pharmacokinetic": 5, "other": 6}
    return (out.assign(_o=out["category"].map(order).fillna(9),
                       _p=(~out["primary"]).astype(int))
            .sort_values(["_p", "_o", "page"]).drop(columns=["_o", "_p"]))


def display_table(rows, url=""):
    if rows is None or rows.empty:
        return pd.DataFrame()
    out = pd.DataFrame({
        "Endpoint": rows["endpoint"],
        "Category": rows["category"].map(CATEGORY_LABEL).fillna("Other"),
        "Result": rows["result"],
        "Direction": rows["direction"].map(DIRECTION_LABEL).fillna("Not stated"),
        "Study": rows["study"],
        "Primary": rows["primary"].map({True: "Yes", False: ""}),
        "Page": rows["page"],
    })
    out["Source"] = [page_link(url, p) for p in rows["page"]]
    return out.reset_index(drop=True)


def tokens_from_row(row):
    found = []
    for col in ID_COLUMNS:
        try:
            value = row[col]
        except (KeyError, IndexError, TypeError):
            continue
        if pd.isna(value):
            continue
        found.extend(ct_tokens(value))
    return sorted(set(norm_id(t) for t in found))


def endpoint_counts(df, confident_only=True):
    if df is None or df.empty:
        return {}
    sub = df[df["has_result"]] if confident_only else df
    return sub.groupby("aid_key").size().to_dict()


def count_for_row(counts, row):
    return sum(counts.get(k, 0) for k in tokens_from_row(row))


def _selected_row(event, frame):
    try:
        picked = list(event.selection.rows)
    except Exception:
        return None
    if not picked or picked[0] >= len(frame):
        return None
    return frame.iloc[picked[0]]


TRANSLATE_MODEL = "claude-haiku-4-5-20251001"


def _wants_english(st):
    try:
        return st.session_state.get("w_src_lang_mode", "Original") == "English"
    except Exception:
        return False


def _cached_translation(text):
    try:
        import source_citations
        return source_citations.translation_for(text)
    except Exception:
        return None


def _api_key(st):
    try:
        return st.secrets.get("ANTHROPIC_API_KEY")
    except Exception:
        return None


def translate_quote(st, text):
    """English for one quote. Tries the pre-built translations file first,
    then one live call, cached in the session so a re-click is free."""
    body = str(text or "").strip()
    if not body:
        return None
    hit = _cached_translation(body)
    if hit:
        return hit
    key = "_ep_tr_" + hashlib.sha1(body.encode("utf-8")).hexdigest()[:16]
    try:
        if key in st.session_state:
            return st.session_state[key]
    except Exception:
        return None
    api_key = _api_key(st)
    if not api_key:
        return None
    try:
        import requests
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": api_key,
                     "anthropic-version": "2023-06-01",
                     "content-type": "application/json"},
            json={"model": TRANSLATE_MODEL, "max_tokens": 700,
                  "system": ("Translate the French health-technology "
                             "assessment passage into English. Keep every "
                             "number, unit and statistical term exactly as "
                             "written. Return only the translation."),
                  "messages": [{"role": "user", "content": body}]},
            timeout=20)
        resp.raise_for_status()
        out = "".join(b.get("text", "") for b in resp.json().get("content", [])
                      if b.get("type") == "text").strip()
    except Exception:
        return None
    if out:
        try:
            st.session_state[key] = out
        except Exception:
            pass
    return out or None


def render_endpoint_detail(st, row, url="", quotes=None):
    if row is None:
        return
    name = str(row.get("endpoint") or "").strip()
    french = str(row.get("endpoint_fr") or "").strip()
    st.markdown(f"#### {name}")
    if french and french.lower() != name.lower():
        st.caption(f"As written in the avis: {french}")

    cols = st.columns(4)
    cols[0].metric("Category", CATEGORY_LABEL.get(row.get("category"), "Other"))
    cols[1].metric("Direction",
                   DIRECTION_LABEL.get(row.get("direction"), "Not stated"))
    cols[2].metric("Primary endpoint", "Yes" if row.get("primary") else "No")
    page = row.get("page")
    cols[3].metric("Page", f"{int(page)}" if pd.notna(page) else "-")

    result = str(row.get("result") or "").strip()
    if result:
        st.markdown("**Result as reported**")
        st.info(result)
    else:
        st.warning("The committee text gave no quantitative result for this "
                   "endpoint, so it is shown without one. Rows like this were "
                   "correct 60% of the time in hand-checking, against 100% "
                   "for rows with a figure.")

    left, right = st.columns(2)
    study = str(row.get("study") or "").strip()
    comparator = str(row.get("comparator") or "").strip()
    left.markdown(f"**Study**  \n{study or '—'}")
    right.markdown(f"**Comparator**  \n{comparator or '—'}")

    found = (quotes or {}).get(row.get("qkey"))
    if found:
        original = found.get("quote", "")
        st.markdown("**What the committee actually wrote**")
        english = translate_quote(st, original) if _wants_english(st) else None
        if english:
            st.markdown(f"> {english}")
            with st.expander("Original French"):
                st.markdown(f"> {original}")
        else:
            st.markdown(f"> {original}")
            if _wants_english(st):
                st.caption("Could not translate this passage just now, so "
                           "the original is shown.")
        hit = found.get("figures_matched", 0)
        tot = found.get("figures_total", 0)
        st.caption(f"Located in the source text - {hit} of {tot} figures in "
                   "the result above appear verbatim in this passage. "
                   "The quote is the document's wording, not a summary.")
    elif result:
        st.caption("The figures in this result could not be located in the "
                   "section text. 2% of extracted endpoints fall here and "
                   "they are the ones most likely to be wrong - check the "
                   "page before relying on it.")

    section = str(row.get("section") or "").strip()
    target = page_link(url, page)
    if target:
        st.link_button(
            f"Open the avis at page {int(page)}" if pd.notna(page)
            else "Open the avis", target)
    if section:
        st.caption(f"Extracted from section {section} of the committee "
                   f"opinion" + (f", pages {int(row['page'])}"
                                 f"-{int(row['page_end'])}"
                                 if pd.notna(row.get("page_end")) else "") + ".")


def render_record_endpoints(st, df, links, row, heading=True, quotes=None):
    keys = tokens_from_row(row)
    if not keys:
        return
    rows = df[df["aid_key"].isin(keys) & df["has_result"]] if (
        df is not None and not df.empty) else None
    if rows is None or rows.empty:
        return
    order = {"mortality": 0, "morbidity": 1, "quality_of_life": 2,
             "safety": 3, "surrogate": 4, "pharmacokinetic": 5, "other": 6}
    rows = rows.assign(_o=rows["category"].map(order).fillna(9),
                       _p=(~rows["primary"]).astype(int)).sort_values(
        ["_p", "_o", "page"])
    if heading:
        st.markdown(f"**Clinical endpoints reported by the committee "
                    f"({len(rows)})**")
    url = next((links.get(k, "") for k in keys if links.get(k)), "")
    table = with_source(display_table(rows, url), rows)
    st.caption("Click a row to open that endpoint in full below.")
    event = st.dataframe(
        table, hide_index=True, width="stretch",
        column_order=RECORD_COLUMNS, column_config=_grid_config(st),
        on_select="rerun", selection_mode="single-row",
        key=f"w_record_endpoints_{keys[0]}")
    st.caption(CAVEAT)
    picked = _selected_row(event, table)
    if picked is not None:
        st.divider()
        render_endpoint_detail(st, _source_row(picked), url, quotes)


def render_appraisal_endpoints(st, df, links, appraisal_id,
                               confident_only=True, key=None):
    rows = endpoints_for(df, appraisal_id, confident_only)
    hidden = 0
    if confident_only:
        every = endpoints_for(df, appraisal_id, confident_only=False)
        hidden = len(every) - len(rows)
    if rows is None or rows.empty:
        st.caption("No clinical endpoints were extracted for this appraisal."
                   + (f" {hidden} without a reported result are hidden."
                      if hidden else ""))
        return
    url = links.get(norm_id(appraisal_id), "")
    table = display_table(rows, url)
    st.dataframe(
        table, hide_index=True, width="stretch",
        column_config={
            "Source": st.column_config.LinkColumn("Source", display_text="page"),
            "Result": st.column_config.TextColumn("Result", width="medium"),
            "Endpoint": st.column_config.TextColumn("Endpoint", width="large"),
        })
    note = CAVEAT
    if hidden:
        note += f" {hidden} endpoint(s) without a result are hidden here."
    st.caption(note)


CT_TOKEN = re.compile(r"CT[\s_-]*(\d{3,6})", re.IGNORECASE)
ID_COLUMNS = ("appraisal_id", "dossier_id", "Appraisal_ID", "appraisal", "id")
LABEL_COLUMNS = ("drug_name", "drug", "brand_name", "generic_name")


def tidy_id(value):
    t = re.sub(r"\s+", "", str(value or ""))
    m = re.match(r"^CT[_-]*(\d+)$", t, re.IGNORECASE)
    return f"CT-{m.group(1)}" if m else str(value or "").strip()


def ct_tokens(value):
    return ["CT" + n for n in CT_TOKEN.findall(str(value or ""))]


def id_columns(frame):
    if frame is None or not hasattr(frame, "columns"):
        return []
    return [c for c in ID_COLUMNS if c in frame.columns]


def ids_from_frame(frame):
    cols = id_columns(frame)
    if not cols:
        return None
    found = []
    for col in cols:
        for value in frame[col].dropna():
            found.extend(ct_tokens(value))
    return sorted(set(found)) or []


def labels_from_frame(frame):
    cols = id_columns(frame)
    if not cols:
        return {}
    label_col = next((c for c in LABEL_COLUMNS if c in frame.columns), None)
    if label_col is None:
        return {}
    out = {}
    for _, row in frame[cols + [label_col]].iterrows():
        name = row[label_col]
        if pd.isna(name):
            continue
        for col in cols:
            for token in ct_tokens(row[col]):
                out.setdefault(norm_id(token), str(name))
    return out


SHOW_COLUMNS = ["Appraisal", "Drug", "Endpoint", "Category", "Result",
                "Direction", "Study", "Primary", "Page", "Source"]
RECORD_COLUMNS = ["Endpoint", "Category", "Result", "Direction", "Study",
                  "Primary", "Page", "Source"]
COLUMN_CONFIG_KEYS = ("Source", "Endpoint", "Result")


def _grid_config(st):
    return {
        "Source": st.column_config.LinkColumn("Source", display_text="page"),
        "Endpoint": st.column_config.TextColumn("Endpoint", width="large"),
        "Result": st.column_config.TextColumn("Result", width="medium"),
    }


def with_source(table, rows):
    joined = table.copy()
    src = rows.reset_index(drop=True)
    for col in ("endpoint_fr", "category", "direction", "primary", "section",
                "page_end", "comparator", "study", "result", "endpoint",
                "page", "aid_key", "qkey"):
        if col in src.columns:
            joined["_" + col] = src[col].to_numpy()
    return joined


def _source_row(table_row):
    return {k[1:]: v for k, v in table_row.items() if k.startswith("_")}


def endpoint_table(df, links, appraisal_ids=None, labels=None,
                   confident_only=True):
    if df is None or df.empty:
        return pd.DataFrame()
    sub = df[df["has_result"]] if confident_only else df
    if appraisal_ids is not None:
        sub = sub[sub["aid_key"].isin({norm_id(a) for a in appraisal_ids})]
    if sub.empty:
        return pd.DataFrame()
    labels = labels or {}
    order = {"mortality": 0, "morbidity": 1, "quality_of_life": 2,
             "safety": 3, "surrogate": 4, "pharmacokinetic": 5, "other": 6}
    sub = sub.assign(_o=sub["category"].map(order).fillna(9),
                     _p=(~sub["primary"]).astype(int)).sort_values(
        ["appraisal_id", "_p", "_o", "page"])
    out = pd.DataFrame({
        "Appraisal": sub["appraisal_id"].map(tidy_id),
        "Drug": [labels.get(k, "") for k in sub["aid_key"]],
        "Endpoint": sub["endpoint"],
        "Category": sub["category"].map(CATEGORY_LABEL).fillna("Other"),
        "Result": sub["result"],
        "Direction": sub["direction"].map(DIRECTION_LABEL).fillna("Not stated"),
        "Study": sub["study"],
        "Primary": sub["primary"].map({True: "Yes", False: ""}),
        "Page": sub["page"],
        "Source": [page_link(links.get(k, ""), p)
                   for k, p in zip(sub["aid_key"], sub["page"])],
    })
    return with_source(out.reset_index(drop=True), sub)


def render_endpoint_table(st, df, links, frame=None, key="has_ep",
                          quotes=None):
    ids = ids_from_frame(frame)
    table = endpoint_table(df, links, ids, labels_from_frame(frame))
    if table.empty:
        if ids is not None and not ids:
            st.caption("The selected decisions carry no CT dossier number, "
                       "so there is nothing to match endpoints against.")
        else:
            st.caption("No extracted clinical endpoints for this selection. "
                       "Endpoints exist for 2,333 HAS appraisals - those with "
                       "a parseable clinical section in the committee "
                       "opinion.")
        return
    term = st.text_input("Filter endpoints", key=f"w_{key}_q",
                         placeholder="e.g. survival, nausea, ASCEND")
    view = table
    if term and term.strip():
        pat = re.escape(term.strip())
        hay = (view["Endpoint"].astype(str) + " " + view["Drug"].astype(str)
               + " " + view["Study"].astype(str) + " "
               + view["Appraisal"].astype(str))
        view = view[hay.str.contains(pat, case=False, na=False)]
    st.caption(f"{len(view):,} endpoints across "
               f"{view['Appraisal'].nunique():,} appraisals. "
               "Click a row to open it in full below.")
    view = view.reset_index(drop=True)
    event = st.dataframe(
        view, hide_index=True, width="stretch", height=460,
        column_order=SHOW_COLUMNS, column_config=_grid_config(st),
        on_select="rerun", selection_mode="single-row", key=f"w_{key}_grid")
    st.caption(CAVEAT)
    picked = _selected_row(event, view)
    if picked is None:
        return
    st.divider()
    data = _source_row(picked)
    render_endpoint_detail(st, data, links.get(data.get("aid_key"), ""),
                           quotes)


def overview(df, appraisal_ids=None, confident_only=True, top=12):
    if df is None or df.empty:
        return {}
    sub = df[df["has_result"]] if confident_only else df
    if appraisal_ids is not None:
        keys = {norm_id(a) for a in appraisal_ids}
        sub = sub[sub["aid_key"].isin(keys)]
    if sub.empty:
        return {}
    cats = (sub.groupby("category")
            .agg(endpoints=("endpoint", "size"),
                 appraisals=("aid_key", "nunique")).reset_index())
    cats["Category"] = cats["category"].map(CATEGORY_LABEL).fillna("Other")
    common = (sub.groupby("endpoint_key")
              .agg(Endpoints=("endpoint", "size"),
                   Appraisals=("aid_key", "nunique"))
              .sort_values("Endpoints", ascending=False).head(top)
              .reset_index().rename(columns={"endpoint_key": "Endpoint"}))
    return {
        "appraisals": sub["aid_key"].nunique(),
        "endpoints": len(sub),
        "categories": cats[["Category", "endpoints", "appraisals"]],
        "commonest": common,
        "directions": sub["direction"].map(DIRECTION_LABEL)
        .fillna("Not stated").value_counts().reset_index(),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--endpoints", default=ENDPOINTS_JSON)
    ap.add_argument("--workbook",
                    default="NICE_HAS_GBA_updated_2026-11-07.xlsx")
    ap.add_argument("--appraisal", default=None)
    ap.add_argument("--list", type=int, default=0)
    ap.add_argument("--has", default=None)
    args = ap.parse_args()

    df = read_endpoints(args.endpoints)
    if df.empty:
        raise SystemExit(f"no endpoints loaded from {args.endpoints}")
    print(f"{len(df):,} endpoints   {df['aid_key'].nunique():,} appraisals")
    print(f"  with a result : {df['has_result'].sum():,} "
          f"({df['has_result'].mean():.0%})")
    merged = df.loc[df['endpoint'].str.strip().str.lower()
                    != df['endpoint_key']]
    print(f"  name variants merged : {len(merged):,}")
    print(f"  duplicates removed   : "
          f"{df.attrs.get('duplicates_removed', 0):,}")
    print(f"  genuinely named study: {df['study'].map(named_study).sum():,} "
          f"({df['study'].map(named_study).mean():.0%})")

    links = {}
    if os.path.exists(args.workbook):
        links = read_links(pd.read_excel(args.workbook, DOC_SHEET))
    print(f"  source links  : {len(links):,} appraisals")

    ov = overview(df)
    print(f"\nconfident subset: {ov['endpoints']:,} endpoints across "
          f"{ov['appraisals']:,} appraisals")
    print("\nby category:")
    print(ov["categories"].to_string(index=False))
    print("\ncommonest:")
    print(ov["commonest"].to_string(index=False))

    if args.has:
        key = norm_id(args.has)
        hit = df[df["aid_key"] == key]
        print(f"\n{args.has}: {len(hit)} endpoints extracted, "
              f"{int(hit['has_result'].sum())} with a result")
        if hit.empty:
            print("  -> nothing was extracted for this appraisal, so an "
                  "empty panel is correct, not a wiring fault")
        return

    if args.list:
        top = (df[df["has_result"]].groupby("appraisal_id")
               .size().sort_values(ascending=False).head(args.list))
        print(f"\n{len(top)} appraisals with the most endpoints to demo with:")
        for aid, n in top.items():
            print(f"  {aid:<12} {n:>3} endpoints")
        return

    aid = args.appraisal or df["appraisal_id"].iloc[0]
    rows = endpoints_for(df, aid)
    url = links.get(norm_id(aid), "")
    print(f"\n=== {aid} ===  {url or '(no link)'}")
    print(display_table(rows, url).to_string(index=False))


if __name__ == "__main__":
    main()
