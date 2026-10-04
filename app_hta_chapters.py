"""
HTA Evidence Explorer - chapter-level citation UI.

Left pane  : ask a question; Claude searches indexed NICE committee sections
Right pane : the cited section in full, with the quoted sentence highlighted,
             its neighbours for context, and a link to the live NICE page

Citations are section-level (TA993 section 3.13), which is how NICE is cited in
practice, and every citation shows its source tier because a committee
discussion chapter and a summary guidance PDF are not equivalent evidence.

Run:
  streamlit run app_hta_chapters.py

Needs: an index built by build_chapter_index.py, and ANTHROPIC_API_KEY set
either in the environment or in .streamlit/secrets.toml
"""

import json
import os
import re

import streamlit as st

INDEX_DIR = "hta_index"
KEYWORD_STORE = os.path.join(INDEX_DIR, "keyword_index.json")
KEYWORD_STORE_GZ = KEYWORD_STORE + ".gz"
CITE_RE = re.compile(r"\[TA(\d+)\s*(?:§|section\s*)\s*(\d+\.\d+)\]", re.IGNORECASE)

TIER_LABEL = {
    "chapter": ("Committee discussion", "the committee's own published reasoning"),
    "fad": ("Final appraisal determination", "the determination document"),
    "committee_papers": ("Committee papers", "the full committee pack"),
    "guidance_pdf": ("Guidance summary", "summary only - not the committee's reasoning"),
    "excel_only": ("Metadata only", "no source document indexed"),
}

st.set_page_config(page_title="HTA Evidence Explorer", layout="wide")


# --------------------------------------------------------------------- index
@st.cache_resource
def load_index():
    """Returns (mode, handle). Prefers the vector index, falls back to keyword."""
    if os.path.exists(os.path.join(INDEX_DIR, "chroma.sqlite3")):
        try:
            import chromadb
            client = chromadb.PersistentClient(path=INDEX_DIR)
            return "vector", client.get_collection("nice_sections")
        except Exception:
            pass
    # Gzipped index first; a plain one from an older build still works.
    if os.path.exists(KEYWORD_STORE_GZ):
        import gzip
        with gzip.open(KEYWORD_STORE_GZ, "rt", encoding="utf-8") as f:
            return "keyword", json.load(f)
    if os.path.exists(KEYWORD_STORE):
        with open(KEYWORD_STORE, encoding="utf-8") as f:
            return "keyword", json.load(f)
    return None, None


def search(query, n=6):
    mode, handle = load_index()
    if mode is None:
        return []
    if mode == "vector":
        res = handle.query(query_texts=[query], n_results=n,
                           include=["documents", "metadatas"])
        return [{"text": res["documents"][0][i].split("\n")[-1],
                 "metadata": res["metadatas"][0][i]}
                for i in range(len(res["documents"][0]))]

    terms = [t for t in re.findall(r"[a-z0-9]+", query.lower()) if len(t) > 2]
    if not terms:
        return []
    scored = []
    for rec in handle:
        md = rec.get("metadata", {})
        hay = " ".join(filter(None, (md.get("theme"), md.get("headline"),
                                     rec.get("text", "")))).lower()
        matched = sum(1 for t in terms if t in hay)
        if not matched:
            continue
        scored.append((matched / len(terms), sum(hay.count(t) for t in terms), rec))
    scored.sort(key=lambda x: (-x[0], -x[1]))
    return [{"text": r["text"], "metadata": r["metadata"]} for _, _, r in scored[:n]]


def get_section(ta, section):
    """Fetch one specific section for the viewer pane."""
    mode, handle = load_index()
    if mode == "keyword":
        for rec in handle:
            md = rec["metadata"]
            if md["ta"] == str(ta) and md["section"] == section:
                return {"text": rec["text"], "metadata": md}
    elif mode == "vector":
        try:
            res = handle.get(ids=[f"TA{ta}_{section}"], include=["documents", "metadatas"])
            if res["ids"]:
                return {"text": res["documents"][0].split("\n")[-1],
                        "metadata": res["metadatas"][0]}
        except Exception:
            pass
    return None


def neighbours(ta, section, span=1):
    """Adjacent sections, so a quote can be read in context."""
    chap, num = section.split(".")
    out = []
    for d in range(-span, span + 1):
        n = int(num) + d
        if n < 1 or d == 0:
            continue
        s = get_section(ta, f"{chap}.{n}")
        if s:
            out.append(s)
    return out


# ------------------------------------------------------------------ the model
SYSTEM = """You answer questions about NICE technology appraisals using only the \
indexed committee sections provided to you by the search_nice_sections tool.

Rules:
- Search before answering. Do not answer from memory about any specific appraisal.
- Cite every substantive claim as [TA<number> §<section>], e.g. [TA993 §3.13].
- Quote the committee's own words when explaining why something was accepted or \
rejected. Keep quotes short and exact.
- If the search returns nothing relevant, say so plainly. Never invent a section \
number or an appraisal.
- Distinguish what the company argued, what the EAG/ERG argued, and what the \
committee concluded. These are routinely confused and the distinction matters.
- If asked about an appraisal that is not in the results, say it is not indexed \
rather than guessing."""

TOOLS = [{
    "name": "search_nice_sections",
    "description": ("Search indexed NICE committee discussion sections. Returns "
                    "whole sections with their appraisal number, section number, "
                    "theme and the committee's headline conclusion."),
    "input_schema": {
        "type": "object",
        "properties": {
            "query": {"type": "string",
                      "description": "What to look for, e.g. 'survival extrapolation "
                                     "rejected implausible' or 'end of life criteria'"},
        },
        "required": ["query"],
    },
}]


def key_sources():
    """Every place a key might be, in precedence order, for display and use.

    st.secrets raises if no secrets file exists anywhere, rather than returning
    a default, so that lookup has to be guarded.
    """
    found = []

    env = os.environ.get("ANTHROPIC_API_KEY", "")
    found.append(("environment", env))

    sec = ""
    try:
        sec = st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        sec = ""
    found.append(("secrets.toml", sec))

    # Both the saved copy and the live widget value, so a key typed this run is
    # picked up even if the save step has not happened yet.
    found.append(("sidebar (saved)", st.session_state.get("api_key_saved", "")))
    found.append(("sidebar (field)", st.session_state.get("api_key_input", "")))

    return found


def get_api_key():
    for _, value in key_sources():
        if value and str(value).strip():
            return str(value).strip()
    return ""


def ask_claude(question, history, model):
    """Agentic loop. Returns (answer_text, retrieved_sections, searches_made)."""
    import anthropic
    key = get_api_key()
    if not key:
        return ("No API key. Paste one into the sidebar, or set ANTHROPIC_API_KEY "
                "in your environment."), [], []

    client = anthropic.Anthropic(api_key=key)
    messages = history + [{"role": "user", "content": question}]
    retrieved, searches = [], []

    for _ in range(6):                      # bounded, so a loop cannot run away
        resp = client.messages.create(
            model=model, max_tokens=2000,
            system=SYSTEM, tools=TOOLS, messages=messages,
        )

        if resp.stop_reason != "tool_use":
            text = "".join(b.text for b in resp.content if hasattr(b, "text"))
            return text, retrieved, searches

        results = []
        for block in resp.content:
            if block.type != "tool_use":
                continue
            q = block.input.get("query", "")
            searches.append(q)
            hits = search(q, n=6)
            retrieved.extend(hits)
            if hits:
                payload = "\n\n".join(
                    f"[TA{h['metadata']['ta']} §{h['metadata']['section']}] "
                    f"({TIER_LABEL.get(h['metadata']['source_tier'], ('?',))[0]})\n"
                    f"theme: {h['metadata']['theme']}\n"
                    f"committee headline: {h['metadata']['headline']}\n"
                    f"{h['text']}"
                    for h in hits)
            else:
                payload = "No matching sections in the index."
            results.append({"type": "tool_result", "tool_use_id": block.id,
                            "content": payload})

        messages.append({"role": "assistant", "content": resp.content})
        messages.append({"role": "user", "content": results})

    return "Stopped after six search rounds without settling on an answer.", retrieved, searches


# ------------------------------------------------------------------------ UI
def highlight(text, quote):
    """Mark quoted spans inside the section text.

    Two things make this harder than a substring search:

    1. Quotes are routinely elided - "too much uncertainty ... preferred ICER" -
       so the quote is shorter than the span it refers to. Extending from the
       start by len(quote) therefore ends mid-word. Each fragment between
       ellipses is located independently instead.
    2. Streamlit renders '==x==' literally; its highlight directive is
       :colour-background[...]. Square brackets inside the span would break
       that directive, so such fragments are left unmarked.
    """
    if not quote:
        return text

    fragments = [f.strip(" .,;:…") for f in re.split(r"…|\.\.\.", quote)]
    fragments = [f for f in fragments if len(f) >= 25 and "[" not in f and "]" not in f]
    if not fragments:
        return text

    spans = []
    low = text.lower()
    for frag in fragments:
        i = low.find(frag.lower())
        if i < 0:                       # tolerate minor whitespace differences
            loose = re.escape(frag.lower())
            loose = re.sub(r"\\\s+", r"\\s+", loose)
            m = re.search(loose, low)
            if not m:
                continue
            i, j = m.span()
        else:
            j = i + len(frag)
        spans.append((i, j))

    if not spans:
        return text

    # Merge overlaps, then apply right-to-left so earlier offsets stay valid.
    spans.sort()
    merged = [spans[0]]
    for a, b in spans[1:]:
        if a <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], b))
        else:
            merged.append((a, b))

    out = text
    for a, b in reversed(merged):
        out = out[:a] + ":orange-background[" + out[a:b] + "]" + out[b:]
    return out


mode, handle = load_index()

st.title("HTA Evidence Explorer")

if mode is None:
    st.error("No index found. Build one first:")
    st.code("python parse_nice_chapters.py\npython build_chapter_index.py --keyword")
    st.stop()

n_indexed = handle.count() if mode == "vector" else len(handle)
tas = (len({r["metadata"]["ta"] for r in handle}) if mode == "keyword" else None)

with st.sidebar:
    st.subheader("Index")
    st.metric("Sections", f"{n_indexed:,}")
    if tas:
        st.metric("Appraisals", tas)
    st.caption(f"Retrieval: {mode}")
    st.divider()
    model = st.selectbox("Model", ["claude-opus-5-5", "claude-sonnet-5-5",
                                   "claude-haiku-4-5-20251001"], index=1)

    # The field is always rendered. Hiding it once a key is entered would let
    # Streamlit discard the widget's state on a later rerun, losing the key.
    typed = st.text_input(
        "Anthropic API key", type="password", key="api_key_input",
        help="Held in this browser session only. Set ANTHROPIC_API_KEY in "
             "your environment to avoid retyping it.")
    if typed and typed.strip():
        st.session_state["api_key_saved"] = typed.strip()

    active = get_api_key()
    if active:
        st.success(f"Key active ({len(active)} chars, ends {active[-4:]})")
    else:
        st.error("No key found")

    with st.expander("Key detection", expanded=not active):
        for name, value in key_sources():
            if value and str(value).strip():
                st.caption(f"{name}: found, {len(str(value).strip())} chars")
            else:
                st.caption(f"{name}: empty")
    st.divider()
    st.caption("Citations are section-level. Source tiers are shown because a "
               "committee discussion and a guidance summary are not equivalent "
               "evidence.")

left, right = st.columns([1, 1])

# ---- left: conversation
with left:
    st.subheader("Ask")

    if "msgs" not in st.session_state:
        st.session_state.msgs = []
    if "api_history" not in st.session_state:
        st.session_state.api_history = []

    for mi, m in enumerate(st.session_state.msgs):
        with st.chat_message(m["role"]):
            st.markdown(m["content"])
            for ci, (ta, sec) in enumerate(m.get("citations", [])):
                if st.button(f"TA{ta} · section {sec}", key=f"c{mi}_{ci}"):
                    st.session_state.selected = (ta, sec)
                    st.session_state.selected_quote = m.get("content", "")
                    st.rerun()
            if m.get("searches"):
                st.caption("searched: " + " · ".join(f"“{s}”" for s in m["searches"]))

    q = st.chat_input("e.g. why was an ICER considered too uncertain?")
    if q:
        st.session_state.msgs.append({"role": "user", "content": q})
        with st.spinner("searching committee sections..."):
            answer, hits, searches = ask_claude(
                q, st.session_state.api_history, model)

        cites = []
        for ta, sec in CITE_RE.findall(answer):
            if (ta, sec) not in cites:
                cites.append((ta, sec))

        st.session_state.msgs.append({
            "role": "assistant", "content": answer,
            "citations": cites, "searches": searches,
        })
        st.session_state.api_history = (
            st.session_state.api_history
            + [{"role": "user", "content": q},
               {"role": "assistant", "content": answer}]
        )[-8:]
        if cites:
            st.session_state.selected = cites[0]
            st.session_state.selected_quote = answer
        st.rerun()

# ---- right: the source
with right:
    st.subheader("Source")

    sel = st.session_state.get("selected")
    if not sel:
        st.info("Ask a question. Cited sections open here in full.")
    else:
        ta, sec = sel
        s = get_section(ta, sec)
        if not s:
            st.warning(f"TA{ta} section {sec} is not in the index.")
        else:
            md = s["metadata"]
            label, caveat = TIER_LABEL.get(md["source_tier"], ("Unknown", ""))

            st.markdown(f"### TA{ta} · section {sec}")
            if md["source_tier"] == "guidance_pdf":
                st.warning(f"**{label}** - {caveat}")
            else:
                st.caption(f"{label} - {caveat}")

            if md.get("theme"):
                st.markdown(f"**{md['theme']}**")
            if md.get("headline"):
                st.markdown(f"*{md['headline']}*")
            st.divider()

            # Every quoted span in the answer, not just the first - a single
            # section is often quoted several times. highlight() splits on
            # ellipses, so joining with one lets it locate each independently.
            answer = st.session_state.get("selected_quote", "")
            quotes = re.findall(r"[“\"]([^”\"]{20,400})[”\"]", answer)
            st.markdown(highlight(s["text"], "…".join(quotes)))

            if md.get("url"):
                st.markdown(f"[Open on nice.org.uk]({md['url']})")

            ctx = neighbours(ta, sec)
            if ctx:
                with st.expander(f"Surrounding sections ({len(ctx)})"):
                    for c in ctx:
                        cmd = c["metadata"]
                        st.markdown(f"**{cmd['section']}** - {cmd['headline']}")
                        st.caption(c["text"][:400] + ("..." if len(c["text"]) > 400 else ""))
                        st.divider()
