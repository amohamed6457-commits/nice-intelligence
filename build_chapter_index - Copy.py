"""
Stage 3: build a searchable index from parsed chapter sections.

Every chunk is one whole NICE section (3.1, 3.14 ...), never a fragment, so a
citation is always a real quotable unit the reader can verify on nice.org.uk.

Each chunk carries a source_tier. These are NOT equivalent evidence and the UI
must distinguish them:
    chapter          committee's published reasoning  (strongest routinely available)
    fad              final appraisal determination
    committee_papers full committee pack
    guidance_pdf     summary guidance only - weaker, it is not the reasoning
    excel_only       metadata row, no document text at all

Embedding: ChromaDB's default model is downloaded once (~80MB) then works
offline. Pass --keyword to skip embeddings entirely and use substring matching,
which needs no download and is often fine given how descriptive NICE's section
headlines are.

Usage:
  python build_chapter_index.py                        # from chapter_sections.json
  python build_chapter_index.py --keyword              # no embedding model
  python build_chapter_index.py --query "survival extrapolation rejected"
"""

import argparse
import json
import os
import re
import sys

DB_DIR = "hta_index"
COLLECTION = "nice_sections"
KEYWORD_STORE = os.path.join(DB_DIR, "keyword_index.json")
KEYWORD_STORE_GZ = KEYWORD_STORE + ".gz"


def log(m):
    print(m, flush=True)


def load_sections(path):
    if not os.path.exists(path):
        log(f"Not found: {path}  (run parse_nice_chapters.py first)")
        return []
    with open(path, encoding="utf-8") as f:
        secs = json.load(f)
    # Drop anything the parser could not attribute to an appraisal.
    keep = [s for s in secs if s.get("ta") and s.get("section") and s.get("text")]
    if len(keep) != len(secs):
        log(f"Skipped {len(secs) - len(keep)} section(s) missing ta/section/text")
    return keep


def chunk_id(s):
    return f"TA{s['ta']}_{s['section']}"


def to_document(s):
    """Text given to the retriever. Headline and theme are included because
    NICE writes the committee's conclusion into the headline, which is often
    the most searchable sentence in the whole section."""
    parts = []
    if s.get("theme"):
        parts.append(s["theme"])
    if s.get("headline"):
        parts.append(s["headline"])
    parts.append(s["text"])
    return "\n".join(parts)


def to_metadata(s, tier="chapter"):
    return {
        "ta": str(s["ta"]),
        "appraisal_id": f"TA{s['ta']}",
        "section": s["section"],
        "headline": s.get("headline") or "",
        "theme": s.get("theme") or "",
        "anchor": s.get("anchor") or "",
        "source_tier": tier,
        "chars": len(s["text"]),
        "url": f"https://www.nice.org.uk/guidance/ta{s['ta']}/chapter/"
               f"{s['section'].split('.')[0]}-Committee-discussion"
               f"#{s.get('anchor', '')}",
    }


# ----------------------------------------------------------------- keyword mode
def build_keyword(sections):
    """Write a gzipped, slim index.

    Slim because the old 'document' field duplicated the section text, and
    headline/theme already live in metadata - it can be rebuilt on load.
    Gzipped because this file has to be committed to a repo for a deployed
    app to see it: ~9,700 sections go from roughly 31 MB to 4 MB, which is
    the difference between a repo GitHub complains about and one it doesn't.
    """
    import gzip
    os.makedirs(DB_DIR, exist_ok=True)
    store = [{
        "id": chunk_id(s),
        "text": s["text"],
        "metadata": to_metadata(s),
    } for s in sections]

    payload = json.dumps(store, ensure_ascii=False).encode("utf-8")
    with gzip.open(KEYWORD_STORE_GZ, "wb", compresslevel=9) as f:
        f.write(payload)

    mb = os.path.getsize(KEYWORD_STORE_GZ) / 1048576
    log(f"Keyword index written: {KEYWORD_STORE_GZ}  "
        f"({len(store)} sections, {mb:.1f} MB compressed from "
        f"{len(payload)/1048576:.1f} MB)")

    # An uncompressed index from an earlier run would shadow this one.
    if os.path.exists(KEYWORD_STORE):
        os.remove(KEYWORD_STORE)
        log(f"Removed the older uncompressed {KEYWORD_STORE}")
    return store


def load_keyword_store():
    """Read the gzipped index, or a plain one left by an older build."""
    import gzip
    if os.path.exists(KEYWORD_STORE_GZ):
        with gzip.open(KEYWORD_STORE_GZ, "rt", encoding="utf-8") as f:
            return json.load(f)
    with open(KEYWORD_STORE, encoding="utf-8") as f:
        return json.load(f)


def searchable_text(rec):
    """Rebuild what used to be the stored 'document' field."""
    md = rec.get("metadata", {})
    return "\n".join(p for p in (md.get("theme"), md.get("headline"),
                                 rec.get("text", "")) if p)


def search_keyword(query, n=5, store=None):
    if store is None:
        store = load_keyword_store()
    terms = [t for t in re.findall(r"[a-z0-9]+", query.lower()) if len(t) > 2]
    if not terms:
        return []
    scored = []
    for rec in store:
        hay = searchable_text(rec).lower()
        hits = sum(hay.count(t) for t in terms)
        matched = sum(1 for t in terms if t in hay)
        if matched == 0:
            continue
        # Require most terms present, then rank by density.
        coverage = matched / len(terms)
        scored.append((coverage, hits, rec))
    scored.sort(key=lambda x: (-x[0], -x[1]))
    return [r for _, _, r in scored[:n]]


# ------------------------------------------------------------------ vector mode
def build_vector(sections):
    import chromadb
    os.makedirs(DB_DIR, exist_ok=True)
    client = chromadb.PersistentClient(path=DB_DIR)
    try:
        client.delete_collection(COLLECTION)
    except Exception:
        pass
    coll = client.create_collection(COLLECTION, metadata={"hnsw:space": "cosine"})

    B = 200
    for i in range(0, len(sections), B):
        batch = sections[i:i + B]
        coll.add(
            ids=[chunk_id(s) for s in batch],
            documents=[to_document(s) for s in batch],
            metadatas=[to_metadata(s) for s in batch],
        )
        log(f"  indexed {min(i + B, len(sections))}/{len(sections)}")
    log(f"Vector index written: {DB_DIR}/  ({coll.count()} sections)")
    return coll


def search_vector(query, n=5):
    import chromadb
    client = chromadb.PersistentClient(path=DB_DIR)
    coll = client.get_collection(COLLECTION)
    res = coll.query(query_texts=[query], n_results=n,
                     include=["documents", "metadatas", "distances"])
    out = []
    for i, doc in enumerate(res["documents"][0]):
        out.append({
            "id": res["ids"][0][i],
            "document": doc,
            "text": doc,
            "metadata": res["metadatas"][0][i],
            "distance": res["distances"][0][i],
        })
    return out


# ------------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sections", default="chapter_sections.json")
    ap.add_argument("--keyword", action="store_true",
                    help="substring index, no embedding model download")
    ap.add_argument("--query", help="search an existing index instead of building")
    ap.add_argument("-n", type=int, default=5)
    args = ap.parse_args()

    if args.query:
        have_keyword = os.path.exists(KEYWORD_STORE_GZ) or os.path.exists(KEYWORD_STORE)
        if args.keyword or (not os.path.exists(os.path.join(DB_DIR, "chroma.sqlite3"))
                            and have_keyword):
            hits = search_keyword(args.query, args.n)
        else:
            hits = search_vector(args.query, args.n)
        if not hits:
            log("No matches.")
            return
        for h in hits:
            md = h["metadata"]
            log(f"\n{md['appraisal_id']} section {md['section']}   [{md['source_tier']}]")
            if md["theme"]:
                log(f"  theme    : {md['theme']}")
            if md["headline"]:
                log(f"  headline : {md['headline']}")
            body = h["text"].split("\n")[-1]
            log(f"  {body[:260]}...")
        return

    sections = load_sections(args.sections)
    if not sections:
        return

    tas = sorted({s["ta"] for s in sections}, key=lambda x: int(x))
    lens = sorted(len(s["text"]) for s in sections)
    log(f"Sections  : {len(sections)}")
    log(f"Appraisals: {len(tas)}  (TA{tas[0]} .. TA{tas[-1]})")
    log(f"Length    : min {lens[0]}  median {lens[len(lens)//2]}  max {lens[-1]} chars")
    log(f"Mode      : {'keyword' if args.keyword else 'vector embeddings'}\n")

    if args.keyword:
        build_keyword(sections)
    else:
        try:
            build_vector(sections)
        except ImportError:
            log("chromadb not installed. Either:")
            log("  pip install chromadb")
            log("  or re-run with --keyword to skip embeddings")
            sys.exit(1)

    log("\nTry a search:")
    log('  python build_chapter_index.py --query "cost effectiveness estimate uncertain"'
        + (" --keyword" if args.keyword else ""))


if __name__ == "__main__":
    main()
