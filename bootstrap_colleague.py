"""
bootstrap_colleague.py
======================
Build a Publications Master CSV from PubMed + CrossRef + an optional CV file.

Sources searched (in order):
  1. PubMed  — via NCBI E-utilities (MEDLINE-indexed journals)
  2. CrossRef — via REST API (catches conference papers, book chapters,
                 and journals not indexed in MEDLINE)
  3. CV file  — optional .docx or .pdf for any entries missing from both

Usage:
    python bootstrap_colleague.py --config my_colleague.yaml

Output (written to current directory):
    <LastnameXX>_publications.csv     — import into Google Sheets
    <LastnameXX>_reconciliation.txt   — log of matches, conflicts, CV-only entries

Requirements:
    pip install pyyaml python-docx pdfminer.six

Python 3.8+
"""

import argparse
import csv
import json
import os
import re
import sys
import time
import unicodedata
import urllib.parse
import urllib.request
from pathlib import Path

# ── Optional imports (PDF and YAML) ─────────────────────────────────
try:
    import yaml
    HAS_YAML = True
except ImportError:
    HAS_YAML = False

try:
    from docx import Document as DocxDocument
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    from pdfminer.high_level import extract_text as pdf_extract_text
    HAS_PDF = True
except ImportError:
    HAS_PDF = False


# ════════════════════════════════════════════════════════════════════
# CONFIG
# ════════════════════════════════════════════════════════════════════

NCBI_BASE        = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"
CROSSREF_BASE    = "https://api.crossref.org/works"
REQUEST_DELAY    = 0.4    # seconds between NCBI requests (≤3/sec without API key)
CROSSREF_DELAY   = 0.2    # CrossRef polite pool allows ~50/sec; 0.2s is conservative
RETMAX_PAGE      = 500    # PMIDs per esearch page
FETCH_CHUNK      = 200    # PMIDs per esummary fetch
CROSSREF_ROWS    = 100    # results per CrossRef query page
CROSSREF_MAX     = 500    # max CrossRef results to fetch per author query
MATCH_THRESHOLD  = 0.65   # minimum title similarity to consider a match
DOI_MATCH        = True   # deduplicate by DOI before title matching

SCHEMA_COLS = [
    "id", "category", "status", "year", "authors", "title", "venue",
    "volume", "issue", "pages", "doi", "pmid",
    "submitted_date", "accepted_date", "published_date", "notes", "raw_text",
]

STOPWORDS = {
    "a", "an", "and", "the", "of", "in", "on", "for", "to", "with",
    "by", "from", "is", "are", "was", "were", "be", "been", "at",
    "as", "that", "this", "it", "or",
}


# ════════════════════════════════════════════════════════════════════
# NCBI E-UTILITIES
# ════════════════════════════════════════════════════════════════════

def _ncbi_get(endpoint, params):
    qs  = urllib.parse.urlencode(params)
    url = f"{NCBI_BASE}/{endpoint}?{qs}"
    req = urllib.request.Request(
        url, headers={"User-Agent": "enigma-cv-tools/1.0"}
    )
    time.sleep(REQUEST_DELAY)
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read())


def esearch_all(term):
    """Return all PMIDs for a query, paginating through results."""
    pmids    = []
    retstart = 0
    while True:
        data = _ncbi_get("esearch.fcgi", {
            "db": "pubmed", "retmode": "json",
            "retmax": RETMAX_PAGE, "retstart": retstart,
            "term": term,
        })
        result   = data.get("esearchresult", {})
        batch    = result.get("idlist", [])
        pmids.extend(batch)
        total    = int(result.get("count", 0))
        retstart += len(batch)
        if retstart >= total or not batch:
            break
    return pmids


def esummary_chunk(pmids):
    """Fetch esummary records for a list of PMIDs (<=FETCH_CHUNK at a time)."""
    records = []
    for i in range(0, len(pmids), FETCH_CHUNK):
        chunk  = pmids[i : i + FETCH_CHUNK]
        data   = _ncbi_get("esummary.fcgi", {
            "db": "pubmed", "retmode": "json",
            "id": ",".join(chunk),
        })
        result = data.get("result", {})
        for pmid in chunk:
            r = result.get(pmid)
            if not r or r.get("error"):
                continue
            doi = ""
            for aid in r.get("articleids", []):
                if aid.get("idtype") == "doi":
                    doi = aid.get("value", "")
            records.append({
                "pmid":    pmid,
                "title":   re.sub(r"\.$", "", r.get("title", "")),
                "journal": r.get("fulljournalname") or r.get("source", ""),
                "year":    (r.get("pubdate") or "")[:4],
                "volume":  r.get("volume", ""),
                "issue":   r.get("issue", ""),
                "pages":   r.get("pages", ""),
                "doi":     doi,
                "authors": ", ".join(
                    a.get("name", "") for a in r.get("authors", [])
                ),
                "source":  "pubmed",
            })
    return records


def fetch_pubmed(config):
    """Build a union query from ORCID + name/affiliation, fetch metadata."""
    terms = []

    orcid = config.get("orcid", "").strip()
    if orcid:
        terms.append(f"{orcid}[Author - Identifier]")

    name_query   = config.get("pubmed_query", "").strip()
    affiliations = config.get("affiliation_hints", [])
    if name_query and affiliations:
        aff_part = " OR ".join(f"{a}[Affiliation]" for a in affiliations)
        terms.append(f"(({name_query}) AND ({aff_part}))")
    elif name_query:
        terms.append(f"({name_query})")

    if not terms:
        print("ERROR: config must include pubmed_query and/or orcid.", file=sys.stderr)
        sys.exit(1)

    full_query = " OR ".join(terms)
    print(f"  PubMed query: {full_query}")

    pmids = list(dict.fromkeys(esearch_all(full_query)))
    print(f"  PMIDs found: {len(pmids)}")
    if not pmids:
        return []

    records = esummary_chunk(pmids)
    print(f"  Metadata fetched: {len(records)}")
    years = [int(r["year"]) for r in records if r["year"].isdigit()]
    if years:
        print(f"  Year range: {min(years)}\u2013{max(years)}")
    return records


# ════════════════════════════════════════════════════════════════════
# CROSSREF
# ════════════════════════════════════════════════════════════════════

def _crossref_get(params, contact_email):
    """Single paginated CrossRef request. contact_email enables polite pool."""
    qs  = urllib.parse.urlencode(params)
    url = f"{CROSSREF_BASE}?{qs}"
    headers = {
        "User-Agent": f"enigma-cv-tools/1.0 (mailto:{contact_email})"
    }
    req = urllib.request.Request(url, headers=headers)
    time.sleep(CROSSREF_DELAY)
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read())
    except Exception as e:
        print(f"  WARNING: CrossRef request failed: {e}", file=sys.stderr)
        return None


def _parse_crossref_item(item):
    """Convert a CrossRef work item into our internal record format."""
    title_list = item.get("title") or []
    title      = title_list[0] if title_list else ""
    title      = re.sub(r"\s+", " ", title).strip()

    container  = item.get("container-title") or []
    journal    = container[0] if container else ""

    # Year: prefer published-print, fall back to published-online or issued
    year = ""
    for date_field in ("published-print", "published-online", "issued"):
        dp = item.get(date_field, {}).get("date-parts", [[]])[0]
        if dp:
            year = str(dp[0])
            break

    # Authors: "Family, Given" -> "Family GI" style to match PubMed
    authors_raw = item.get("author", [])
    author_parts = []
    for a in authors_raw:
        family = a.get("family", "")
        given  = a.get("given", "")
        if family:
            initials = "".join(w[0] for w in given.split() if w) if given else ""
            author_parts.append(f"{family} {initials}".strip())
    authors = ", ".join(author_parts)

    return {
        "pmid":    "",
        "title":   title,
        "journal": journal,
        "year":    year,
        "volume":  item.get("volume", ""),
        "issue":   item.get("issue", ""),
        "pages":   item.get("page", ""),
        "doi":     item.get("DOI", ""),
        "authors": authors,
        "source":  "crossref",
        "type":    item.get("type", ""),   # e.g. "journal-article", "book-chapter"
    }


def fetch_crossref(config):
    """
    Search CrossRef by author name (and optionally affiliation).
    Returns a list of records in our internal format.

    CrossRef author search is fuzzier than PubMed — we fetch generously
    and rely on deduplication + title similarity to filter.
    """
    name = config.get("name", "").strip()
    if not name:
        print("  WARNING: no 'name' in config; skipping CrossRef search.", file=sys.stderr)
        return []

    # CrossRef author search uses "family" and "given" query params.
    # We split on the last word as family name; everything else as given.
    # e.g. "Elisabeth A. Wilde" -> family="Wilde", given="Elisabeth A."
    parts  = name.rsplit(" ", 1)
    family = parts[-1]
    given  = parts[0] if len(parts) > 1 else ""

    contact_email = config.get("contact_email", "enigma-cv-tools@example.com")
    affiliations  = config.get("affiliation_hints", [])

    records  = []
    fetched  = 0
    offset   = 0

    print(f"  CrossRef query: author={name}" +
          (f", affiliations={affiliations}" if affiliations else ""))

    while fetched < CROSSREF_MAX:
        rows = min(CROSSREF_ROWS, CROSSREF_MAX - fetched)
        params = {
            "query.author": f"{given} {family}".strip(),
            "rows":         rows,
            "offset":       offset,
            "select":       "DOI,title,author,container-title,published-print,"
                            "published-online,issued,volume,issue,page,type",
        }
        data = _crossref_get(params, contact_email)
        if not data:
            break

        items = (data.get("message") or {}).get("items") or []
        if not items:
            break

        for item in items:
            rec = _parse_crossref_item(item)
            if not rec["title"]:
                continue

            # Filter: the author must actually appear in the author list.
            # CrossRef's query.author is a relevance search, not a filter,
            # so results can include works by unrelated authors.
            author_str = rec["authors"].lower()
            if family.lower() not in author_str:
                continue

            # Affiliation filter: if hints provided, at least one must appear
            # in the raw item's affiliation strings.
            if affiliations:
                raw_affiliations = []
                for a in item.get("author", []):
                    for aff in a.get("affiliation", []):
                        raw_affiliations.append(aff.get("name", "").lower())
                aff_text = " ".join(raw_affiliations)
                if aff_text and not any(
                    hint.lower() in aff_text for hint in affiliations
                ):
                    continue

            records.append(rec)

        fetched += len(items)
        offset  += len(items)
        if len(items) < rows:
            break   # no more results

    print(f"  CrossRef records fetched: {fetched}, passed author filter: {len(records)}")
    return records


def deduplicate(pubmed_records, crossref_records, threshold=MATCH_THRESHOLD):
    """
    Remove CrossRef records that are already in pubmed_records.
    Deduplication strategy:
      1. Exact DOI match (fast, reliable)
      2. Title similarity >= threshold (catches same paper with different DOI
         or missing DOI on one side)
    Returns (crossref_only, doi_dupes, title_dupes) for reporting.
    """
    pubmed_dois   = {
        r["doi"].lower().strip()
        for r in pubmed_records
        if r.get("doi", "").strip()
    }

    crossref_only = []
    doi_dupes     = 0
    title_dupes   = 0

    for rec in crossref_records:
        doi = rec.get("doi", "").lower().strip()

        # 1. DOI match
        if doi and doi in pubmed_dois:
            doi_dupes += 1
            continue

        # 2. Title similarity match
        _, score = best_match(rec["title"], pubmed_records, threshold)
        if score >= threshold:
            title_dupes += 1
            continue

        crossref_only.append(rec)

    return crossref_only, doi_dupes, title_dupes


# ════════════════════════════════════════════════════════════════════
# CV PARSING
# ════════════════════════════════════════════════════════════════════

def _clean(text):
    """Normalize unicode and collapse whitespace."""
    text = unicodedata.normalize("NFKC", text)
    return re.sub(r"\s+", " ", text).strip()


def _extract_text_docx(path):
    if not HAS_DOCX:
        print("WARNING: python-docx not installed; skipping CV parsing.", file=sys.stderr)
        return ""
    doc = DocxDocument(path)
    return "\n".join(_clean(p.text) for p in doc.paragraphs if p.text.strip())


def _extract_text_pdf(path):
    if not HAS_PDF:
        print("WARNING: pdfminer.six not installed; skipping PDF CV parsing.", file=sys.stderr)
        return ""
    return _clean(pdf_extract_text(str(path)))


def extract_cv_titles(path):
    """
    Heuristic title extraction from a CV.
    Returns a list of candidate title strings.
    """
    p = Path(path)
    if not p.exists():
        print(f"WARNING: CV file not found: {path}", file=sys.stderr)
        return []

    suffix = p.suffix.lower()
    if suffix == ".docx":
        text = _extract_text_docx(p)
    elif suffix == ".pdf":
        text = _extract_text_pdf(p)
    else:
        print(f"WARNING: Unsupported CV format '{suffix}'. Use .docx or .pdf.", file=sys.stderr)
        return []

    titles = []
    for line in text.splitlines():
        line = line.strip()
        if 40 <= len(line) <= 250:
            if line[0].isupper() and not line.isupper():
                if not re.match(r"^\d{4}[;\s]", line):
                    if not re.match(r"^[A-Z][a-z]+\s[A-Z]", line):
                        titles.append(line)
    return titles


# ════════════════════════════════════════════════════════════════════
# TITLE MATCHING
# ════════════════════════════════════════════════════════════════════

def _tokenize(s):
    return [
        t for t in re.sub(r"[^a-z0-9\s]", " ", s.lower()).split()
        if t and t not in STOPWORDS
    ]


def title_similarity(a, b):
    ta, tb = _tokenize(a), _tokenize(b)
    if not ta or not tb:
        return 0.0
    sa, sb  = set(ta), set(tb)
    inter   = len(sa & sb)
    jaccard = inter / (len(sa) + len(sb) - inter)
    n       = min(8, len(ta), len(tb))
    ordered = sum(1 for i in range(n) if ta[i] == tb[i])
    return 0.75 * jaccard + 0.25 * (ordered / n if n else 0)


def best_match(query_title, candidates, threshold=MATCH_THRESHOLD):
    """Return (best_record, score) or (None, 0) if nothing meets threshold."""
    best, best_score = None, 0.0
    for c in candidates:
        s = title_similarity(query_title, c["title"])
        if s > best_score:
            best, best_score = c, s
    if best_score >= threshold:
        return best, best_score
    return None, 0.0


# ════════════════════════════════════════════════════════════════════
# BUILD OUTPUT ROWS
# ════════════════════════════════════════════════════════════════════

def _guess_category(record):
    """Rough category guess from CrossRef type, journal name, and title."""
    cr_type = record.get("type", "")
    title   = record.get("title", "").lower()
    journal = record.get("journal", "").lower()

    if cr_type == "book-chapter":
        return "book_chapter"
    if cr_type in ("proceedings-article", "conference-paper"):
        return "conference"
    if any(k in journal for k in ["review", "reviews"]):
        return "review"
    if any(k in title for k in ["editorial", "letter to"]):
        return "editorial"
    if any(k in journal for k in ["proceedings", "conference", "congress"]):
        return "conference"
    return "journal"


def record_to_row(idx, record):
    row = {col: "" for col in SCHEMA_COLS}
    row["id"]       = f"PUB{str(idx).zfill(4)}"
    row["category"] = _guess_category(record)
    row["status"]   = "published"
    row["year"]     = record.get("year", "")
    row["authors"]  = record.get("authors", "")
    row["title"]    = record.get("title", "")
    row["venue"]    = record.get("journal", "")
    row["volume"]   = record.get("volume", "")
    row["issue"]    = record.get("issue", "")
    row["pages"]    = record.get("pages", "")
    row["doi"]      = record.get("doi", "")
    row["pmid"]     = record.get("pmid", "")
    # Tag the source so users can see where each record came from
    src = record.get("source", "")
    if src == "crossref":
        row["notes"] = "Source: CrossRef"
    elif src == "cv_only":
        row["notes"] = "CV only — not matched in PubMed or CrossRef; verify manually"
    return row


def cv_only_to_row(idx, title):
    row = {col: "" for col in SCHEMA_COLS}
    row["id"]       = f"PUB{str(idx).zfill(4)}"
    row["category"] = "journal"
    row["status"]   = "published"
    row["title"]    = title
    row["notes"]    = "CV only — not matched in PubMed or CrossRef; verify manually"
    return row


# ════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════

def load_config(path):
    p = Path(path)
    if not p.exists():
        print(f"ERROR: Config file not found: {path}", file=sys.stderr)
        sys.exit(1)
    text = p.read_text(encoding="utf-8")
    if p.suffix.lower() in (".yaml", ".yml"):
        if not HAS_YAML:
            print("ERROR: pyyaml not installed. Run: pip install pyyaml", file=sys.stderr)
            sys.exit(1)
        return yaml.safe_load(text)
    elif p.suffix.lower() == ".json":
        return json.loads(text)
    else:
        print(f"ERROR: Config must be .yaml or .json, got: {p.suffix}", file=sys.stderr)
        sys.exit(1)


def slug(config):
    """Short filesystem-safe identifier from the pubmed_query or name."""
    q = config.get("pubmed_query", config.get("name", "output"))
    return re.sub(r"[^\w]", "_", q.split("[")[0].strip())


def main():
    parser = argparse.ArgumentParser(
        description="Bootstrap a Publications Master CSV from PubMed + CrossRef + CV."
    )
    parser.add_argument("--config",    required=True,  help="Path to YAML or JSON config file")
    parser.add_argument("--no-crossref", action="store_true",
                        help="Skip CrossRef search (PubMed + CV only)")
    parser.add_argument(
        "--threshold", type=float, default=MATCH_THRESHOLD,
        help=f"Title similarity threshold for deduplication and CV matching (default: {MATCH_THRESHOLD})"
    )
    args = parser.parse_args()

    config = load_config(args.config)
    name   = config.get("name", "Unknown")
    print(f"\nBootstrapping publications for: {name}")

    # ── 1. PubMed ──────────────────────────────────────────────────
    print("\nQuerying PubMed...")
    pubmed_records = fetch_pubmed(config)
    if not pubmed_records:
        print("WARNING: No PubMed records found. Continuing with CrossRef only.")

    # ── 2. CrossRef ────────────────────────────────────────────────
    crossref_only_records = []
    if not args.no_crossref:
        print("\nQuerying CrossRef...")
        crossref_records = fetch_crossref(config)
        if crossref_records:
            crossref_only_records, doi_dupes, title_dupes = deduplicate(
                pubmed_records, crossref_records, args.threshold
            )
            print(f"  Deduplicated: {doi_dupes} DOI matches, "
                  f"{title_dupes} title matches removed")
            print(f"  CrossRef-only (new): {len(crossref_only_records)}")
    else:
        print("\nSkipping CrossRef (--no-crossref flag set).")

    all_records = pubmed_records + crossref_only_records

    if not all_records:
        print("ERROR: No records found from any source. Check your config.")
        sys.exit(1)

    # ── 3. CV ──────────────────────────────────────────────────────
    cv_titles = []
    cv_path   = config.get("cv_path", "").strip() if config.get("cv_path") else ""
    if cv_path:
        print(f"\nParsing CV: {cv_path}")
        cv_titles = extract_cv_titles(cv_path)
        print(f"  CV title candidates extracted: {len(cv_titles)}")

    # ── 4. Match CV titles to PubMed + CrossRef ───────────────────
    recon_lines   = [f"Reconciliation report for {name}", "=" * 60, ""]
    matched_titles = set()
    cv_only        = []

    if cv_titles:
        recon_lines.append("CV title matching results:")
        recon_lines.append("-" * 40)
        for cv_title in cv_titles:
            rec, score = best_match(cv_title, all_records, args.threshold)
            if rec:
                matched_titles.add(rec["title"])
                src = rec.get("source", "pubmed")
                recon_lines.append(
                    f"  MATCHED ({score:.2f}) [{src}]\n"
                    f"    CV:     {cv_title[:100]}\n"
                    f"    Found:  {rec['title'][:100]}\n"
                )
            else:
                cv_only.append(cv_title)
                recon_lines.append(f"  CV ONLY: {cv_title[:100]}\n")

        recon_lines.append(
            f"\nSummary: {len(matched_titles)} matched, "
            f"{len(cv_only)} CV-only (not found in PubMed or CrossRef)\n"
        )

    # ── 5. Build rows ──────────────────────────────────────────────
    rows = []
    idx  = 1
    for rec in all_records:
        rows.append(record_to_row(idx, rec))
        idx += 1
    for title in cv_only:
        rows.append(cv_only_to_row(idx, title))
        idx += 1

    # ── 6. Write CSV ───────────────────────────────────────────────
    prefix   = slug(config)
    csv_path = f"{prefix}_publications.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=SCHEMA_COLS)
        writer.writeheader()
        writer.writerows(rows)
    print(f"\nWrote {len(rows)} rows to: {csv_path}")
    print(f"  PubMed: {len(pubmed_records)}  |  "
          f"CrossRef-only: {len(crossref_only_records)}  |  "
          f"CV-only: {len(cv_only)}")

    # ── 7. Write reconciliation report ────────────────────────────
    recon_lines.extend([
        "CrossRef-only entries (not in PubMed):",
        "-" * 40,
    ])
    for r in crossref_only_records:
        recon_lines.append(f"  [{r['year']}] {r['title'][:100]}")
        if r.get("doi"):
            recon_lines.append(f"         doi:{r['doi']}")

    recon_lines.extend([
        "",
        "PubMed entries (for reference):",
        "-" * 40,
    ])
    for r in pubmed_records:
        recon_lines.append(f"  [{r['year']}] {r['title'][:100]}")

    recon_path = f"{prefix}_reconciliation.txt"
    Path(recon_path).write_text("\n".join(recon_lines), encoding="utf-8")
    print(f"Wrote reconciliation report to: {recon_path}")
    print("\nDone. Import the CSV into your Publications Master Google Sheet.")
    print("Tip: review CrossRef-only entries in the reconciliation report — ")
    print("     they may include conference abstracts or preprints to exclude.\n")


if __name__ == "__main__":
    main()
