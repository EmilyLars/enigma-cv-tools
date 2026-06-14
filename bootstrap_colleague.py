"""
bootstrap_colleague.py
======================
Build a Publications Master CSV from PubMed + an optional CV file.

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

NCBI_BASE      = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"
REQUEST_DELAY  = 0.4          # seconds between NCBI requests (≤3/sec without API key)
RETMAX_PAGE    = 500          # PMIDs per esearch page
FETCH_CHUNK    = 200          # PMIDs per esummary fetch
MATCH_THRESHOLD = 0.65        # minimum title similarity to consider a match

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
    pmids = []
    retstart = 0
    while True:
        data = _ncbi_get("esearch.fcgi", {
            "db": "pubmed", "retmode": "json",
            "retmax": RETMAX_PAGE, "retstart": retstart,
            "term": term,
        })
        result  = data.get("esearchresult", {})
        batch   = result.get("idlist", [])
        pmids.extend(batch)
        total   = int(result.get("count", 0))
        retstart += len(batch)
        if retstart >= total or not batch:
            break
    return pmids


def esummary_chunk(pmids):
    """Fetch esummary records for a list of PMIDs (≤FETCH_CHUNK at a time)."""
    records = []
    for i in range(0, len(pmids), FETCH_CHUNK):
        chunk = pmids[i : i + FETCH_CHUNK]
        data  = _ncbi_get("esummary.fcgi", {
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
            })
    return records


def fetch_pubmed(config):
    """
    Build a union query from ORCID (if provided) and name+affiliation,
    de-duplicate, fetch metadata.
    """
    terms = []

    orcid = config.get("orcid", "").strip()
    if orcid:
        terms.append(f"{orcid}[Author - Identifier]")

    name_query = config.get("pubmed_query", "").strip()
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

    pmids = list(dict.fromkeys(esearch_all(full_query)))   # preserve order, de-dup
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
        # Heuristic: lines that look like publication titles are typically
        # 40–250 chars, start with a capital letter, and are not all-caps
        # section headers.
        if 40 <= len(line) <= 250:
            if line[0].isupper() and not line.isupper():
                # Exclude lines that are likely venue / author lines
                if not re.match(r"^\d{4}[;\s]", line):       # starts with year
                    if not re.match(r"^[A-Z][a-z]+\s[A-Z]", line):  # "Author Initials"
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
    sa, sb = set(ta), set(tb)
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
    """Very rough category guess from journal/title keywords."""
    title   = record.get("title", "").lower()
    journal = record.get("journal", "").lower()
    if any(k in journal for k in ["review", "reviews"]):
        return "review"
    if any(k in title for k in ["editorial", "letter to"]):
        return "editorial"
    if any(k in journal for k in ["proceedings", "conference", "congress"]):
        return "conference"
    return "journal"


def pubmed_to_row(pmid_idx, record):
    row = {col: "" for col in SCHEMA_COLS}
    row["id"]       = f"PUB{str(pmid_idx).zfill(4)}"
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
    return row


def cv_only_to_row(pmid_idx, title):
    row = {col: "" for col in SCHEMA_COLS}
    row["id"]       = f"PUB{str(pmid_idx).zfill(4)}"
    row["category"] = "journal"
    row["status"]   = "published"
    row["title"]    = title
    row["notes"]    = "CV only — not matched in PubMed; verify manually"
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
    """Short filesystem-safe identifier from the pubmed_query."""
    q = config.get("pubmed_query", config.get("name", "output"))
    return re.sub(r"[^\w]", "_", q.split("[")[0].strip())


def main():
    parser = argparse.ArgumentParser(
        description="Bootstrap a Publications Master CSV from PubMed + CV."
    )
    parser.add_argument("--config", required=True, help="Path to YAML or JSON config file")
    parser.add_argument(
        "--threshold", type=float, default=MATCH_THRESHOLD,
        help=f"Title similarity threshold for CV matching (default: {MATCH_THRESHOLD})"
    )
    args = parser.parse_args()

    config = load_config(args.config)
    name   = config.get("name", "Unknown")
    print(f"\nBootstrapping publications for: {name}")

    # ── 1. PubMed ──────────────────────────────────────────────────
    print("\nQuerying PubMed...")
    pubmed_records = fetch_pubmed(config)
    if not pubmed_records:
        print("No PubMed records found. Check your config and try again.")
        sys.exit(1)

    # ── 2. CV ──────────────────────────────────────────────────────
    cv_titles = []
    cv_path   = config.get("cv_path", "").strip()
    if cv_path:
        print(f"\nParsing CV: {cv_path}")
        cv_titles = extract_cv_titles(cv_path)
        print(f"  CV title candidates extracted: {len(cv_titles)}")

    # ── 3. Match CV titles to PubMed ──────────────────────────────
    recon_lines  = [f"Reconciliation report for {name}", "=" * 60, ""]
    matched_pmids = set()
    cv_only       = []

    if cv_titles:
        recon_lines.append("CV title matching results:")
        recon_lines.append("-" * 40)
        for cv_title in cv_titles:
            rec, score = best_match(cv_title, pubmed_records, args.threshold)
            if rec:
                matched_pmids.add(rec["pmid"])
                recon_lines.append(
                    f"  MATCHED ({score:.2f})\n"
                    f"    CV:     {cv_title[:100]}\n"
                    f"    PubMed: {rec['title'][:100]}\n"
                )
            else:
                cv_only.append(cv_title)
                recon_lines.append(f"  CV ONLY: {cv_title[:100]}\n")

        recon_lines.append(
            f"\nSummary: {len(matched_pmids)} matched, "
            f"{len(cv_only)} CV-only (not found in PubMed)\n"
        )

    # ── 4. Build rows ──────────────────────────────────────────────
    rows   = []
    idx    = 1
    for rec in pubmed_records:
        rows.append(pubmed_to_row(idx, rec))
        idx += 1
    for title in cv_only:
        rows.append(cv_only_to_row(idx, title))
        idx += 1

    # ── 5. Write CSV ───────────────────────────────────────────────
    prefix   = slug(config)
    csv_path = f"{prefix}_publications.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=SCHEMA_COLS)
        writer.writeheader()
        writer.writerows(rows)
    print(f"\nWrote {len(rows)} rows to: {csv_path}")

    # ── 6. Write reconciliation report ────────────────────────────
    recon_lines.extend([
        "PubMed-only entries (not in CV or CV not provided):",
        "-" * 40,
    ])
    pubmed_only = [r for r in pubmed_records if r["pmid"] not in matched_pmids]
    for r in pubmed_only:
        recon_lines.append(f"  [{r['year']}] {r['title'][:100]}")

    recon_path = f"{prefix}_reconciliation.txt"
    Path(recon_path).write_text("\n".join(recon_lines), encoding="utf-8")
    print(f"Wrote reconciliation report to: {recon_path}")
    print("\nDone. Import the CSV into your Publications Master Google Sheet.")
    print("Tip: review the reconciliation report before importing.\n")


if __name__ == "__main__":
    main()
