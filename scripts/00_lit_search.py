"""Automated PubMed search for TB delay literature."""
from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import sys
import time
from pathlib import Path
from typing import Dict, Iterable, List, Optional

from Bio import Entrez

PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PATH = PROJECT_ROOT / "lit" / "lit_db.csv"
LOG_PATH = PROJECT_ROOT / "lit" / "lit_search.log"
DEFAULT_QUERY = (
    "(tuberculosis[MeSH Terms] OR tuberculosis[Title/Abstract]) AND "
    "(delay OR timeliness OR diagnosis delay OR treatment delay) AND India"
)
RATE_LIMIT_SECONDS = 0.34  # NCBI recommends <= 3 requests/sec with API key


def configure_logging() -> logging.Logger:
    """Configure root logger for the script."""
    logger = logging.getLogger("lit_search")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )
    stream = logging.StreamHandler(sys.stdout)
    stream.setFormatter(formatter)
    file_handler = logging.FileHandler(LOG_PATH)
    file_handler.setFormatter(formatter)
    logger.addHandler(stream)
    logger.addHandler(file_handler)
    return logger


def ensure_output_file(headers: Iterable[str]) -> None:
    """Create the output CSV with headers if it is missing."""
    if OUTPUT_PATH.exists() and OUTPUT_PATH.stat().st_size > 0:
        return
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT_PATH.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(headers))
        writer.writeheader()


def load_entrez_credentials(logger: logging.Logger) -> bool:
    """Set Entrez credentials from environment variables."""
    email = os.environ.get("ENTREZ_EMAIL")
    if not email:
        logger.error("ENTREZ_EMAIL not set. Set it to comply with NCBI policies.")
        return False
    Entrez.email = email
    api_key = os.environ.get("ENTREZ_API_KEY")
    if api_key:
        Entrez.api_key = api_key
        logger.info("Using Entrez API key for higher rate limits.")
    return True


def chunked(iterable: List[str], size: int) -> Iterable[List[str]]:
    """Yield successive chunks from a list."""
    for i in range(0, len(iterable), size):
        yield iterable[i : i + size]


def fetch_pmids(query: str, retmax: int, logger: logging.Logger) -> List[str]:
    """Fetch PMIDs for a given query."""
    handle = Entrez.esearch(db="pubmed", term=query, retmax=retmax)
    results = Entrez.read(handle)
    handle.close()
    ids = results.get("IdList", [])
    logger.info("Retrieved %s PMIDs for query.", len(ids))
    return ids


def parse_authors(article: Dict) -> str:
    """Concatenate author names from PubMed article."""
    authors = []
    auth_list = (
        article.get("MedlineCitation", {})
        .get("Article", {})
        .get("AuthorList", [])
    )
    for author in auth_list:
        last = author.get("LastName")
        fore = author.get("ForeName") or author.get("Initials")
        if last and fore:
            authors.append(f"{fore} {last}")
        elif last:
            authors.append(last)
    return "; ".join(authors)


def parse_article(article: Dict) -> Dict[str, Optional[str]]:
    """Extract selected fields from a PubMed article record."""
    citation = article.get("MedlineCitation", {})
    article_data = citation.get("Article", {})
    abstract = article_data.get("Abstract", {}).get("AbstractText")
    if isinstance(abstract, list):
        abstract = " ".join(str(x) for x in abstract)
    journal = article_data.get("Journal", {}).get("Title")
    article_ids = citation.get("PMID", {})
    pmid = article_ids if isinstance(article_ids, str) else article_ids.get("#text")
    year = None
    date_created = citation.get("DateCompleted") or citation.get("DateCreated")
    if date_created:
        year = date_created.get("Year")
    if not year:
        journal_issue = article_data.get("Journal", {}).get("JournalIssue", {})
        pub_date = journal_issue.get("PubDate", {})
        year = pub_date.get("Year")
    return {
        "pmid": pmid,
        "title": article_data.get("ArticleTitle"),
        "journal": journal,
        "year": year,
        "authors": parse_authors(article),
        "abstract": abstract,
    }


def fetch_article_details(pmids: List[str], logger: logging.Logger) -> List[Dict[str, str]]:
    """Fetch article metadata for a set of PMIDs."""
    records: List[Dict[str, str]] = []
    for batch in chunked(pmids, 100):
        logger.info("Fetching metadata for PMIDs %s", ",".join(batch))
        with Entrez.efetch(
            db="pubmed", id=",".join(batch), rettype="medline", retmode="xml"
        ) as handle:
            fetch_result = Entrez.read(handle)
        articles = fetch_result.get("PubmedArticle", [])
        for article in articles:
            records.append(parse_article(article))
        time.sleep(RATE_LIMIT_SECONDS)
    return records


def save_records(records: List[Dict[str, str]]) -> None:
    """Save the records to the CSV database, replacing existing content."""
    ensure_output_file(
        ["pmid", "title", "journal", "year", "authors", "abstract", "query"]
    )
    with OUTPUT_PATH.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(
            fh,
            fieldnames=[
                "pmid",
                "title",
                "journal",
                "year",
                "authors",
                "abstract",
                "query",
            ],
        )
        writer.writeheader()
        for record in records:
            writer.writerow(record)


def parse_arguments() -> argparse.Namespace:
    """Parse CLI arguments."""
    parser = argparse.ArgumentParser(
        description="Query PubMed for TB delay literature and save metadata."
    )
    parser.add_argument(
        "--query",
        default=DEFAULT_QUERY,
        help="Custom Entrez query string.",
    )
    parser.add_argument(
        "--retmax",
        type=int,
        default=200,
        help="Maximum number of PubMed records to retrieve.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Skip API calls and log the query only.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_arguments()
    logger = configure_logging()
    logger.info("Starting PubMed search (retmax=%s)", args.retmax)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)

    if not load_entrez_credentials(logger):
        ensure_output_file(
            ["pmid", "title", "journal", "year", "authors", "abstract", "query"]
        )
        logger.warning("Skipping PubMed call because credentials are missing.")
        return

    if args.dry_run:
        logger.info("Dry run mode: query would be '%s'", args.query)
        return

    try:
        pmids = fetch_pmids(args.query, args.retmax, logger)
        if not pmids:
            logger.warning("No PMIDs returned for the query.")
            ensure_output_file(
                [
                    "pmid",
                    "title",
                    "journal",
                    "year",
                    "authors",
                    "abstract",
                    "query",
                ]
            )
            return
        records = fetch_article_details(pmids, logger)
        for record in records:
            record["query"] = args.query
        save_records(records)
        logger.info("Saved %s PubMed records to %s", len(records), OUTPUT_PATH)
    except Exception as exc:  # noqa: BLE001
        logger.exception("Failed to complete PubMed search: %s", exc)
        ensure_output_file(
            ["pmid", "title", "journal", "year", "authors", "abstract", "query"]
        )


if __name__ == "__main__":
    main()
