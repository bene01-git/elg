#!/usr/bin/env python3
from __future__ import annotations
"""
DPS Document Search Engine
============================
Searches through extracted text files for keywords and returns the file name,
matching sentence, document metadata, and a formatted legal citation.

Reads metadata (case number, company, date, etc.) from the ELG_Database.xlsx
spreadsheet to enrich results with paper numbers and legal citations.

Usage:
    python search_docs.py "keyword"
    python search_docs.py "rate increase" --dir searchable_pdfs/text
    python search_docs.py "capital expenditure" --case 25-G-0073
    python search_docs.py "joint proposal" --exact
    python search_docs.py "climate" --context 2
    python search_docs.py "penalty" --export results.csv

Requirements:
    pip install openpyxl
"""

import argparse
import csv
import os
import re
import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
DEFAULT_TEXT_DIR = os.path.join("searchable_pdfs", "text")
DEFAULT_EXCEL = "ELG_Database.xlsx"


# ---------------------------------------------------------------------------
# Excel metadata loading
# ---------------------------------------------------------------------------

def load_metadata(excel_path: str) -> list:
    """
    Load document metadata from the ELG_Database spreadsheet.
    Returns a list of dicts, one per row.
    """
    if not HAS_OPENPYXL:
        return []
    if not os.path.exists(excel_path):
        return []

    wb = load_workbook(excel_path)
    ws = wb.active
    records = []

    for row_idx in range(2, ws.max_row + 1):
        case_number = str(ws.cell(row=row_idx, column=1).value or "").strip()
        company_name = str(ws.cell(row=row_idx, column=2).value or "").strip()
        company_type = str(ws.cell(row=row_idx, column=3).value or "").strip()
        date_filed = str(ws.cell(row=row_idx, column=4).value or "").strip()
        doc_type = str(ws.cell(row=row_idx, column=5).value or "").strip()
        doc_title = str(ws.cell(row=row_idx, column=6).value or "").strip()
        filing_on_behalf = str(ws.cell(row=row_idx, column=7).value or "").strip()
        file_name = str(ws.cell(row=row_idx, column=8).value or "").strip()
        file_size = str(ws.cell(row=row_idx, column=9).value or "").strip()

        # Get link — check hyperlink first, then cell value
        cell = ws.cell(row=row_idx, column=10)
        link = ""
        if cell.hyperlink:
            link = cell.hyperlink.target or ""
        if not link:
            link = str(cell.value or "")

        # Extract paper/item number from the link if it contains FilingSeq
        item_no = ""
        filing_match = re.search(r'FilingSeq=(\d+)', link)
        if filing_match:
            item_no = filing_match.group(1)

        if not case_number and not doc_title:
            continue

        records.append({
            "case_number": case_number,
            "company_name": company_name,
            "company_type": company_type,
            "date_filed": date_filed,
            "doc_type": doc_type,
            "doc_title": doc_title,
            "filing_on_behalf": filing_on_behalf,
            "file_name": file_name,
            "file_size": file_size,
            "link": link,
            "item_no": item_no,
        })

    wb.close()
    return records


def match_file_to_metadata(filename: str, records: list) -> dict:
    """
    Match a text filename back to an Excel row.
    Filenames follow: {case} - {company} - {doc_type} - {doc_title}.txt
    """
    name = filename.replace(".txt", "").replace(".pdf", "")
    name_lower = name.lower()

    best_match = None
    best_score = 0

    for rec in records:
        score = 0

        # Check if case number is in the filename
        if rec["case_number"] and rec["case_number"].lower() in name_lower:
            score += 3

        # Check if document title is in the filename
        if rec["doc_title"]:
            title_clean = rec["doc_title"][:60].lower()
            if title_clean in name_lower:
                score += 5
            elif title_clean[:30] in name_lower:
                score += 3

        # Check if doc type is in the filename
        if rec["doc_type"] and rec["doc_type"].lower() in name_lower:
            score += 2

        # Check if company name is in the filename
        if rec["company_name"]:
            company_short = rec["company_name"].split(",")[0].lower()
            if company_short in name_lower:
                score += 1

        if score > best_score:
            best_score = score
            best_match = rec

    # Only return if we have a reasonable match
    if best_score >= 5:
        return best_match
    return None


def format_citation(meta: dict) -> str:
    """
    Format a legal citation from document metadata.

    Format:
        Case [case_number], [case_title], [doc_type]: [doc_title],
        filed by [filing_on_behalf] on behalf of [company_name],
        p. [item_no] ([date_filed]).
    """
    parts = []

    # Case number
    if meta.get("case_number"):
        parts.append(f"Case {meta['case_number']}")

    # Document type and title
    if meta.get("doc_type") and meta.get("doc_title"):
        parts.append(f"{meta['doc_type']}: {meta['doc_title']}")
    elif meta.get("doc_title"):
        parts.append(meta["doc_title"])

    # Filing party and company
    if meta.get("filing_on_behalf"):
        parts.append(f"filed by {meta['filing_on_behalf']}")
    if meta.get("company_name"):
        company = meta["company_name"]
        # Don't repeat if filing_on_behalf is the same as company
        if meta.get("filing_on_behalf", "").lower() != company.lower():
            parts.append(f"on behalf of {company}")

    # Paper/Item number
    if meta.get("item_no"):
        parts.append(f"Paper No. {meta['item_no']}")

    # Date
    if meta.get("date_filed"):
        parts.append(f"({meta['date_filed']})")

    return ", ".join(parts)


# ---------------------------------------------------------------------------
# Text splitting
# ---------------------------------------------------------------------------

def split_sentences(text: str) -> list:
    """Split text into sentences."""
    text = re.sub(r'---\s*Page\s+\d+\s*---', ' ', text)
    text = re.sub(r'\n+', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()

    sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z"])', text)

    result = []
    for s in sentences:
        s = s.strip()
        if not s:
            continue
        if len(s) > 1000:
            parts = re.split(r'[;:]\s+', s)
            result.extend(p.strip() for p in parts if p.strip())
        else:
            result.append(s)

    return result


# ---------------------------------------------------------------------------
# Search logic
# ---------------------------------------------------------------------------

def search_file(filepath: str, keyword: str, exact: bool, context: int) -> list:
    """Search a single text file for a keyword."""
    try:
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            text = f.read()
    except Exception:
        return []

    if not text.strip():
        return []

    sentences = split_sentences(text)
    matches = []

    if exact:
        pattern = re.compile(re.escape(keyword), re.IGNORECASE)
    else:
        words = keyword.lower().split()
        pattern = None

    for i, sentence in enumerate(sentences):
        found = False

        if exact:
            if pattern.search(sentence):
                found = True
        else:
            sentence_lower = sentence.lower()
            if all(w in sentence_lower for w in words):
                found = True

        if found:
            start = max(0, i - context)
            end = min(len(sentences), i + context + 1)
            context_sentences = sentences[start:end]

            matches.append({
                "sentence": sentence.strip(),
                "context": " ".join(context_sentences).strip(),
                "sentence_num": i + 1,
            })

    return matches


def search_all_files(text_dir: str, keyword: str, exact: bool, context: int,
                     case_filter: str = None) -> list:
    """Search all .txt files in a directory."""
    results = []

    if not os.path.isdir(text_dir):
        print(f"[ERROR] Directory not found: {text_dir}")
        sys.exit(1)

    txt_files = sorted(Path(text_dir).glob("*.txt"))

    if not txt_files:
        print(f"[ERROR] No .txt files found in: {text_dir}")
        sys.exit(1)

    for filepath in txt_files:
        filename = filepath.name

        if case_filter and case_filter.lower() not in filename.lower():
            continue

        matches = search_file(str(filepath), keyword, exact, context)

        if matches:
            results.append({
                "file": filename,
                "path": str(filepath),
                "matches": matches,
            })

    return results


# ---------------------------------------------------------------------------
# Display
# ---------------------------------------------------------------------------

def highlight(text: str, keyword: str, exact: bool) -> str:
    """Add >>> <<< markers around the keyword."""
    if exact:
        pattern = re.compile(f'({re.escape(keyword)})', re.IGNORECASE)
        return pattern.sub(r'>>>\1<<<', text)
    else:
        words = keyword.split()
        result = text
        for word in words:
            pattern = re.compile(f'({re.escape(word)})', re.IGNORECASE)
            result = pattern.sub(r'>>>\1<<<', result)
        return result


def display_results(results: list, keyword: str, exact: bool, show_context: bool,
                    metadata_records: list):
    """Print search results with metadata and citations."""
    total_matches = sum(len(r["matches"]) for r in results)
    total_files = len(results)

    print()
    print(f"  Found {total_matches} match(es) across {total_files} file(s)")
    print()

    for r in results:
        # Try to match this file to Excel metadata
        meta = match_file_to_metadata(r["file"], metadata_records)

        print("=" * 78)
        print(f"  FILE: {r['file']}")

        if meta:
            print(f"  Case:       {meta.get('case_number', 'N/A')}")
            print(f"  Company:    {meta.get('company_name', 'N/A')}")
            print(f"  Doc Type:   {meta.get('doc_type', 'N/A')}")
            print(f"  Doc Title:  {meta.get('doc_title', 'N/A')}")
            print(f"  Filed By:   {meta.get('filing_on_behalf', 'N/A')}")
            print(f"  Date Filed: {meta.get('date_filed', 'N/A')}")
            if meta.get("item_no"):
                print(f"  Paper No:   {meta['item_no']}")
            print()
            print(f"  CITATION: {format_citation(meta)}")
        else:
            print(f"  (No metadata found — ensure ELG_Database.xlsx is in the same folder)")

        print(f"  Matches: {len(r['matches'])}")
        print("-" * 78)

        for j, m in enumerate(r["matches"], 1):
            if show_context:
                display_text = m["context"]
            else:
                display_text = m["sentence"]

            highlighted = highlight(display_text, keyword, exact)
            print(f"\n  [{j}] {highlighted}")

        print()


def export_results(results: list, keyword: str, output_path: str,
                   metadata_records: list):
    """Export results to a CSV file with metadata and citations."""
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            "File",
            "Case Number",
            "Company Name",
            "Document Type",
            "Document Title",
            "Filing On Behalf Of",
            "Date Filed",
            "Paper No.",
            "Citation",
            "Match #",
            "Sentence",
        ])

        for r in results:
            meta = match_file_to_metadata(r["file"], metadata_records)

            for j, m in enumerate(r["matches"], 1):
                if meta:
                    writer.writerow([
                        r["file"],
                        meta.get("case_number", ""),
                        meta.get("company_name", ""),
                        meta.get("doc_type", ""),
                        meta.get("doc_title", ""),
                        meta.get("filing_on_behalf", ""),
                        meta.get("date_filed", ""),
                        meta.get("item_no", ""),
                        format_citation(meta),
                        j,
                        m["sentence"],
                    ])
                else:
                    writer.writerow([
                        r["file"], "", "", "", "", "", "", "", "",
                        j, m["sentence"],
                    ])

    print(f"  Results exported to: {output_path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Search DPS document text files with metadata and legal citations.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python search_docs.py "rate increase"
  python search_docs.py "capital expenditure" --case 25-G-0073
  python search_docs.py "joint proposal" --exact
  python search_docs.py "climate" --context 2
  python search_docs.py "penalty" --export results.csv
  python search_docs.py "penalty" --excel path/to/ELG_Database.xlsx
        """,
    )
    parser.add_argument("keyword", help="Keyword or phrase to search for")
    parser.add_argument(
        "--dir", "-d", default=DEFAULT_TEXT_DIR,
        help=f"Directory containing .txt files (default: {DEFAULT_TEXT_DIR})",
    )
    parser.add_argument(
        "--excel", default=DEFAULT_EXCEL,
        help=f"Path to ELG_Database.xlsx for metadata (default: {DEFAULT_EXCEL})",
    )
    parser.add_argument(
        "--case", "-c", default=None,
        help="Only search files for a specific case number",
    )
    parser.add_argument(
        "--exact", "-e", action="store_true",
        help="Match the exact phrase (default: all words any order)",
    )
    parser.add_argument(
        "--context", type=int, default=0,
        help="Number of surrounding sentences to show (default: 0)",
    )
    parser.add_argument(
        "--export", default=None,
        help="Export results to a CSV file",
    )
    args = parser.parse_args()

    # Load metadata from Excel
    metadata_records = []
    if os.path.exists(args.excel):
        if HAS_OPENPYXL:
            print(f"  Loading metadata from {args.excel}...")
            metadata_records = load_metadata(args.excel)
            print(f"  Loaded {len(metadata_records)} document records.")
        else:
            print("  [NOTE] openpyxl not installed — metadata/citations unavailable.")
            print("         Install with: pip install openpyxl")
    else:
        print(f"  [NOTE] {args.excel} not found — metadata/citations unavailable.")

    print()
    print("=" * 78)
    print("  DPS Document Search Engine")
    print("=" * 78)
    print(f"  Keyword:   \"{args.keyword}\"")
    print(f"  Directory: {args.dir}")
    print(f"  Mode:      {'Exact phrase' if args.exact else 'All words (any order)'}")
    if args.case:
        print(f"  Case:      {args.case}")
    if args.context > 0:
        print(f"  Context:   {args.context} surrounding sentence(s)")

    results = search_all_files(
        text_dir=args.dir,
        keyword=args.keyword,
        exact=args.exact,
        context=args.context,
        case_filter=args.case,
    )

    if not results:
        print(f"\n  No matches found for \"{args.keyword}\".")
        sys.exit(0)

    display_results(results, args.keyword, args.exact,
                    show_context=args.context > 0,
                    metadata_records=metadata_records)

    if args.export:
        export_results(results, args.keyword, args.export, metadata_records)


if __name__ == "__main__":
    main()
