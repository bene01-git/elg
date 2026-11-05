#!/usr/bin/env python3
from __future__ import annotations
"""
Build Search Index — with Inverted Index, TF-IDF data, and Suggestions
========================================================================
Outputs four files into site/:
  - metadata.json    (small — doc info, citations)
  - sentences.json   (large — sentences keyed by doc id)
  - word_index.json  (medium — inverted index: word → doc ids + doc frequency)
  - suggestions.json (tiny — top search terms for autocomplete)

Usage:
    python build_index.py
    python build_index.py --text-dir searchable_pdfs/text --output-dir site
"""

import argparse
import json
import math
import os
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

try:
    from openpyxl import load_workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

DEFAULT_TEXT_DIR = os.path.join("searchable_pdfs", "text")
DEFAULT_EXCEL = "ELG_Database.xlsx"
DEFAULT_OUTPUT_DIR = "site"

# Common English stop words to exclude from the index and suggestions
STOP_WORDS = frozenset("""
a an and are as at be by for from has have in is it its of on or that the
this to was were will with not but they them their than also been can could
did does each had has have how into may more most must no nor only other
our out own said same she should so some such than that the their then
there these they this those through too under until very was we were what
when where which while who whom why will with would you your about after
all also am among any are before being both between came come could did
do does during each few get got had has have here him his how if into its
just let like made make many me might more much must my never next now
of off often old only other our out over own per put quite rather really
right round said see seem shall she should show side since so some still
such take tell than that the their them then there these they thing think
this those though through time to too two under until up upon us use used
very want was way we well went were what when where which while who why
will with without work would year you your upon shall thereof herein
hereof pursuant provided section subsection paragraph
""".split())

MIN_WORD_LENGTH = 3
MIN_DOC_FREQUENCY = 1       # word must appear in at least this many docs
MAX_DOC_FREQUENCY_PCT = 0.9  # skip words in >90% of docs (too common)
NUM_SUGGESTIONS = 800


# ---------------------------------------------------------------------------
# Metadata loading (same as before)
# ---------------------------------------------------------------------------

def load_metadata(excel_path):
    if not HAS_OPENPYXL or not os.path.exists(excel_path):
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
        cell = ws.cell(row=row_idx, column=10)
        link = ""
        if cell.hyperlink:
            link = cell.hyperlink.target or ""
        if not link:
            link = str(cell.value or "")
        item_no = ""
        m = re.search(r'FilingSeq=(\d+)', link)
        if m:
            item_no = m.group(1)
        if not case_number and not doc_title:
            continue
        records.append({
            "case_number": case_number, "company_name": company_name,
            "company_type": company_type, "date_filed": date_filed,
            "doc_type": doc_type, "doc_title": doc_title,
            "filing_on_behalf": filing_on_behalf, "link": link, "item_no": item_no,
        })
    wb.close()
    return records


def match_file_to_metadata(filename, records):
    name = filename.replace(".txt", "").replace(".pdf", "").lower()
    best_match = None
    best_score = 0
    for rec in records:
        score = 0
        if rec["case_number"] and rec["case_number"].lower() in name:
            score += 3
        if rec["doc_title"]:
            t = rec["doc_title"][:60].lower()
            if t in name:
                score += 5
            elif t[:30] in name:
                score += 3
        if rec["doc_type"] and rec["doc_type"].lower() in name:
            score += 2
        if rec["company_name"]:
            c = rec["company_name"].split(",")[0].lower()
            if c in name:
                score += 1
        if score > best_score:
            best_score = score
            best_match = rec
    return best_match if best_score >= 5 else None


def format_citation(meta):
    parts = []
    if meta.get("case_number"):
        parts.append(f"Case {meta['case_number']}")
    if meta.get("doc_type") and meta.get("doc_title"):
        parts.append(f"{meta['doc_type']}: {meta['doc_title']}")
    elif meta.get("doc_title"):
        parts.append(meta["doc_title"])
    if meta.get("filing_on_behalf"):
        parts.append(f"filed by {meta['filing_on_behalf']}")
    if meta.get("company_name"):
        if meta.get("filing_on_behalf", "").lower() != meta["company_name"].lower():
            parts.append(f"on behalf of {meta['company_name']}")
    if meta.get("item_no"):
        parts.append(f"Paper No. {meta['item_no']}")
    if meta.get("date_filed"):
        parts.append(f"({meta['date_filed']})")
    return ", ".join(parts)


def split_sentences(text):
    text = re.sub(r'---\s*Page\s+\d+\s*---', ' ', text)
    text = re.sub(r'\n+', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z"])', text)
    result = []
    for s in sentences:
        s = s.strip()
        if not s or len(s) < 10:
            continue
        if len(s) > 800:
            parts = re.split(r'[;:]\s+', s)
            result.extend(p.strip() for p in parts if p.strip() and len(p.strip()) >= 10)
        else:
            result.append(s)
    return result


def tokenize(text):
    """Split text into lowercase words, filtering out stop words and short words."""
    words = re.findall(r'[a-zA-Z]+', text.lower())
    return [w for w in words if len(w) >= MIN_WORD_LENGTH and w not in STOP_WORDS]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Build search index files.")
    parser.add_argument("--text-dir", default=DEFAULT_TEXT_DIR)
    parser.add_argument("--excel", default=DEFAULT_EXCEL)
    parser.add_argument("--output-dir", "-o", default=DEFAULT_OUTPUT_DIR)
    args = parser.parse_args()

    if not os.path.isdir(args.text_dir):
        print(f"[ERROR] Text directory not found: {args.text_dir}")
        sys.exit(1)

    print(f"Loading metadata from {args.excel}...")
    records = load_metadata(args.excel)
    print(f"Loaded {len(records)} metadata records.")

    txt_files = sorted(Path(args.text_dir).glob("*.txt"))
    print(f"Found {len(txt_files)} text files.")

    metadata_list = []
    sentences_dict = {}

    # For inverted index
    inverted = defaultdict(set)   # word → set of doc_id strings
    word_counts = Counter()       # global word frequency for suggestions
    doc_word_counts = {}          # doc_id → total word count (for TF-IDF)

    for i, filepath in enumerate(txt_files):
        try:
            with open(filepath, "r", encoding="utf-8", errors="replace") as f:
                content = f.read()
        except Exception:
            continue

        if not content.strip():
            continue

        sentences = split_sentences(content)
        if not sentences:
            continue

        meta = match_file_to_metadata(filepath.name, records)
        citation = format_citation(meta) if meta else ""

        doc_id = str(i)

        # Metadata
        metadata_list.append({
            "id": doc_id,
            "filename": filepath.name,
            "case_number": meta.get("case_number", "") if meta else "",
            "company_name": meta.get("company_name", "") if meta else "",
            "doc_type": meta.get("doc_type", "") if meta else "",
            "doc_title": meta.get("doc_title", "") if meta else "",
            "filing_on_behalf": meta.get("filing_on_behalf", "") if meta else "",
            "date_filed": meta.get("date_filed", "") if meta else "",
            "item_no": meta.get("item_no", "") if meta else "",
            "link": meta.get("link", "") if meta else "",
            "citation": citation,
        })

        # Sentences
        sentences_dict[doc_id] = sentences

        # Tokenize for inverted index
        all_words = []
        for s in sentences:
            all_words.extend(tokenize(s))

        doc_word_counts[doc_id] = len(all_words)
        unique_words = set(all_words)
        for w in unique_words:
            inverted[w].add(doc_id)
        word_counts.update(all_words)

        if (i + 1) % 50 == 0:
            print(f"  Processed {i + 1}/{len(txt_files)} files...")

    total_docs = len(metadata_list)
    max_df = int(total_docs * MAX_DOC_FREQUENCY_PCT)

    # Build final inverted index — filter out too-rare and too-common words
    print(f"\nBuilding inverted index from {len(inverted)} unique terms...")
    word_index = {}
    for word, doc_ids in inverted.items():
        df = len(doc_ids)
        if df < MIN_DOC_FREQUENCY:
            continue
        if df > max_df and total_docs > 10:
            continue
        word_index[word] = {
            "df": df,
            "docs": sorted(doc_ids),
        }
    print(f"Index contains {len(word_index)} terms after filtering.")

    # Build suggestion list — top N terms by frequency (excluding filtered ones)
    print("Building autocomplete suggestions...")
    suggestion_candidates = [
        (word, count) for word, count in word_counts.most_common()
        if word in word_index and len(word) >= 4
    ]
    suggestions = [w for w, _ in suggestion_candidates[:NUM_SUGGESTIONS]]

    # ---- Write output files ----
    os.makedirs(args.output_dir, exist_ok=True)

    # 1. metadata.json
    meta_path = os.path.join(args.output_dir, "metadata.json")
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump({"documents": metadata_list, "count": total_docs}, f)
    print(f"  metadata.json    {os.path.getsize(meta_path) / 1024:.0f} KB")

    # 2. sentences.json
    sent_path = os.path.join(args.output_dir, "sentences.json")
    with open(sent_path, "w", encoding="utf-8") as f:
        json.dump(sentences_dict, f)
    print(f"  sentences.json   {os.path.getsize(sent_path) / (1024*1024):.1f} MB")

    # 3. word_index.json
    idx_path = os.path.join(args.output_dir, "word_index.json")
    with open(idx_path, "w", encoding="utf-8") as f:
        json.dump({
            "total_docs": total_docs,
            "doc_word_counts": doc_word_counts,
            "index": word_index,
        }, f)
    print(f"  word_index.json  {os.path.getsize(idx_path) / 1024:.0f} KB")

    # 4. suggestions.json
    sug_path = os.path.join(args.output_dir, "suggestions.json")
    with open(sug_path, "w", encoding="utf-8") as f:
        json.dump(suggestions, f)
    print(f"  suggestions.json {os.path.getsize(sug_path) / 1024:.0f} KB")

    # Clean up old single file
    old_path = os.path.join(args.output_dir, "search_data.json")
    if os.path.exists(old_path):
        os.remove(old_path)

    print(f"\nDone. {total_docs} documents indexed.")


if __name__ == "__main__":
    main()
