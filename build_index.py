#!/usr/bin/env python3
from __future__ import annotations
"""
Build Search Index — Chunked Architecture
===========================================
Outputs to site/:
  - metadata.json      (small — doc info, citations)
  - word_index.json    (medium — inverted index for fast lookup)
  - suggestions.json   (tiny — autocomplete terms)
  - chunks/0.json      (sentences for docs 0–29)
  - chunks/1.json      (sentences for docs 30–59)
  - ...

Sentences are split into small chunk files so no single file is too large
for GitHub Pages to handle.

Usage:
    python build_index.py
    python build_index.py --text-dir searchable_pdfs/text --output-dir site
"""

import argparse
import json
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
CHUNK_SIZE = 30  # docs per chunk file

STOP_WORDS = frozenset("""
a an and are as at be by for from has have in is it its of on or that the
this to was were will with not but they them their than also been can could
did does each had how into may more most must nor only other our out own
said same she should so some such then there these those through too under
until very we what when where which while who whom why would you your about
after all before both between come get here him his just like made make
many much never now off often upon shall thereof herein hereof pursuant
provided section subsection paragraph
""".split())

MIN_WORD_LENGTH = 3
MAX_DOC_FREQUENCY_PCT = 0.9
NUM_SUGGESTIONS = 800


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
    best = None
    best_score = 0
    for rec in records:
        s = 0
        if rec["case_number"] and rec["case_number"].lower() in name: s += 3
        if rec["doc_title"]:
            t = rec["doc_title"][:60].lower()
            if t in name: s += 5
            elif t[:30] in name: s += 3
        if rec["doc_type"] and rec["doc_type"].lower() in name: s += 2
        if rec["company_name"]:
            c = rec["company_name"].split(",")[0].lower()
            if c in name: s += 1
        if s > best_score:
            best_score = s
            best = rec
    return best if best_score >= 5 else None


def format_citation(meta):
    parts = []
    if meta.get("case_number"): parts.append(f"Case {meta['case_number']}")
    if meta.get("doc_type") and meta.get("doc_title"):
        parts.append(f"{meta['doc_type']}: {meta['doc_title']}")
    elif meta.get("doc_title"): parts.append(meta["doc_title"])
    if meta.get("filing_on_behalf"): parts.append(f"filed by {meta['filing_on_behalf']}")
    if meta.get("company_name"):
        if meta.get("filing_on_behalf", "").lower() != meta["company_name"].lower():
            parts.append(f"on behalf of {meta['company_name']}")
    if meta.get("item_no"): parts.append(f"Paper No. {meta['item_no']}")
    if meta.get("date_filed"): parts.append(f"({meta['date_filed']})")
    return ", ".join(parts)


def split_sentences(text):
    text = re.sub(r'---\s*Page\s+\d+\s*---', ' ', text)
    text = re.sub(r'\n+', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z"])', text)
    result = []
    for s in sentences:
        s = s.strip()
        if not s or len(s) < 10: continue
        if len(s) > 800:
            result.extend(p.strip() for p in re.split(r'[;:]\s+', s) if len(p.strip()) >= 10)
        else:
            result.append(s)
    return result


def tokenize(text):
    words = re.findall(r'[a-zA-Z]+', text.lower())
    return [w for w in words if len(w) >= MIN_WORD_LENGTH and w not in STOP_WORDS]


def main():
    parser = argparse.ArgumentParser(description="Build chunked search index.")
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
    all_sentences = {}       # doc_id → sentences (temporary, written to chunks)
    inverted = defaultdict(set)
    word_counts = Counter()
    doc_word_counts = {}

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
        chunk_num = i // CHUNK_SIZE

        metadata_list.append({
            "id": doc_id,
            "chunk": chunk_num,
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

        all_sentences[doc_id] = sentences

        all_words = []
        for s in sentences:
            all_words.extend(tokenize(s))
        doc_word_counts[doc_id] = len(all_words)
        for w in set(all_words):
            inverted[w].add(doc_id)
        word_counts.update(all_words)

        if (i + 1) % 50 == 0:
            print(f"  Processed {i + 1}/{len(txt_files)} files...")

    total_docs = len(metadata_list)
    max_df = int(total_docs * MAX_DOC_FREQUENCY_PCT)

    # Build inverted index
    print(f"\nBuilding inverted index from {len(inverted)} terms...")
    word_index = {}
    for word, doc_ids in inverted.items():
        df = len(doc_ids)
        if df > max_df and total_docs > 10:
            continue
        word_index[word] = {"df": df, "docs": sorted(doc_ids)}
    print(f"Index: {len(word_index)} terms after filtering.")

    # Suggestions
    suggestions = [w for w, _ in word_counts.most_common()
                   if w in word_index and len(w) >= 4][:NUM_SUGGESTIONS]

    # ---- Write files ----
    os.makedirs(args.output_dir, exist_ok=True)
    chunks_dir = os.path.join(args.output_dir, "chunks")
    os.makedirs(chunks_dir, exist_ok=True)

    # Clear old chunks
    for old in Path(chunks_dir).glob("*.json"):
        old.unlink()

    # 1. metadata.json
    p = os.path.join(args.output_dir, "metadata.json")
    with open(p, "w", encoding="utf-8") as f:
        json.dump({"documents": metadata_list, "count": total_docs}, f)
    print(f"  metadata.json      {os.path.getsize(p) / 1024:.0f} KB")

    # 2. word_index.json
    p = os.path.join(args.output_dir, "word_index.json")
    with open(p, "w", encoding="utf-8") as f:
        json.dump({
            "total_docs": total_docs,
            "doc_word_counts": doc_word_counts,
            "index": word_index,
        }, f)
    print(f"  word_index.json    {os.path.getsize(p) / 1024:.0f} KB")

    # 3. suggestions.json
    p = os.path.join(args.output_dir, "suggestions.json")
    with open(p, "w", encoding="utf-8") as f:
        json.dump(suggestions, f)
    print(f"  suggestions.json   {os.path.getsize(p) / 1024:.0f} KB")

    # 4. Chunked sentences
    num_chunks = (total_docs + CHUNK_SIZE - 1) // CHUNK_SIZE
    total_chunk_size = 0
    for c in range(num_chunks):
        chunk_data = {}
        for doc_id, sents in all_sentences.items():
            if int(doc_id) // CHUNK_SIZE == c:
                chunk_data[doc_id] = sents
        cp = os.path.join(chunks_dir, f"{c}.json")
        with open(cp, "w", encoding="utf-8") as f:
            json.dump(chunk_data, f)
        total_chunk_size += os.path.getsize(cp)

    print(f"  chunks/            {num_chunks} files, {total_chunk_size / (1024*1024):.1f} MB total")

    # Clean up old monolithic sentences.json
    for old_file in ["sentences.json", "search_data.json"]:
        old_path = os.path.join(args.output_dir, old_file)
        if os.path.exists(old_path):
            os.remove(old_path)
            print(f"  Removed old {old_file}")

    print(f"\nDone. {total_docs} documents indexed across {num_chunks} chunks.")


if __name__ == "__main__":
    main()
