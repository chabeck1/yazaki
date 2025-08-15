#!/usr/bin/env python3
"""
Extract and Sort Data from YC SOV PDF

Story (Software): Extract and Sort Data from YC SOV pdf
Doesn't Story (Software): Format Data to YNA SOV Excel Sheet

Usage:
  python extract_sort_sov.py input.pdf [-o output.csv]

Dependencies:
  pip install pdfplumber
"""
import argparse
import csv
import re
import sys
from pathlib import Path

import pdfplumber

def parse_sov_pdf(pdf_path: Path):
    """
    Parse the first table on each page of the SOV PDF,
    returning a list of dicts with part_number, part_name, drawing_no, quantity.
    """
    entries = []
    for page in pdfplumber.open(str(pdf_path)).pages:
        tables = page.extract_tables()
        if not tables:
            continue
        tbl = tables[0]
        # locate header row
        hdr_i = next(
            (i for i, row in enumerate(tbl[:6])
             if any(c and re.search(r'(?i)\blevel\b', str(c)) for c in row)
             and any(c and re.search(r'(?i)\bpart number\b', str(c)) for c in row)),
            0
        )
        hdr = tbl[hdr_i]
        # map columns
        idx = {}
        for ci, cell in enumerate(hdr):
            txt = str(cell or "").lower()
            if "part number" in txt:
                idx['part_number'] = ci
            elif "part name" in txt:
                idx['part_name'] = ci
            elif any(x in txt for x in ("draw.no","draw no","draw.")):
                idx['drawing_no'] = ci
            elif "qty" in txt:
                idx['quantity'] = ci
        # defaults
        idx.setdefault('part_number', 1)
        idx.setdefault('part_name', 2)
        idx.setdefault('drawing_no', idx['part_number'])
        idx.setdefault('quantity', 3)

        # parse rows
        for row in tbl[hdr_i+1:]:
            pn = row[idx['part_number']]
            if not pn or not str(pn).strip():
                continue
            entries.append({
                'part_number': str(pn).strip(),
                'part_name':   str(row[idx['part_name']] or "").strip(),
                'drawing_no':  str(row[idx['drawing_no']] or "").strip(),
                'quantity':    str(row[idx['quantity']] or "").strip(),
            })
    return entries

def main():
    parser = argparse.ArgumentParser(description="Extract & sort data from YC SOV PDF")
    parser.add_argument('pdf', type=Path, help='Input SOV PDF')
    parser.add_argument('-o', '--output', type=Path, help='Output CSV file (default stdout)')
    args = parser.parse_args()

    entries = parse_sov_pdf(args.pdf)
    # sort by part_number
    entries.sort(key=lambda e: e['part_number'])

    # write CSV
    out = sys.stdout
    if args.output:
        out = open(args.output, 'w', newline='', encoding='utf-8')
    writer = csv.DictWriter(out, fieldnames=['part_number','part_name','drawing_no','quantity'])
    writer.writeheader()
    for e in entries:
        writer.writerow(e)
    if args.output:
        out.close()

if __name__ == "__main__":
    main()
