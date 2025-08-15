#!/usr/bin/env python3
"""
BOM PDF → SOV-of-Variants Excel Generator (Combined Single Sheet)

Features:
- Parses multiple BOM PDFs into one "Combined" sheet
- Hierarchical Part-Type indent columns (1–5 levels)
- Part Name/Part Number/Drawing Number/Rev. all centered and bordered
- Part Name column always 43.55 wide
- ΔX headers per variant (max revision)
"""
import argparse
import logging
import re
import subprocess
import sys
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pdfplumber
import pandas as pd

# ──────────────────────────────────────────────────────────────────────────────
BLANK_WIDTHS = [4.22, 16.67, 43.47, 18.22, 17.78, 6.78, 13.78, 13.78, 13.78]
BLANK_HEIGHTS = [
    82.5, 16.5, 60.8, 16.2, 16.5, 16.5, 25.5, 16.5, 19.2,
    361.5, 33.0, 171.8, 97.1, 81.8
]
SMALL_WIDTH = 4.36
TOTAL_WIDTH = 21.18  # fixed width for Part-Type area
PART_NAME_W = 43.55
START_ROW = 14

JAPANESE_RE = re.compile(r"[\u3000-\u30FF\u4E00-\u9FAF]")
CUT_THRESHOLD = 4


def setup_logging(level: str) -> None:
    """Configure basic logging."""
    lvl = getattr(logging, level.upper(), None)
    if not isinstance(lvl, int):
        raise ValueError(f"Invalid log level: {level}")
    logging.basicConfig(
        format="%(asctime)s %(levelname)-8s %(message)s",
        level=lvl,
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def clean_cell(text: Optional[str]) -> str:
    """Normalize cell text: strip newlines and whitespace."""
    return "" if text is None else str(text).replace("\n", " ").strip()


def strip_non_ascii(s: str) -> str:
    """Remove non-ASCII characters."""
    return re.sub(r"[^\x00-\x7F]+", "", s)


def normalize_text(s: str) -> str:
    """Normalize Unicode text to NFKC form."""
    return unicodedata.normalize("NFKC", s)


def parse_bom_pdf(pdf_path: Path) -> List[Dict]:
    """Extract BOM entries from a PDF into structured records."""
    entries: List[Dict] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables:
                continue
            tbl = tables[0]

            # locate header row
            hdr_i = next(
                (
                    i for i, row in enumerate(tbl[:6])
                    if any(clean_cell(c).lower().startswith("level") for c in row)
                    and any(clean_cell(c).lower().startswith("part number") for c in row)
                ),
                2
            )
            hdr = tbl[hdr_i]

            # build column map
            cmap: Dict[str, int] = {}
            for ci, cell in enumerate(hdr):
                txt = clean_cell(cell).lower()
                if "part number" in txt:
                    cmap['part_number'] = ci
                elif "part name" in txt:
                    cmap['part_name'] = ci
                elif "qty" in txt:
                    cmap['quantity'] = ci
                elif "change" in txt or txt.startswith("rev"):
                    cmap['change'] = ci
            cmap.setdefault('part_number', 6)
            cmap.setdefault('part_name', 19)
            cmap.setdefault('quantity', None)
            cmap.setdefault('change', None)

            # process rows
            for row in tbl[hdr_i + 1:]:
                idx = clean_cell(row[0])
                if not idx or not re.match(r"^\d+(?:\.\d+)?$", idx):
                    continue

                # indent level
                level = next(
                    (ci for ci in range(1, 6) if ci < len(row) and clean_cell(row[ci])),
                    1
                )

                # part name extraction
                raw = row[cmap['part_name']] if cmap['part_name'] < len(row) else ""
                lines = [ln.strip() for ln in str(raw).splitlines() if ln.strip()]
                content = lines[1:] if len(lines) > 1 else []

                # drop second line if Japanese or too short
                if content:
                    nxt = content[0]
                    eng = re.sub(r'[^A-Za-z]', '', nxt)
                    if JAPANESE_RE.search(nxt) or len(eng) < CUT_THRESHOLD:
                        content = content[1:]

                # merge broken fragments (e.g., AS + SY → ASSY)
                merged: List[str] = []
                i = 0
                while i < len(content):
                    curr = content[i]
                    if i + 1 < len(content):
                        a = re.sub(r'[^A-Za-z]', '', curr)
                        b = re.sub(r'[^A-Za-z]', '', content[i + 1])
                        if a.isupper() and b.isupper() and len(a) <= CUT_THRESHOLD and len(b) <= CUT_THRESHOLD:
                            merged.append(a + b)
                            i += 2
                            continue
                    merged.append(curr)
                    i += 1
                content = merged

                # assemble text
                text = " ".join(content) if content else (lines[-1] if lines else "")

                # Camel-case with acronyms preserved
                words = strip_non_ascii(text).split()
                pname = " ".join(
                    w if w.isupper() and len(w) <= 3 else w.lower().capitalize()
                    for w in words
                )

                # part/drawing number
                pn = normalize_text(clean_cell(row[cmap['part_number']]))
                dn = pn

                # revision
                raw_chg = clean_cell(row[cmap['change']]) if cmap['change'] is not None and cmap['change'] < len(row) else ""
                nums = re.findall(r"\d+", raw_chg)
                chg = nums[-1] if nums else "0"

                # quantity
                qt = clean_cell(row[cmap['quantity']]) if cmap['quantity'] is not None and cmap['quantity'] < len(row) else ""
                try:
                    qty = int(qt)
                except ValueError:
                    try:
                        qty = float(qt)
                    except ValueError:
                        qty = 0

                entries.append({
                    'level': level,
                    'part_name': pname,
                    'part_number': pn,
                    'drawing_no': dn,
                    'change': chg,
                    'quantity': qty,
                })

    return entries


def parse_pdf_metadata(pdf_path: Path) -> Tuple[str, str]:
    """Extract customer and product from PDF text without index errors."""
    # Concatenate all page text
    with pdfplumber.open(str(pdf_path)) as pdf:
        full_text = "".join(page.extract_text() or "" for page in pdf.pages)
    # Split into non-empty lines
    lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]
    # Look for the 'PRODUCT NUMBER' line mentioning 'CUSTOMER'
    for i, line in enumerate(lines):
        up = line.upper()
        if up.startswith("PRODUCT NUMBER") and "CUSTOMER" in up:
            if i > 0:
                parts = lines[i-1].split()
                if len(parts) >= 2:
                    prod = strip_non_ascii(normalize_text(parts[0]))
                    cust = strip_non_ascii(normalize_text(parts[1]))
                    return cust, prod
            break
    return "", ""


def write_combined_excel(sheets: List[Tuple[str, List[Dict], str, str]], out_path: Path) -> None:
    """Generate and save the combined Excel sheet."""
    # prepare parts
    parts, order, var_max = {}, [], []
    for _, ents, _, _ in sheets:
        var_max.append(max((int(e['change']) for e in ents), default=0))
    for vi, (_, ents, _, _) in enumerate(sheets):
        for r in ents:
            pn = r['part_number']
            if pn not in parts:
                parts[pn] = {**r, 'qtys': [0] * len(sheets)}
                order.append(pn)
            parts[pn]['qtys'][vi] = r['quantity']

    max_lvl = max(parts[p]['level'] for p in order) if order else 1
    PN_COL = max_lvl
    GROUP_ST = PN_COL + 1
    QTY_ST = GROUP_ST + 3

    writer = pd.ExcelWriter(str(out_path), engine='xlsxwriter')
    wb = writer.book
    ws = wb.add_worksheet('Combined')

    # formats
    base = {'font_name': 'Arial', 'font_size': 12, 'border': 1, 'text_wrap': True}
    merge_fmt = wb.add_format({**base, 'align': 'center', 'valign': 'vcenter'})
    title_fmt = wb.add_format({'font_name': 'Arial', 'font_size': 26, 'align': 'center', 'valign': 'vcenter'})
    rev_fmt = wb.add_format({**base, 'font_name': 'Symbol', 'align': 'center', 'valign': 'vcenter'})
    vert_fmt = wb.add_format({**base, 'rotation': 90, 'align': 'center', 'valign': 'vcenter'})
    left_fmt = wb.add_format({**base, 'align': 'left', 'valign': 'vcenter'})
    blank_fmt = wb.add_format({**base, 'border': 0, 'bg_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter'})
    grey_fmt = wb.add_format({'bg_color': '#D3D3D3', 'border': 0})

    # set dimensions
    for i, h in enumerate(BLANK_HEIGHTS):
        ws.set_row(i, h)
    lvl_ws = [SMALL_WIDTH] * (max_lvl - 1) + [TOTAL_WIDTH - SMALL_WIDTH * (max_lvl - 1)]
    for i, w in enumerate(lvl_ws):
        ws.set_column(i, i, w)
    ws.set_column(PN_COL, PN_COL, PART_NAME_W)
    for j, w in enumerate(BLANK_WIDTHS[3:6]):
        ws.set_column(GROUP_ST + j, GROUP_ST + j, w)

    # title
    ws.merge_range(0, 0, START_ROW - 1, PN_COL, 'Spreadsheet of Variants', title_fmt)

    # merged boxes across all variant columns for rows 1–9 (zero-indexed 0–8)
    for row in range(0, 9):
        ws.merge_range(
            row,
            QTY_ST,
            row,
            QTY_ST + len(sheets) - 1,
            '',
            merge_fmt
        )

    # empty vertical boxes in rows 10–14 (zero-indexed 9–13)
    for row in range(9, START_ROW):
        for vi in range(len(sheets)):
            ws.write(row, QTY_ST + vi, "", vert_fmt)

    # labels
    labels = [
        'Yazaki Assy Drawing Number', 'Yazaki Assy Drawing Rev',
        'Yazaki MTS Part Number', 'Yazaki MTS Rev',
        'Yazaki S-Characteristics Drawing Number', 'Yazaki S-Characteristics Drawing Rev',
        'Customer Drawing Number', 'Customer Drawing Rev',
        'Vehicle Code', 'Customer Part Description',
        'Model Year', 'Manufacturing Plant',
        'Customer Part Number', 'Yazaki Assembly Number'
    ]
    for i, txt in enumerate(labels):
        ws.merge_range(i, GROUP_ST, i, GROUP_ST + 2, txt, merge_fmt)

    # metadata
    for vi, (_, _, cust, prod) in enumerate(sheets):
        ws.write(12, QTY_ST + vi, cust, vert_fmt)
        ws.write(13, QTY_ST + vi, prod, vert_fmt)

    # headers
    ws.merge_range(START_ROW, 0, START_ROW, PN_COL - 1, 'Part Type', merge_fmt)
    ws.write(START_ROW, PN_COL, 'Part Name', merge_fmt)
    ws.write(START_ROW, GROUP_ST, 'Part Number', merge_fmt)
    ws.write(START_ROW, GROUP_ST + 1, 'Drawing Number', merge_fmt)
    ws.write(START_ROW, GROUP_ST + 2, 'Rev.', merge_fmt)
    for vi, mx in enumerate(var_max):
        ws.write(START_ROW, QTY_ST + vi, f'D{mx}', rev_fmt)

    # data rows
    for idx, pn in enumerate(order):
        r = START_ROW + 1 + idx
        rec = parts[pn]
        sc = rec['level'] - 1
        if rec['level'] >= 2:
            for c in range(sc):
                ws.write(r, c, '', blank_fmt)
        ec = PN_COL - 1
        fmt = merge_fmt if rec['level'] == 1 else left_fmt
        if sc < ec:
            ws.merge_range(r, sc, r, ec, rec['part_name'], fmt)
        else:
            ws.write(r, sc, rec['part_name'], fmt)
        ws.write(r, PN_COL, rec['part_name'], merge_fmt)
        ws.write(r, GROUP_ST, rec['part_number'], merge_fmt)
        ws.write(r, GROUP_ST + 1, rec['part_number'], merge_fmt)
        ws.write(r, GROUP_ST + 2, f"D{rec['change']}", rev_fmt)
        for vi, q in enumerate(rec['qtys']):
            ws.write_number(r, QTY_ST + vi, q, merge_fmt)

    # shade unused
    last_r = START_ROW + len(order)
    last_c = QTY_ST + len(sheets) - 1
    for rr in range(START_ROW, last_r + 1):
        for cc in range(last_c + 1, GROUP_ST + 3):
            ws.write(rr, cc, '', grey_fmt)
    for rr in range(last_r + 1, last_r + 1 + len(BLANK_HEIGHTS)):
        for cc in range(last_c + 1):
            ws.write(rr, cc, '', grey_fmt)
    for rr in range(START_ROW, last_r + 1):
        ws.set_row(rr, 32.3)

    writer.close()
    logging.info('Written Combined → %s', out_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description='Generate Combined SOV from BOM PDFs'
    )
    parser.add_argument('-s', '--sheet', action='append', nargs=2,
                        metavar=('VAR', 'PDF'), required=True,
                        help='Variant name + PDF path')
    parser.add_argument('-o', '--output', required=True,
                        help='Output XLSX path')
    parser.add_argument('--log', default='INFO', help='Log level')
    args = parser.parse_args()

    setup_logging(args.log)
    sheets: List[Tuple[str, List[Dict], str, str]] = []
    for var, pdf in args.sheet:
        p = Path(pdf)
        if not p.is_file():
            logging.error('PDF not found: %s', pdf)
            sys.exit(1)
        ents = parse_bom_pdf(p)
        cust, prod = parse_pdf_metadata(p)
        sheets.append((var, ents, cust, prod))

    write_combined_excel(sheets, Path(args.output))

    try:
        if sys.platform == 'win32':
            subprocess.run(['start', args.output], check=False, shell=True)
        elif sys.platform == 'darwin':
            subprocess.run(['open', args.output], check=False)
        else:
            subprocess.run(['xdg-open', args.output], check=False)
    except Exception:
        logging.warning("Couldn't auto-open %s", args.output)


if __name__ == '__main__':
    main()
