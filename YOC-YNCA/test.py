#!/usr/bin/env python3
"""
BOM PDF → SOV-of-Variants Excel Generator (Combined Single Sheet)

Highlights / flags:
- Yellow  : Part Name (and Part Type) when we had to stitch multiple ENGLISH lines together
- Light red: Part Name when NOTE/備考 is used as the visible name
- Lavender : Both conditions above true

Only multi-line *English* (post-filter) triggers yellow.
"""

import argparse
import logging
import re
import subprocess
import sys
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Set

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
START_ROW = 14  # zero-indexed: Excel row 15

JAPANESE_RE = re.compile(r"[\u3000-\u30FF\u4E00-\u9FAF]")
CUT_THRESHOLD = 3

# ───────────── NOTE handling ─────────────
_TRIVIAL_NOTE_REGEX = re.compile(
    r"^\s*(欠図\s*)?(in\s+preparation)\s*$",
    re.IGNORECASE
)

def setup_logging(level: str) -> None:
    lvl = getattr(logging, level.upper(), None)
    if not isinstance(lvl, int):
        raise ValueError(f"Invalid log level: {level}")
    logging.basicConfig(
        format="%(asctime)s %(levelname)-8s %(message)s",
        level=lvl,
        datefmt="%Y-%m-%d %H:%M:%S",
    )

def clean_cell(text: Optional[str]) -> str:
    return "" if text is None else str(text).replace("\n", " ").strip()

def strip_non_ascii(s: str) -> str:
    return re.sub(r"[^\x00-\x7F]+", "", s)

def normalize_text(s: str) -> str:
    return unicodedata.normalize("NFKC", s)

def is_trivial_note(note: str) -> bool:
    if not note:
        return False
    n = normalize_text(note).strip()
    if _TRIVIAL_NOTE_REGEX.match(n):
        return True
    if n == "欠図":
        return True
    return False

def normalize_note(s: str) -> str:
    """Canonicalize note text for matching across variants, treating trivial as empty."""
    if not s or is_trivial_note(s):
        return ""
    s = normalize_text(s).strip()
    s = re.sub(r"\s+", " ", s)
    s = s.strip(" ;:-")
    return s.lower()

# ───────────── RED/STRIKE DELETION DETECTION ─────────────
def get_deleted_indices(page: pdfplumber.page.Page) -> Set[int]:
    """Find gutter indices marked red (line/text) and return set of ints (e.g., 68 from '68.1')."""
    deleted: Set[int] = set()

    for l in getattr(page, "lines", []):
        if l.get('stroking_color') == (1.0, 0.0, 0.0):
            s = page.crop((0, l['top'] - 13, 103, l['bottom'] + 10)).extract_text()
            if s:
                m = re.search(r"\d+", s)
                if m:
                    try:
                        deleted.add(int(m.group(0)))
                    except ValueError:
                        pass

    for c in getattr(page, "curves", []):
        if c.get('stroking_color') == (1.0, 0.0, 0.0):
            s = page.crop((0, c['top'] - 13, 103, c['bottom'] + 10)).extract_text()
            if s:
                m = re.search(r"\d+", s)
                if m:
                    try:
                        deleted.add(int(m.group(0)))
                    except ValueError:
                        pass

    for ch in getattr(page, "chars", []):
        if ch.get('non_stroking_color') == (1.0, 0.0, 0.0) and ch.get('x0', 9999) < 103:
            s = page.crop((0, ch['top'] - 13, 103, ch['bottom'] + 10)).extract_text()
            if s:
                m = re.search(r"\d+", s)
                if m:
                    try:
                        deleted.add(int(m.group(0)))
                    except ValueError:
                        pass

    return deleted

def parse_index_int(idx_text: str) -> Optional[int]:
    m = re.match(r"^(\d+)", idx_text.strip())
    if not m:
        return None
    try:
        return int(m.group(1))
    except ValueError:
        return None

# ───────────── PDF PARSE ─────────────
def parse_bom_pdf(pdf_path: Path) -> List[Dict]:
    entries: List[Dict] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            deleted_idxs = get_deleted_indices(page)

            tbls = page.extract_tables()
            if not tbls:
                continue
            tbl = tbls[0]

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

            # map columns
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
                elif "note" in txt or "備" in txt:  # 備考 NOTE
                    cmap['note'] = ci
            cmap.setdefault('part_number', 6)
            cmap.setdefault('part_name', 19)
            cmap.setdefault('quantity', None)
            cmap.setdefault('change', None)
            cmap.setdefault('note', None)

            # extract entries
            for row in tbl[hdr_i + 1:]:
                idx = clean_cell(row[0])
                if not idx or not re.match(r"^\d+(?:\.\d+)?$", idx):
                    continue

                # skip if flagged deleted
                idx_int = parse_index_int(idx)
                if idx_int is not None and idx_int in deleted_idxs:
                    continue

                # detect level (first non-empty among columns 1..5)
                level = next(
                    (ci for ci in range(1, 6) if ci < len(row) and clean_cell(row[ci])),
                    1
                )

                # raw part-name cell + early check for outline drawing
                raw = row[cmap['part_name']] if cmap['part_name'] < len(row) else ""
                raw_lines_all = [ln.strip() for ln in str(raw).splitlines() if ln.strip()]
                raw_text = " ".join(raw_lines_all).lower()
                if level == 1 and "outline drawing" in raw_text:
                    continue

                # NOTE (備考)
                note = ""
                if cmap['note'] is not None and cmap['note'] < len(row):
                    note = clean_cell(row[cmap['note']])
                if is_trivial_note(note):
                    note = ""

                # Build EN content:
                # Start from everything except the first line (often JP),
                # then optionally drop the first EN candidate if it’s JP/too-short.
                content = raw_lines_all[1:] if len(raw_lines_all) > 1 else raw_lines_all[:1]
                if content:
                    cand = content[0]
                    eng = re.sub(r'[^A-Za-z]', '', cand)
                    if JAPANESE_RE.search(cand) or len(eng) < CUT_THRESHOLD:
                        content = content[1:]

                # >>> Only flag multiline if we STILL have 2+ EN lines after filtering <<<
                had_multiline_english = len(content) > 1

                # flatten lines: default add space, join only if true mid-word split
                flat = ""
                for i, ln in enumerate(content):
                    ln = ln.strip()
                    if i == 0:
                        flat = ln
                    else:
                        if flat and flat[-1].isalpha() and ln and ln[0].isalpha() and not flat.endswith(" "):
                            flat += ln
                        else:
                            flat += " " + ln

                words = strip_non_ascii(flat).split()

                def fmt(w: str) -> str:
                    if any(ch.isdigit() for ch in w):
                        return w.upper()
                    if w.isupper() and len(w) <= CUT_THRESHOLD:
                        return w
                    return w.lower().capitalize()

                pname = " ".join(fmt(w) for w in words).strip()
                if not pname:
                    continue

                pn = normalize_text(clean_cell(row[cmap['part_number']])) if cmap['part_number'] < len(row) else ""
                dn = pn
                raw_chg = clean_cell(row[cmap['change']]) if (cmap['change'] is not None and cmap['change'] < len(row)) else ""
                nums = re.findall(r"\d+", raw_chg)
                chg = nums[-1] if nums else "0"

                qt = clean_cell(row[cmap['quantity']]) if (cmap['quantity'] is not None and cmap['quantity'] < len(row)) else ""
                try:
                    qty = int(qt) if qt else None
                except ValueError:
                    try:
                        qty = float(qt)
                    except ValueError:
                        qty = None

                note_used_as_display = bool(note)

                entries.append({
                    'level': level,
                    'part_name': pname,                      # canonical parsed name (for Part Type)
                    'display_name': (note or pname),         # visible "Part Name" cell
                    'note': note,
                    'note_norm': normalize_note(note),       # for merge logic
                    'part_number': pn,
                    'drawing_no': dn,
                    'change': chg,
                    'quantity': qty,
                    'flag_multiline_en': had_multiline_english,  # << ONLY EN multiline
                    'flag_note_used': note_used_as_display,      # << NOTE used
                })

    return entries

# ───────────── METADATA ─────────────
def parse_pdf_metadata(pdf_path: Path) -> Tuple[str, str]:
    with pdfplumber.open(str(pdf_path)) as pdf:
        full_text = "".join(page.extract_text() or "" for page in pdf.pages)
    lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]
    for i, line in enumerate(lines):
        up = line.upper()
        if up.startswith("PRODUCT NUMBER") and "CUSTOMER" in up:
            if i > 0:
                parts = lines[i-1].split()
                if len(parts) >= 2:
                    return strip_non_ascii(normalize_text(parts[1])), strip_non_ascii(normalize_text(parts[0]))
            break
    return "", ""

# ───────────── EXCEL WRITE ─────────────
def write_combined_excel(sheets: List[Tuple[str, List[Dict], str, str]], out_path: Path) -> None:
    parts: Dict[str, Dict] = {}
    initial_order: List[str] = []

    def make_key(r: Dict, vi: int, seq: int) -> str:
        if r['level'] == 1:
            nn = r.get('note_norm', "")
            return f"{r['part_number']}@@n={nn}"
        return f"{r['part_number']}@@v{vi}@@i{seq}"

    # build parts dict
    seq_counter = 0
    for vi, (_, ents, _, _) in enumerate(sheets):
        for r in ents:
            key = make_key(r, vi, seq_counter)
            seq_counter += 1

            if key not in parts:
                parts[key] = {**r, 'qtys': [None] * len(sheets)}

            existing = parts[key]['qtys'][vi]
            q_add = r['quantity']
            if q_add is not None:
                new_total = (existing or 0) + q_add
                parts[key]['qtys'][vi] = None if new_total == 0 else new_total

            if key not in initial_order:
                initial_order.append(key)

    # segment into level-1 clusters
    segments: List[List[str]] = []
    current: List[str] = []
    for key in initial_order:
        if parts[key]['level'] == 1:
            if current:
                segments.append(current)
            current = [key]
        else:
            current.append(key)
    if current:
        segments.append(current)

    # group segments by their level-1 canonical part_name (not display_name)
    seg_by_name: Dict[str, List[List[str]]] = {}
    name_order: List[str] = []
    for seg in segments:
        name = parts[seg[0]]['part_name']
        if name not in seg_by_name:
            seg_by_name[name] = []
            name_order.append(name)
        seg_by_name[name].append(seg)

    # final order
    ordered_keys: List[str] = []
    for name in name_order:
        for seg in seg_by_name[name]:
            ordered_keys.extend(seg)

    # column positions
    max_lvl = max(parts[k]['level'] for k in ordered_keys) if ordered_keys else 1
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
    grey_fmt = wb.add_format({'bg_color': '#A6A6A6', 'border': 0})

    # highlight formats
    hl_multiline_center = wb.add_format({**base, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFF2CC', 'pattern': 1})
    hl_multiline_left   = wb.add_format({**base, 'align': 'left',   'valign': 'vcenter', 'bg_color': '#FFF2CC', 'pattern': 1})
    hl_note_fmt         = wb.add_format({**base, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#F4CCCC', 'pattern': 1})
    hl_both_fmt         = wb.add_format({**base, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9D2E9', 'pattern': 1})

    # prefill canvas
    GRID_ROWS = 500
    GRID_COLS = 50
    for rr in range(GRID_ROWS):
        for cc in range(GRID_COLS):
            ws.write_blank(rr, cc, None, grey_fmt)

    # set row heights & core column widths
    for i, h in enumerate(BLANK_HEIGHTS):
        ws.set_row(i, h)
    lvl_ws = [SMALL_WIDTH] * (max_lvl - 1) + [TOTAL_WIDTH - SMALL_WIDTH * (max_lvl - 1)]
    for i, w in enumerate(lvl_ws):
        ws.set_column(i, i, w)
    ws.set_column(PN_COL, PN_COL, PART_NAME_W)
    for j, w in enumerate(BLANK_WIDTHS[3:6]):
        ws.set_column(GROUP_ST + j, GROUP_ST + j, w)

    # each variant quantity column width
    for vi in range(len(sheets)):
        ws.set_column(QTY_ST + vi, QTY_ST + vi, 13.91)

    # title block
    ws.merge_range(0, 0, START_ROW - 1, PN_COL, 'Spreadsheet of Variants', title_fmt)

    # merged/bordered boxes rows 1–9
    for r in range(0, 9):
        if len(sheets) > 1:
            ws.merge_range(r, QTY_ST, r, QTY_ST + len(sheets) - 1, '', merge_fmt)
        else:
            ws.write(r, QTY_ST, '', merge_fmt)

    # vertical boxes rows 10–14
    for r in range(9, START_ROW):
        for vi in range(len(sheets)):
            ws.write(r, QTY_ST + vi, "", vert_fmt)

    # labels rows 1–14
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

    # header row 15
    ws.merge_range(START_ROW, 0, START_ROW, PN_COL - 1, 'Part Type', merge_fmt)
    ws.write(START_ROW, PN_COL, 'Part Name', merge_fmt)
    ws.write(START_ROW, GROUP_ST, 'Part Number', merge_fmt)
    ws.write(START_ROW, GROUP_ST + 1, 'Drawing Number', merge_fmt)
    ws.write(START_ROW, GROUP_ST + 2, 'Rev.', merge_fmt)
    for vi in range(len(sheets)):
        ws.write(START_ROW, QTY_ST + vi, '', merge_fmt)

    # data rows
    for idx, key in enumerate(ordered_keys):
        r = START_ROW + 1 + idx
        rec = parts[key]
        sc = rec['level'] - 1
        if sc > 0:
            for c in range(sc):
                ws.write(r, c, '', blank_fmt)
        ec = PN_COL - 1

        # choose Part Type format (highlight if multiline-EN)
        pt_fmt = hl_multiline_left if rec.get('flag_multiline_en') and rec['level'] > 1 else (
                 hl_multiline_center if rec.get('flag_multiline_en') else
                 (merge_fmt if rec['level'] == 1 else left_fmt)
        )

        # Part Type area = canonical part_name (NOT the note)
        if sc < ec:
            ws.merge_range(r, sc, r, ec, rec['part_name'], pt_fmt)
        else:
            ws.write(r, sc, rec['part_name'], pt_fmt)

        # Part Name column = note-or-name + HIGHLIGHTS
        visible_name = rec.get('display_name') or rec['part_name']
        flag_multiline = bool(rec.get('flag_multiline_en'))
        flag_note_used = bool(rec.get('flag_note_used'))

        if flag_multiline and flag_note_used:
            pn_fmt = hl_both_fmt
        elif flag_note_used:
            pn_fmt = hl_note_fmt
        elif flag_multiline:
            pn_fmt = hl_multiline_center
        else:
            pn_fmt = merge_fmt

        ws.write(r, PN_COL, visible_name, pn_fmt)

        ws.write(r, GROUP_ST, rec['part_number'], merge_fmt)
        ws.write(r, GROUP_ST + 1, rec['part_number'], merge_fmt)
        ws.write(r, GROUP_ST + 2, f"D{rec['change']}", rev_fmt)

        # quantities: blank for None or 0
        for vi, q in enumerate(rec['qtys']):
            col = QTY_ST + vi
            if q is None or q == 0:
                ws.write(r, col, "", merge_fmt)
            else:
                try:
                    ws.write_number(r, col, float(q), merge_fmt)
                except Exception:
                    ws.write(r, col, str(q), merge_fmt)

    last_r = START_ROW + len(ordered_keys)
    for rr in range(START_ROW, last_r + 1):
        ws.set_row(rr, 32.3)

    writer.close()
    logging.info('Written Combined → %s', out_path)

# ───────────── CLI ─────────────
def main() -> None:
    parser = argparse.ArgumentParser(description='Generate Combined SOV from BOM PDFs')
    parser.add_argument('-s', '--sheet', action='append', nargs=2,
                        metavar=('VAR', 'PDF'), required=True,
                        help='Variant name + PDF path')
    parser.add_argument('-o', '--output', required=True, help='Output XLSX path')
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
