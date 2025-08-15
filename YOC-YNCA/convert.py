import re
import pdfplumber
import sqlite3
import os
import sys
import argparse

def clean_cell(text):
    """Normalize a cell: turn None → ""; strip newlines/whitespace."""
    if text is None:
        return ""
    return str(text).replace('\n', ' ').strip()

def strip_non_ascii(s: str) -> str:
    """Remove any non‑ASCII characters (e.g. Japanese) from the string."""
    return re.sub(r'[^\x00-\x7F]+', '', s)

def parse_bom(pdf_path, db_path):
    # --- 1) Sanity checks & reset ---
    if not os.path.exists(pdf_path):
        print(f"PDF not found: {pdf_path}")
        sys.exit(1)
    if os.path.exists(db_path):
        print(f"Removing old database '{db_path}'…")
        os.remove(db_path)

    all_entries = []
    print(f"Opening PDF '{pdf_path}'…")

    with pdfplumber.open(pdf_path) as pdf:
        # metadata from first page
        first_tbl = pdf.pages[0].extract_tables()[0]
        product_name         = clean_cell(first_tbl[0][7])
        product_number       = clean_cell(first_tbl[1][7])
        customer_part_number = clean_cell(first_tbl[1][21])
        print(f"  • Product Name:  {product_name}")
        print(f"  • Product No.:   {product_number}")
        print(f"  • Cust P/N:      {customer_part_number}")

        # iterate pages
        for pg, page in enumerate(pdf.pages, start=1):
            print(f"  • Page {pg}…")
            tables = page.extract_tables()
            if not tables:
                continue
            tbl = tables[0]

            # find header row
            header_idx = None
            for i, row in enumerate(tbl[:6]):
                low = [clean_cell(c).lower() for c in row]
                if any("level" in c for c in low) \
                and any("part number" in c for c in low) \
                and any("qty" in c for c in low):
                    header_idx = i
                    break
            if header_idx is None:
                header_idx = 2

            header = tbl[header_idx]
            # build column map (including “change”)
            col_map = {}
            for i, cell in enumerate(header):
                t = clean_cell(cell).lower()
                if "part number" in t:
                    col_map["part_number"] = i
                elif any(x in t for x in ("draw.no", "draw no", "draw.")) and "product number" not in t:
                    col_map.setdefault("draw_no", i)
                elif "part name" in t:
                    col_map["part_name"] = i
                elif "qty" in t:
                    col_map["quantity"] = i
                elif "note" in t:
                    col_map["note"] = i
                elif "change" in t:
                    col_map["change"] = i

            # fallbacks
            col_map.setdefault("part_number", 6)
            col_map.setdefault("draw_no",     17)
            col_map.setdefault("part_name",   19)
            col_map.setdefault("quantity",    23)
            col_map.setdefault("note",        24)
            col_map.setdefault("change",      None)

            # scan data rows
            for row in tbl[header_idx + 1 :]:
                item_no = clean_cell(row[0])
                if not item_no or not re.match(r'^\d+(\.\d+)?$', item_no):
                    continue

                # level from cols 1–5
                level = ""
                for lvl in range(1, 6):
                    if lvl < len(row):
                        v = clean_cell(row[lvl])
                        if v:
                            level = v
                            break

                # part_name (English only)
                raw_name = clean_cell(row[col_map["part_name"]]) if col_map["part_name"] < len(row) else ""
                eng_name = strip_non_ascii(raw_name)

                # change column
                raw_chg = ""
                if col_map["change"] is not None and col_map["change"] < len(row):
                    raw_chg = clean_cell(row[col_map["change"]])
                change_val = raw_chg or ""

                entry = {
                    "page":        pg,
                    "item_no":     item_no,
                    "level":       level,
                    "part_number": clean_cell(row[col_map["part_number"]]) if col_map["part_number"] < len(row) else "",
                    "draw_no":     clean_cell(row[col_map["draw_no"]])     if col_map["draw_no"]     < len(row) else "",
                    "part_name":   eng_name,
                    "quantity":    clean_cell(row[col_map["quantity"]])    if col_map["quantity"]    < len(row) else "",
                    "note":        clean_cell(row[col_map["note"]])        if col_map["note"]        < len(row) else "",
                    "change":      change_val,
                }
                all_entries.append(entry)

    if not all_entries:
        print("No SOV entries found.")
        return

    print(f"\nExtracted {len(all_entries)} entries.")

    # --- 4) Write to SQLite ---
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS bom_data (
        id                    INTEGER PRIMARY KEY AUTOINCREMENT,
        source_pdf            TEXT,
        product_name          TEXT,
        product_number        TEXT,
        customer_part_number  TEXT,
        page                  INTEGER,
        item_no               TEXT,
        level                 TEXT,
        part_number           TEXT,
        draw_no               TEXT,
        part_name             TEXT,
        quantity              TEXT,
        note                  TEXT,
        change                TEXT,
        UNIQUE(product_number, item_no, page)
    )
    """)
    rows = [
        (
            pdf_path,
            product_name,
            product_number,
            customer_part_number,
            e["page"],
            e["item_no"],
            e["level"],
            e["part_number"],
            e["draw_no"],
            e["part_name"],
            e["quantity"],
            e["note"],
            e["change"]
        )
        for e in all_entries
    ]
    cur.executemany("""
    INSERT OR IGNORE INTO bom_data (
        source_pdf, product_name, product_number, customer_part_number,
        page, item_no, level, part_number, draw_no, part_name,
        quantity, note, change
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, rows)
    conn.commit()
    conn.close()
    print(f"Done — {len(rows)} rows written to '{db_path}'.")

if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Parse BOM PDF into SQLite")
    p.add_argument("pdf_path", help="Path to the BOM PDF file")
    p.add_argument(
        "-d", "--db",
        default="yazaki_bom.db",
        help="SQLite database path (will be overwritten if exists)"
    )
    args = p.parse_args()
    parse_bom(args.pdf_path, args.db)
