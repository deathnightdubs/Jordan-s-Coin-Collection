#!/usr/bin/env python3
"""Parse OCR'd Hartill Chapter 22 (Qing) pages into a structured spreadsheet.

Source of truth: per-page OCR text in _pages/txt/pNNN.txt (produced by EasyOCR).
Output: CoinDatabase_Hartill_Qing.xlsx  (+ a CSV copy of the Coins sheet)

Design notes:
- OCR corrupts Chinese glyphs and even romanized names, so reign assignment is
  driven by the reliable reign DATE-RANGE in section headers, e.g. "(1662-1722)".
- Each Hartill entry number (e.g. 22.140) anchors one catalog row. Description text
  is the OCR text between this number and the next number token.
- A separate "Raw OCR by Page" sheet preserves the full text for verification.
"""
import os
import re
import glob
import csv

TXT_DIR = "_pages/txt"
OUT_XLSX = "CoinDatabase_Hartill_Qing.xlsx"
OUT_CSV = "CoinDatabase_Hartill_Qing.csv"

# Map reign START YEAR -> (Era name, full reign-date string, temple/personal name)
REIGN_BY_START = {
    "1644": ("Shunzhi", "1644-1661", "Shi Zu"),
    "1662": ("Kangxi", "1662-1722", "Sheng Zu"),
    "1723": ("Yongzheng", "1723-1735", "Shi Zong"),
    "1736": ("Qianlong", "1736-1795", "Gao Zong"),
    "1796": ("Jiaqing", "1796-1820", "Ren Zong"),
    "1821": ("Daoguang", "1821-1850", "Xuan Zong"),
    "1851": ("Xianfeng", "1851-1861", "Wen Zong"),
    "1862": ("Tongzhi", "1862-1874", "Mu Zong"),
    "1875": ("Guangxu", "1875-1908", "De Zong"),
    "1909": ("Xuantong", "1909-1912", "Puyi"),
}

# A reign header is a date range whose start year is one of the known reign starts.
REIGN_HEADER_RE = re.compile(
    r"\((" + "|".join(REIGN_BY_START.keys()) + r")[\s.\-–](\d{2,4})\)"
)

# Catalog entry token: 22.NNNN or 22,NNNN possibly with stray punctuation/spaces.
ENTRY_RE = re.compile(r"\b22[.,]\s?(\d{1,4})\b")

# Obverse inscription pinyin like "Shun Zhi tong bao" (allow OCR noise in names).
OBV_RE = re.compile(r"Obv[:;]\s*(.{0,40}?tong\s*bao)", re.IGNORECASE)

# A date range inside an entry, e.g. 1667-70, 1878-83, 1909-11
ENTRY_DATE_RE = re.compile(r"\b(1[6789]\d{2})[.\-](\d{2,4})\b")


def clean(s):
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip(" .,;:。")


# Book page = 281 + pdf page index (verified: pdf 0 -> p.281 ... pdf 144 -> p.425).
# This is far more reliable than OCR of the printed page number.
BOOK_PAGE_BASE = 281
MAX_SUBNO = 1521  # highest legitimate Qing entry (22.1521, Puyi/Xuantong)


def parse():
    rows = []  # each: dict
    current = {"era": "", "reign_dates": "", "temple": "", "obv": ""}

    files = sorted(glob.glob(os.path.join(TXT_DIR, "p*.txt")))
    pages = []  # (pdf_index, book_page, text)

    for f in files:
        idx = int(re.search(r"p(\d+)\.txt", f).group(1))
        with open(f) as fh:
            text = fh.read()
        pages.append((idx, str(BOOK_PAGE_BASE + idx), text))

    for pdf_idx, book_page, text in pages:
        # Build ordered event list of reign headers + entry tokens.
        events = []
        for m in REIGN_HEADER_RE.finditer(text):
            events.append(("reign", m.start(), m.end(), m))
        for m in ENTRY_RE.finditer(text):
            events.append(("entry", m.start(), m.end(), m))
        events.sort(key=lambda e: e[1])

        prev_end = 0
        for i, (kind, start, end, m) in enumerate(events):
            next_start = events[i + 1][1] if i + 1 < len(events) else len(text)
            if kind == "reign":
                sy = m.group(1)
                era, dates, temple = REIGN_BY_START[sy]
                current = {"era": era, "reign_dates": dates, "temple": temple, "obv": ""}
                ctx = text[end:next_start]
                ob = OBV_RE.search(ctx)
                if ob:
                    current["obv"] = clean(ob.group(1))
                prev_end = end
                continue
            # entry
            subno = m.group(1)
            catno = f"22.{subno}"
            desc = clean(text[end:next_start])
            preceding = clean(text[prev_end:start])
            # Entry-level date if present
            dm = ENTRY_DATE_RE.search(desc)
            entry_date = f"{dm.group(1)}-{dm.group(2)}" if dm else ""
            rows.append({
                "Catalog No": catno,
                "_sub": int(subno),
                "Reign Era": current["era"],
                "Reign Dates": current["reign_dates"],
                "Temple/Personal Name": current["temple"],
                "Obverse Inscription": current["obv"],
                "Preceding Context (mint?)": preceding,
                "Description": desc,
                "Entry Date": entry_date,
                "Book Page": book_page,
                "PDF Page": pdf_idx + 1,
                "_desclen": len(desc),
            })
            prev_end = end

    # Dedup: caption-column numbers create near-empty duplicates. Keep, per catalog
    # number, the row with the longest description.
    best = {}
    for r in rows:
        if r["_sub"] > MAX_SUBNO:
            continue  # drop OCR-noise numbers above the real max (22.1521)
        k = r["Catalog No"]
        if k not in best or r["_desclen"] > best[k]["_desclen"]:
            best[k] = r
    out = list(best.values())
    out.sort(key=lambda r: r["_sub"])
    return out, pages


def write_outputs(rows, pages):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter

    cols = [
        "Catalog No", "Reign Era", "Reign Dates", "Temple/Personal Name",
        "Obverse Inscription", "Preceding Context (mint?)", "Description",
        "Entry Date", "Book Page", "PDF Page",
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Coins"
    ws.append(cols)
    for c in range(1, len(cols) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)
    for r in rows:
        ws.append([r.get(c, "") for c in cols])
    widths = [10, 12, 12, 18, 22, 34, 70, 12, 10, 9]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # Raw OCR sheet for verification/traceability
    ws2 = wb.create_sheet("Raw OCR by Page")
    ws2.append(["PDF Page", "Book Page", "Full OCR Text"])
    for c in range(1, 4):
        ws2.cell(row=1, column=c).font = Font(bold=True)
    for pdf_idx, book_page, text in pages:
        ws2.append([pdf_idx + 1, book_page, text])
    ws2.column_dimensions["A"].width = 9
    ws2.column_dimensions["B"].width = 10
    ws2.column_dimensions["C"].width = 120
    ws2.freeze_panes = "A2"

    wb.save(OUT_XLSX)

    with open(OUT_CSV, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow(r)


if __name__ == "__main__":
    rows, pages = parse()
    write_outputs(rows, pages)
    # Summary
    from collections import Counter
    by_era = Counter(r["Reign Era"] or "(unassigned)" for r in rows)
    print(f"Total catalog entries: {len(rows)}")
    print("By reign era:")
    for era in ["Shunzhi","Kangxi","Yongzheng","Qianlong","Jiaqing","Daoguang",
                "Xianfeng","Tongzhi","Guangxu","Xuantong","(unassigned)"]:
        if by_era.get(era):
            print(f"  {era:12s} {by_era[era]}")
    print(f"Wrote {OUT_XLSX} and {OUT_CSV}")
