#!/usr/bin/env python3
"""Parse OCR'd Hartill Chapter 22 (Qing) pages into a structured spreadsheet,
organised by Reign -> Mint -> Variety, with a sortable numeric Rarity column.

Source of truth: per-page OCR text in _pages/txt/pNNN.txt (produced by EasyOCR).
Output: CoinDatabase_Hartill_Qing.xlsx  (+ a CSV copy of the Coins sheet)

Design notes
------------
* OCR corrupts Chinese glyphs and even romanized names, so reign assignment is
  driven by the reliable reign DATE-RANGE in section headers, e.g. "(1662-1722)".
* Mint is extracted on a BEST-EFFORT basis from sub-headings in the OCR text:
    - "Place, Province" / "Place; Province" where Province is a known Qing province,
    - the two Peking central mints: "Board of Revenue" (Boo chiowan) and
      "Board of Works" (Boo yuwan), plus "Iron Coin Mint",
    - Xinjiang city mints (Aksu, Ili, Kashgar, ...).
  The detected mint is carried forward to following entries until the next heading.
* Rarity is Hartill's rarity figure, a trailing 1-2 digit number after the catalogue
  cross-references (FD/TD = Ding, S = Schjoth). IMPORTANT: on Hartill's scale a LOWER
  number is RARER (0 = rarest; common cash are ~15-20). Because the book is printed in
  two columns interleaved with coin images, the OCR sometimes detaches the rarity figure
  from its entry, so this column is partial -- verify against the book before relying on it.
* Each Hartill entry number (e.g. 22.140) anchors one catalog row.
"""
import os
import re
import glob
import csv

TXT_DIR = "_pages/txt"
OUT_XLSX = "CoinDatabase_Hartill_Qing.xlsx"
OUT_CSV = "CoinDatabase_Hartill_Qing.csv"

# ---------------------------------------------------------------------------
# Reigns
# ---------------------------------------------------------------------------
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
# Chronological order for sorting (reign era -> index)
REIGN_ORDER = {era: i for i, era in enumerate(
    ["Shunzhi", "Kangxi", "Yongzheng", "Qianlong", "Jiaqing",
     "Daoguang", "Xianfeng", "Tongzhi", "Guangxu", "Xuantong"]
)}

# A reign header is a date range whose start year is one of the known reign starts.
REIGN_HEADER_RE = re.compile(
    r"\((" + "|".join(REIGN_BY_START.keys()) + r")[\s.\-\u2013](\d{2,4})\)"
)

# ---------------------------------------------------------------------------
# Mints
# ---------------------------------------------------------------------------
# Known Qing provinces (romanisations + common OCR variants).  Used to spot
# "Place, Province" mint sub-headings.
PROVINCES = [
    "Zhili", "Chihli", "Fengtian", "Shengjing", "Shanxi", "Shaanxi", "Shensi",
    "Gansu", "Gansy", "Shandong", "Henan", "Jiangsu", "Jiangnan", "Jiangxi",
    "Anhui", "Anhwei", "Fujian", "Zhejiang", "Hubei", "Hunan", "Guangdong",
    "Guangxi", "Sichuan", "Szechuan", "Yunnan", "Guizhou", "Xinjiang",
    "Taiwan", "Jilin", "Heilongjiang",
]
PROV_ALT = "|".join(sorted(PROVINCES, key=len, reverse=True))
# "Ruzhou, Fujian"  /  "Kaifeng; Henan"  /  "Yanghe garrison, Shanxi"
MINT_PLACE_RE = re.compile(
    r"([A-Z][A-Za-z\u4e00-\u9fff'()]+(?:\s+(?:garrison|branch)?\s*)?)[,;]\s*(" + PROV_ALT + r")\b"
)

# Central / special mints recognised by phrase (order matters: most specific first).
SPECIAL_MINTS = [
    (re.compile(r"Iron Coin Mint", re.I), "Iron Coin Mint (Peking)"),
    (re.compile(r"Board of Revenue|Boo[\s-]*chiowan", re.I), "Board of Revenue (Boo-chiowan, Peking)"),
    (re.compile(r"Board of Works|Boo[\s-]*yuwan", re.I), "Board of Works (Boo-yuwan, Peking)"),
]
# Xinjiang / frontier city mints (often appear as "Aksu, Xinjiang" too, but also alone).
CITY_MINTS = [
    "Aksu", "Ili", "Kashgar", "Yarkand", "Kucha", "Kuche", "Ushi", "Khotan",
    "Urumqi", "Kuldja", "Tarbagatai",
]
CITY_MINT_RE = re.compile(r"\b(" + "|".join(CITY_MINTS) + r")\b")

# Normalise common OCR garbles in mint place names so identical mints group together.
MINT_FIXES = {
    "Gansy": "Gansu", "Anhwei": "Anhui", "Shensi": "Shaanxi", "Szechuan": "Sichuan",
    "Chihli": "Zhili", "Guangzbou": "Guangzhou", "Ki'an": "Xi'an", "Kian": "Xi'an",
    "Xian": "Xi'an", "Jinnn": "Jinan", "Jinnn": "Jinan", "Puzhou": "Fuzhou",
    "Dihua)": "Dihua", "Dihua": "Dihua", "Kuche": "Kucha",
}
GARRISON_WORDS = {"garrison", "barrison", "carrison", "branch", "mint"}


def normalise_place(place):
    place = re.sub(r"\s+", " ", place).strip(" .,;:()")
    # strip leading OCR fragments of reverse-legend words ("right"/"left"/"above")
    place = re.sub(r"^(?:Tight|Right|Left|Above|Below|Rev|Obv)\s*", "", place)
    # strip trailing 'garrison'/'branch' descriptors
    place = re.sub(r"\s+(garrison|barrison|branch|mint)\b.*$", "", place, flags=re.I).strip(" .,;:()")
    if place.lower() in GARRISON_WORDS:
        return ""    # valid province, unnamed garrison
    if len(place) < 3:
        return None  # too garbled -> suppress (keep carried-forward mint)
    return MINT_FIXES.get(place, place)


def make_mint(place, prov):
    p = normalise_place(place)
    if p is None:
        return None
    if p == "":
        return f"{prov} (unnamed garrison)"
    return f"{p}, {prov}"

# ---------------------------------------------------------------------------
# Entries / inscriptions / dates
# ---------------------------------------------------------------------------
# Catalog entry token: 22.NNNN or 22,NNNN possibly with stray punctuation/spaces.
ENTRY_RE = re.compile(r"\b22[.,]\s?(\d{1,4})\b")

# Obverse inscription pinyin like "Shun Zhi tong bao" (allow OCR noise in names).
OBV_RE = re.compile(r"Obv[:;]\s*(.{0,40}?tong\s*bao)", re.IGNORECASE)

# A date range inside an entry, e.g. 1667-70, 1878-83, 1909-11
ENTRY_DATE_RE = re.compile(r"\b(1[6789]\d{2})[.\-](\d{2,4})\b")

# Catalogue cross-reference (Ding = FD/TD/D, Schjoth = S) followed by the rarity figure.
# Hartill prints rarity as a small trailing integer; we look for it after a ref number.
REF_THEN_RARITY_RE = re.compile(r"(?:[FTDS尺卫][\sD]?\d{3,4}[,，\s]*){1,3}(\d{1,2})\b")

BOOK_PAGE_BASE = 281      # Book page = 281 + pdf page index (verified).
MAX_SUBNO = 1521          # highest legitimate Qing entry (22.1521, Puyi/Xuantong)


def clean(s):
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip(" .,;:\u3002")


def detect_mint(text):
    """Return a normalised mint name if `text` contains a mint sub-heading, else None."""
    for rx, name in SPECIAL_MINTS:
        if rx.search(text):
            return name
    m = MINT_PLACE_RE.search(text)
    if m:
        return make_mint(m.group(1), m.group(2))
    m = CITY_MINT_RE.search(text)
    if m:
        return f"{MINT_FIXES.get(m.group(1), m.group(1))}, Xinjiang"
    return None


# A trailing 1-2 digit rarity that follows a typical entry-ending token
# (date, catalogue ref, period, or descriptive word). Guards against grabbing
# fragments of years / measurements.
TRAILING_RARITY_RE = re.compile(
    r"(?:\d{2,4}|bao|tong|right|left|above|below|head|qian|style|\))[\s.,\u3002]+(\d{1,2})\s*$",
    re.I,
)


def extract_rarity(desc):
    """Best-effort Hartill rarity figure (lower = rarer). Returns int or ''."""
    matches = REF_THEN_RARITY_RE.findall(desc)
    if matches:
        val = int(matches[-1])
        if 0 <= val <= 25:
            return val
    m = TRAILING_RARITY_RE.search(desc)
    if m:
        val = int(m.group(1))
        if 0 <= val <= 25:
            return val
    return ""


def parse():
    rows = []
    # Chapter 22 always opens with Shunzhi; default so caption tokens before the
    # first "Emperor" header on page 1 aren't left unassigned.
    current = {"era": "Shunzhi", "reign_dates": "1644-1661", "temple": "Shi Zu", "obv": ""}
    current_mint = ""

    files = sorted(glob.glob(os.path.join(TXT_DIR, "p*.txt")))
    pages = []  # (pdf_index, book_page, text)
    for f in files:
        idx = int(re.search(r"p(\d+)\.txt", f).group(1))
        with open(f) as fh:
            text = fh.read()
        pages.append((idx, str(BOOK_PAGE_BASE + idx), text))

    for pdf_idx, book_page, text in pages:
        # Build ordered event list: reign headers, mint headings, entry tokens.
        events = []
        for m in REIGN_HEADER_RE.finditer(text):
            events.append(("reign", m.start(), m.end(), m))
        for m in ENTRY_RE.finditer(text):
            events.append(("entry", m.start(), m.end(), m))
        # Mint headings: scan for province/special/city matches.
        for rx, name in SPECIAL_MINTS:
            for m in rx.finditer(text):
                events.append(("mint", m.start(), m.end(), name))
        for m in MINT_PLACE_RE.finditer(text):
            mint = make_mint(m.group(1), m.group(2))
            if mint:
                events.append(("mint", m.start(), m.end(), mint))
        for m in CITY_MINT_RE.finditer(text):
            city = MINT_FIXES.get(m.group(1), m.group(1))
            events.append(("mint", m.start(), m.end(), f"{city}, Xinjiang"))

        events.sort(key=lambda e: e[1])

        prev_end = 0
        for i, (kind, start, end, payload) in enumerate(events):
            next_start = events[i + 1][1] if i + 1 < len(events) else len(text)
            if kind == "reign":
                sy = payload.group(1)
                # Skip the chapter running-title "The Qing Dynasty (1644-1912)",
                # which appears atop every page and is NOT a Shunzhi reign header.
                if sy == "1644" and payload.group(2) == "1912":
                    continue
                era, dates, temple = REIGN_BY_START[sy]
                current = {"era": era, "reign_dates": dates, "temple": temple, "obv": ""}
                ctx = text[end:next_start]
                ob = OBV_RE.search(ctx)
                if ob:
                    current["obv"] = clean(ob.group(1))
                prev_end = end
                continue
            if kind == "mint":
                current_mint = payload
                prev_end = end
                continue
            # entry
            subno = payload.group(1)
            catno = f"22.{subno}"
            desc = clean(text[end:next_start])
            preceding = clean(text[prev_end:start])
            dm = ENTRY_DATE_RE.search(desc)
            entry_date = f"{dm.group(1)}-{dm.group(2)}" if dm else ""
            rarity = extract_rarity(desc)
            rows.append({
                "Catalog No": catno,
                "_sub": int(subno),
                "Reign Era": current["era"],
                "Reign Dates": current["reign_dates"],
                "Temple/Personal Name": current["temple"],
                "Mint": current_mint,
                "Obverse Inscription": current["obv"],
                "Description": desc,
                "Rarity (Hartill: lower=rarer)": rarity,
                "Entry Date": entry_date,
                "Book Page": book_page,
                "PDF Page": pdf_idx + 1,
                "_desclen": len(desc),
                "_rarity_num": rarity if rarity != "" else 999,
            })
            prev_end = end

    # Dedup: caption-column numbers create near-empty duplicates. Keep, per catalog
    # number, the row with the longest description (and prefer one that has a mint/rarity).
    best = {}
    for r in rows:
        if r["_sub"] > MAX_SUBNO:
            continue
        k = r["Catalog No"]
        if k not in best:
            best[k] = r
            continue
        cur = best[k]
        # Prefer richer rows: more description, then having a mint, then having rarity.
        score_new = (r["_desclen"], bool(r["Mint"]), r["Rarity (Hartill: lower=rarer)"] != "")
        score_cur = (cur["_desclen"], bool(cur["Mint"]), cur["Rarity (Hartill: lower=rarer)"] != "")
        if score_new > score_cur:
            best[k] = r
    out = list(best.values())

    # Sort by Reign (chronological) -> Mint -> catalog sub-number.
    out.sort(key=lambda r: (
        REIGN_ORDER.get(r["Reign Era"], 99),
        r["Mint"] or "\uffff",   # blank mints sort last within a reign
        r["_sub"],
    ))
    return out, pages


def write_outputs(rows, pages):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter

    cols = [
        "Catalog No", "Reign Era", "Reign Dates", "Temple/Personal Name",
        "Mint", "Obverse Inscription", "Description",
        "Rarity (Hartill: lower=rarer)", "Entry Date", "Book Page", "PDF Page",
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Coins"
    ws.append(cols)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for c in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    for r in rows:
        ws.append([r.get(c, "") for c in cols])

    widths = [10, 12, 12, 18, 30, 22, 60, 14, 12, 10, 9]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    # AutoFilter over the whole table so the user can filter Reign/Mint and sort Rarity.
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"

    # ---- Notes / methodology sheet ----
    ws3 = wb.create_sheet("Notes")
    notes = [
        ("Hartill Qing (Chapter 22) coin database - methodology & caveats", ""),
        ("", ""),
        ("Organisation", "Rows are sorted by Reign (chronological) -> Mint -> Hartill catalogue number."),
        ("Filtering", "Row 1 of the Coins sheet has AutoFilter. Click the dropdowns to filter by Reign Era or Mint."),
        ("Sorting rarity", "Use the AutoFilter dropdown on the Rarity column to sort. LOWER number = RARER."),
        ("Rarity scale", "Hartill's rarity figure: 0 = rarest; common Qing cash are typically ~15-20."),
        ("Mint extraction", "Mint is read from OCR sub-headings (Place, Province / Board of Revenue / Board of Works / Iron Coin Mint / Xinjiang cities) and carried forward to following entries. Best-effort only."),
        ("Rarity extraction", "Rarity is the trailing figure after the Ding (FD/TD) and Schjoth (S) cross-references. The book's two-column layout with coin images sometimes detaches this figure during OCR, so the column is PARTIAL. Verify against the book before relying on it."),
        ("Source text", "EasyOCR of 'Hartill Qing Coins Only.pdf' (book pages 281-425). Chinese glyphs and personal/mint names are often garbled by OCR."),
        ("Book page", "Book Page = 281 + PDF page index (verified)."),
        ("Verification", "The 'Raw OCR by Page' sheet preserves full OCR text per page for cross-checking."),
    ]
    for r in notes:
        ws3.append(list(r))
    ws3.cell(row=1, column=1).font = Font(bold=True, size=12)
    ws3.column_dimensions["A"].width = 22
    ws3.column_dimensions["B"].width = 110
    for row in ws3.iter_rows(min_row=1):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # ---- Raw OCR sheet (traceability) ----
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
    from collections import Counter
    by_era = Counter(r["Reign Era"] or "(unassigned)" for r in rows)
    with_mint = sum(1 for r in rows if r["Mint"])
    with_rar = sum(1 for r in rows if r["Rarity (Hartill: lower=rarer)"] != "")
    print(f"Total catalog entries: {len(rows)}")
    print("By reign era:")
    for era in ["Shunzhi", "Kangxi", "Yongzheng", "Qianlong", "Jiaqing", "Daoguang",
                "Xianfeng", "Tongzhi", "Guangxu", "Xuantong", "(unassigned)"]:
        if by_era.get(era):
            print(f"  {era:12s} {by_era[era]}")
    print(f"Rows with a Mint assigned : {with_mint} / {len(rows)}")
    print(f"Rows with a Rarity figure : {with_rar} / {len(rows)}")
    print(f"Wrote {OUT_XLSX} and {OUT_CSV}")
