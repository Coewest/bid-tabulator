#!/usr/bin/env python3
"""Bid Tabulator MVP - Flask backend.

POST /api/upload  (multipart, field 'files', up to 10 files: pdf/xlsx/xls/csv)
  -> { project_name, mode, bidders:[{name, location, file}],
       items:[{description, qty, unit, values:[{amount, rate, status}...]}],
       submitted_totals:[...], subtotals:[...], taxes:[...],
       warnings:[...], bidder_confidence:'high'|'review' }

GET  /            -> index.html
GET  /api/health  -> {ok:true}
"""
import csv
import io
import json
import os
import re
import tempfile

from flask import Flask, request, jsonify, send_from_directory

app = Flask(__name__)
BASE = os.path.dirname(os.path.abspath(__file__))
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

# ---------------------------------------------------------------- parsing utils

def clean_num_text(s):
    """'1 ,456.00' -> '1456.00', '9 1.00' -> '91.00', '1 .00' -> '1.00'"""
    t = s.replace("$", "").strip()
    # drop spaces between digits, around commas, and around decimal points
    t = re.sub(r"(?<=\d)\s+(?=[\d.])", "", t)
    t = re.sub(r"(?<=\.)\s+(?=\d)", "", t)
    t = re.sub(r"\s*,\s*", ",", t)
    t = t.replace(",", "")
    return t.strip()


def parse_money(raw):
    """Returns (value_or_None, status). status: priced | not_priced | included | text"""
    if raw is None:
        return None, "not_priced"
    s = str(raw).strip()
    if not s or s in ("-", "--", "---", "\u2014", "N/A", "n/a", "NA", "n.a.", "*", "x", "X"):
        return None, "not_priced"
    low = s.lower()
    if low in ("included", "incl", "incl.", "nic", "no charge", "free", "inc."):
        return 0.0, "included"
    neg = False
    t = s.strip()
    if t.startswith("(") and t.endswith(")"):
        neg, t = True, t[1:-1]
    t = clean_num_text(t)
    if not re.fullmatch(r"-?\d+(\.\d+)?", t):
        return None, "text"
    try:
        v = float(t)
        return (-v if neg else v), "priced"
    except ValueError:
        return None, "text"


def norm_desc(s):
    s = (s or "").lower()
    s = re.sub(r"\(note \d+\)", "", s)
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def desc_match(a, b, thresh=0.55):
    ta, tb = set(norm_desc(a).split()), set(norm_desc(b).split())
    if not ta or not tb:
        return False
    return len(ta & tb) / max(len(ta), len(tb)) >= thresh


HEADER_WORDS = ("item", "description", "scope", "work description", "bid item")
TOTAL_WORDS = ("total", "grand total", "bid total", "total bid", "total amount")
SUBTOTAL_WORDS = ("subtotal", "sub total")
TAX_WORDS = ("tax", "sales tax")

COMPANY_HINTS = ("llc", "inc", "corp", "co.", "construction", "electric", "builders",
                 "contracting", "company", "services", "enterprises", "associates")


def is_header_row(row):
    cells = [norm_desc(c) for c in row if c and str(c).strip()]
    joined = " ".join(cells)
    return any(w in joined for w in HEADER_WORDS)


def row_label(row):
    return norm_desc(row[0]) if row and row[0] else ""


def is_total_row(row):
    lab = row_label(row)
    return any(w == lab or lab.startswith(w + " ") for w in TOTAL_WORDS) or lab in TOTAL_WORDS


def is_subtotal_row(row):
    return row_label(row) in SUBTOTAL_WORDS


def is_tax_row(row):
    return any(w in row_label(row) for w in TAX_WORDS)


# ---------------------------------------------------------------- PDF (multi-bidder tab mode)

def _pdf_tables_with_boxes(path):
    import pdfplumber
    out = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for t in page.find_tables():
                data = t.extract()
                if not data:
                    continue
                # cell x-centers per column from Column bboxes
                col_x = []
                try:
                    for col in t.columns:
                        b = col.bbox
                        col_x.append((b[0] + b[2]) / 2)
                except (AttributeError, TypeError, IndexError):
                    pass
                out.append({"data": data, "col_x": col_x,
                            "bbox": t.bbox, "page_width": page.width})
    return out


def _name_xcenters_from_words(path):
    """Fallback: company-ish words with x positions (for name row detection)."""
    import pdfplumber
    out = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for w in page.extract_words():
                txt = w["text"]
                if any(h in txt.lower() for h in COMPANY_HINTS) and len(txt) > 2:
                    out.append((txt, (w["x0"] + w["x1"]) / 2, w["top"]))
    return out


def extract_pdf_tab(path, filename):
    """Parse a bid-tabulation PDF: one table, multiple bidder column groups."""
    tables = _pdf_tables_with_boxes(path)
    if not tables:
        return None, ["no tables found in PDF"]
    # largest table by cells
    tab = max(tables, key=lambda t: len(t["data"]) * max((len(r) for r in t["data"]), default=0))
    data, col_x = tab["data"], tab["col_x"]
    warnings = []

    # find header row
    hdr_idx = next((i for i, r in enumerate(data) if is_header_row(r)), None)
    if hdr_idx is None:
        return None, ["could not find ITEM header row"]
    header = data[hdr_idx]

    # bidder groups: columns after item/qty/unit (first 3 cols), detect RATE/TOTAL pairs
    # find first data column: first col whose header is not item/qty/unit-ish
    first_bid_col = 3
    for i, h in enumerate(header):
        hn = norm_desc(h)
        if hn in ("rate", "unit price", "price", "total", "amount", "extended", "bid"):
            first_bid_col = i
            break
    n_cols = len(header)
    # group remaining cols into pairs (RATE,TOTAL) or singles
    groups = []  # list of (total_col_idx, rate_col_idx_or_None)
    i = first_bid_col
    while i < n_cols:
        hn = norm_desc(header[i]) if i < len(header) else ""
        hn2 = norm_desc(header[i + 1]) if i + 1 < n_cols else ""
        if "rate" in hn and "total" in hn2:
            groups.append((i + 1, i))
            i += 2
        elif "total" in hn or "amount" in hn or "bid" in hn:
            groups.append((i, None))
            i += 1
        else:
            # unknown column: treat as a price column if any numeric below
            col_vals = [r[i] if i < len(r) else "" for r in data[hdr_idx + 1:hdr_idx + 6]]
            if any(parse_money(v)[1] == "priced" for v in col_vals):
                groups.append((i, None))
            i += 1
    if not groups:
        return None, ["no bidder price columns detected"]

    # bidder names: row(s) above header — pick the row with the most company-hint cells
    name_row_idx = None
    best_score = -1
    for ri in range(max(0, hdr_idx - 3), hdr_idx):
        row = data[ri]
        score = 0
        for ci in range(first_bid_col, len(row)):
            cell = (row[ci] or "")
            if any(h in cell.lower() for h in COMPANY_HINTS) and len(cell.strip()) > 3:
                score += 1
        if score > best_score:
            best_score, name_row_idx = score, ri

    n_bidders = len(groups)
    bidder_names = [""] * n_bidders
    bidder_locs = [""] * n_bidders
    name_conf = "low"
    if name_row_idx is not None and best_score > 0:
        name_row = data[name_row_idx]
        # map each group's x-center to nearest name cell x-center
        name_cells = []  # (text, x_center)
        for ci in range(first_bid_col, len(name_row)):
            cell = (name_row[ci] or "").strip()
            if cell and any(h in cell.lower() for h in COMPANY_HINTS):
                xc = col_x[ci] if ci < len(col_x) else None
                if xc is not None:
                    name_cells.append((cell, xc))
        group_x = []
        for (tc, rc) in groups:
            xs = [col_x[c] for c in (tc, rc) if c is not None and c < len(col_x)]
            group_x.append(sum(xs) / len(xs) if xs else None)
        used = set()
        for gi, gx in enumerate(group_x):
            if gx is None:
                continue
            best, bestd = None, 1e18
            for ni, (txt, nx) in enumerate(name_cells):
                if ni in used:
                    continue
                d = abs(nx - gx)
                if d < bestd:
                    best, bestd = ni, d
            if best is not None:
                used.add(best)
                txt = name_cells[best][0]
                parts = [p.strip() for p in txt.split("\n") if p.strip()]
                bidder_names[gi] = parts[0] if parts else txt
                bidder_locs[gi] = parts[1] if len(parts) > 1 else ""
        if all(bidder_names):
            name_conf = "high"
        elif any(bidder_names):
            name_conf = "medium"
            warnings.append("some bidder names could not be mapped to columns - please confirm")
        else:
            warnings.append("bidder names not found - please enter them manually")
    else:
        warnings.append("no bidder name row detected - please enter bidder names manually")

    for gi in range(n_bidders):
        if not bidder_names[gi]:
            bidder_names[gi] = f"Bidder {gi + 1}"

    # item rows
    items = []
    submitted, subtotals, taxes = [None] * n_bidders, [None] * n_bidders, [None] * n_bidders
    for r in data[hdr_idx + 1:]:
        if not r or not any((c or "").strip() for c in r):
            continue
        if is_total_row(r):
            for gi, (tc, _) in enumerate(groups):
                v, st = parse_money(r[tc] if tc < len(r) else "")
                if st == "priced":
                    submitted[gi] = v
            continue
        if is_subtotal_row(r):
            for gi, (tc, _) in enumerate(groups):
                v, st = parse_money(r[tc] if tc < len(r) else "")
                if st == "priced":
                    subtotals[gi] = v
            continue
        if is_tax_row(r):
            for gi, (tc, _) in enumerate(groups):
                v, st = parse_money(r[tc] if tc < len(r) else "")
                if st == "priced":
                    taxes[gi] = v
            continue
        desc = (r[0] or "").strip()
        if not desc or len(desc) < 2:
            continue
        qty_raw = (r[1] or "").strip() if len(r) > 1 else ""
        unit_raw = (r[2] or "").strip() if len(r) > 2 else ""
        try:
            qty = float(clean_num_text(qty_raw)) if qty_raw else None
        except ValueError:
            qty = None
        vals = []
        for gi, (tc, rc) in enumerate(groups):
            amt_raw = r[tc] if tc < len(r) else ""
            rate_raw = r[rc] if rc is not None and rc < len(r) else ""
            amt, st = parse_money(amt_raw)
            rate, _ = parse_money(rate_raw)
            vals.append({"amount": amt, "rate": rate, "status": st,
                         "raw": str(amt_raw or "").strip()})
        # skip rows with no priced data at all (section headers etc.)
        if not any(v["status"] in ("priced", "included") for v in vals):
            # keep rows where at least one bidder has not_priced but others priced
            if all(v["status"] == "not_priced" or v["status"] == "text" for v in vals):
                continue
        items.append({"description": desc, "qty": qty, "unit": unit_raw, "values": vals})

    # math verification warnings
    for gi in range(n_bidders):
        calc = sum(v["amount"] for it in items for v in [it["values"][gi]]
                   if v["status"] in ("priced", "included") and v["amount"] is not None)
        base = subtotals[gi] if subtotals[gi] is not None else submitted[gi]
        if base is not None and base != 0:
            if abs(calc - base) > max(1.0, abs(base) * 0.002):
                warnings.append(
                    f"{bidder_names[gi]}: line items sum to "
                    f"${calc:,.2f} but {'subtotal' if subtotals[gi] is not None else 'total'} "
                    f"shows ${base:,.2f} - possible math error in source")

    return {
        "project_name": os.path.splitext(filename)[0].replace("_", " ").replace("-", " "),
        "mode": "tab",
        "bidders": [{"name": n, "location": l, "file": filename}
                    for n, l in zip(bidder_names, bidder_locs)],
        "items": items,
        "submitted_totals": submitted,
        "subtotals": subtotals,
        "taxes": taxes,
        "warnings": warnings,
        "bidder_confidence": name_conf,
    }, []


# ---------------------------------------------------------------- single-bid files (excel / csv / pdf-text)

def _rows_to_bidder(rows, filename):
    """rows: list of lists (strings). Find item/amount table -> one bidder's items."""
    # find header
    hdr = next((i for i, r in enumerate(rows) if is_header_row(r)), None)
    start = hdr + 1 if hdr is not None else 0
    header = rows[hdr] if hdr is not None else []
    # amount column: prefer TOTAL/AMOUNT, else last numeric-ish column
    amt_col = None
    if header:
        for i, h in enumerate(header):
            if norm_desc(h) in ("total", "amount", "extended", "extended total", "bid", "price"):
                amt_col = i
                break
    items = []
    for r in rows[start:]:
        if not r or not any(str(c or "").strip() for c in r):
            continue
        if is_total_row(r) or is_subtotal_row(r) or is_tax_row(r):
            continue
        desc = str(r[0] or "").strip()
        if not desc or len(desc) < 2:
            continue
        # amount: chosen col or scan for last money-like cell
        cands = []
        order = [amt_col] if amt_col is not None else []
        order += [i for i in range(len(r) - 1, -1, -1) if i != amt_col]
        val, status, rate = None, "not_priced", None
        for i in order:
            if i is None or i >= len(r):
                continue
            v, st = parse_money(r[i])
            if st in ("priced", "included", "not_priced"):
                val, status = v, st
                break
        if status == "text":
            continue
        items.append({"description": desc, "qty": None, "unit": "",
                      "raw_amount": val, "raw_status": status})
    # bidder name from filename
    name = os.path.splitext(filename)[0].replace("_", " ").replace("-", " ")
    return {"name": name, "location": "", "file": filename,
            "items": items}


def extract_excel(path, filename):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    # pick sheet with most rows
    best = None
    for ws in wb.worksheets:
        rows = [[str(c.value).strip() if c.value is not None else "" for c in row]
                for row in ws.iter_rows()]
        rows = [r for r in rows if any(r)]
        if best is None or len(rows) > len(best):
            best = rows
    if not best:
        return None, ["empty workbook"]
    return _rows_to_bidder(best, filename), []


def extract_csv_file(path, filename):
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = [r for r in csv.reader(f)]
    if not rows:
        return None, ["empty csv"]
    return _rows_to_bidder(rows, filename), []


def extract_pdf_single(path, filename):
    """Single-bidder PDF: text tables -> item/amount list."""
    import pdfplumber
    rows = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for t in page.extract_tables():
                for r in t:
                    rows.append([(c or "").strip() for c in r])
    if not rows:
        return None, ["no tables found in PDF"]
    return _rows_to_bidder(rows, filename), []


def merge_single_bids(bidders):
    """Merge per-bidder item lists into shared normalized items (fuzzy desc match)."""
    items = []  # {description, values:[per bidder {amount,status}]}
    for bi, b in enumerate(bidders):
        for it in b["items"]:
            hit = next((m for m in items if desc_match(m["description"], it["description"])), None)
            if hit is None:
                hit = {"description": it["description"], "qty": None, "unit": "",
                       "values": [{"amount": None, "rate": None, "status": "not_priced",
                                   "raw": ""} for _ in bidders]}
                items.append(hit)
            hit["values"][bi] = {"amount": it["raw_amount"], "rate": None,
                                 "status": it["raw_status"], "raw": ""}
    return items


# ---------------------------------------------------------------- routes

@app.route("/")
def index():
    return send_from_directory(BASE, "index.html")


@app.route("/api/health")
def health():
    return jsonify(ok=True)


@app.route("/api/upload", methods=["POST"])
def upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify(error="no files uploaded"), 400
    if len(files) > 10:
        return jsonify(error="max 10 files"), 400

    saved = []
    try:
        for f in files:
            ext = os.path.splitext(f.filename or "")[1].lower()
            if ext not in (".pdf", ".xlsx", ".xls", ".csv"):
                return jsonify(error=f"unsupported file type: {f.filename}"), 400
            fd, path = tempfile.mkstemp(suffix=ext)
            os.close(fd)
            f.save(path)
            saved.append((path, f.filename, ext))

        # decide mode: single PDF that parses as a multi-bidder tab?
        if len(saved) == 1 and saved[0][2] == ".pdf":
            data, errs = extract_pdf_tab(saved[0][0], saved[0][1])
            if data and len(data["bidders"]) >= 2 and data["items"]:
                return jsonify(data)
            # fall through to single-bidder mode
            b, errs2 = extract_pdf_single(saved[0][0], saved[0][1])
            bidders = [b] if b else []
            warnings = errs + errs2
        else:
            bidders, warnings = [], []
            for path, filename, ext in saved:
                if ext == ".pdf":
                    # try tab first
                    data, errs = extract_pdf_tab(path, filename)
                    if data and len(data["bidders"]) >= 2 and data["items"]:
                        return jsonify(data)
                    b, errs2 = extract_pdf_single(path, filename)
                    warnings += errs + errs2
                elif ext in (".xlsx", ".xls"):
                    b, errs2 = extract_excel(path, filename)
                    warnings += errs2
                else:
                    b, errs2 = extract_csv_file(path, filename)
                    warnings += errs2
                if b:
                    bidders.append(b)
            if not bidders:
                return jsonify(error="could not extract bid data",
                               warnings=warnings), 422

        items = merge_single_bids(bidders)
        return jsonify({
            "project_name": "Bid package",
            "mode": "bids",
            "bidders": [{"name": b["name"], "location": b.get("location", ""),
                         "file": b["file"]} for b in bidders],
            "items": items,
            "submitted_totals": [None] * len(bidders),
            "subtotals": [None] * len(bidders),
            "taxes": [None] * len(bidders),
            "warnings": warnings + [
                "single-bid files: bidder names taken from filenames - please confirm"],
            "bidder_confidence": "low",
        })
    finally:
        for path, _, _ in saved:
            try:
                os.unlink(path)
            except OSError:
                pass


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
