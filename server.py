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
    # percent detection: "1.79%" or "1.79 percent"
    pct = re.match(r"^\s*\(?\s*(\d+(?:\.\d+)?)\s*%\s*\)?\s*$", s)
    if pct:
        try:
            return float(pct.group(1)) / 100.0, "percent"
        except ValueError:
            pass
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


def detect_row_type(description, header_text=""):
    """Classify a row: amount | rate | percent | lump | fee_component.
    Never sum rate/percent rows into totals."""
    d = (description or "").lower()
    h = (header_text or "").lower()
    combined = d + " " + h
    # percent rows: explicit % or "percent"/"fee %" language
    if "%" in description or "percent" in combined or re.search(r"\bfee\s*%", combined):
        return "percent"
    # rate rows: $/unit language without a total
    if re.search(r"\$\s*/\s*\w|per\s+(hour|hr|lf|sf|unit|each|sq|ton|day)", combined):
        return "rate"
    # fee components (CMAR/design-build): fee, general conditions, etc.
    if re.search(r"\b(fee|general conditions|preconstruction|cm fee|construction fee)\b", d):
        return "fee_component"
    return "amount"


def verify_cell(qty, rate, amount, tolerance=0.02):
    """Three-state verification: verified | mismatch | unverified.
    Confidence comes from arithmetic, never from OCR."""
    if qty is not None and rate is not None and amount is not None:
        expected = qty * rate
        if abs(expected - amount) <= max(tolerance, abs(expected) * 0.001):
            return "verified"
        return "mismatch"
    return "unverified"


def parse_money_all(raw):
    """All money-like values stacked in one cell (as-bid over corrected, etc.).
    Returns list of (value, text) with dupes removed, preserving order."""
    if raw is None:
        return []
    out, seen = [], set()
    for p in re.split(r"[\r\n]+", str(raw)):
        v, st = parse_money(p)
        if st == "priced" and v not in seen:
            seen.add(v)
            out.append((v, p.strip()))
    return out


def parse_minmax(desc):
    """Item-level min/max bid constraints, e.g. 'Minimum Bid of $7,500.00 Lump Sum'.
    Returns (min_or_None, max_or_None)."""
    d = desc or ""
    mn = mx = None
    m = re.search(r"minimum\s+bid\s+(?:of\s+)?\$?\s*([\d,]+(?:\.\d+)?)", d, re.I)
    if m:
        try:
            mn = float(m.group(1).replace(",", ""))
        except ValueError:
            pass
    m = re.search(r"maximum\s+bid\s+(?:of\s+)?\$?\s*([\d,]+(?:\.\d+)?)", d, re.I)
    if m:
        try:
            mx = float(m.group(1).replace(",", ""))
        except ValueError:
            pass
    return mn, mx


def detect_tax_rate(row):
    """Find '8.50%' style rate in a tax row. Returns decimal or None."""
    joined = " ".join(str(c or "") for c in row)
    m = re.search(r"(\d+(?:\.\d+)?)\s*%", joined)
    if m:
        try:
            return float(m.group(1)) / 100.0
        except ValueError:
            pass
    return None


SERVICE_WORDS = ("mow", "maintenance", "cleaning", "janitorial", "landscape",
                 "lawn", "fertiliz", "pruning", "cleanup", "snow removal",
                 "per visit", "weekly", "monthly", "per season", "service")


def compute_basis_state(items):
    """Separate 'math verified' from 'basis verified'.
    Returns (state, notes): verified | suspect | unknown.
    'unknown' = recurring-service bid with no quantity/frequency dimension —
    the comparison basis is void and the product must refuse to rank."""
    amount_items = [it for it in items if it.get("row_type", "amount") == "amount"]
    if not amount_items:
        return "unknown", ["no amount rows extracted"]
    with_qty = sum(1 for it in amount_items if it.get("qty") is not None)
    notes = []
    service_like = any(
        any(w in (it.get("description") or "").lower() for w in SERVICE_WORDS)
        for it in amount_items)
    if with_qty == 0 and service_like:
        return "unknown", [
            "no quantity or frequency stated anywhere — recurring-service bid. "
            "Per-site prices cannot be compared without a visits-per-period basis."]
    if with_qty == 0:
        notes.append("no quantities stated — lump-sum comparison only; "
                     "spreads reflect scope interpretation, not unit pricing")
    elif with_qty < len(amount_items) * 0.8:
        notes.append(f"only {with_qty}/{len(amount_items)} items carry quantities")
    noncomp = sum(1 for it in amount_items if it.get("comparable") is False)
    if noncomp:
        notes.append(f"{noncomp} row(s) flagged not directly comparable")
    return ("suspect" if notes else "verified"), notes


def compute_ranking(items, bidder_excluded, bidder_names, extraction_quality,
                    low_conf_counts, blocked=False):
    """Ranking with honesty states: ranked | too_close | blocked.
    'blocked' = service bid with no quantity basis — refuse to rank.
    'too_close' = top-two margin below extraction uncertainty."""
    totals = []
    for gi in range(len(bidder_names)):
        if bidder_excluded[gi]:
            totals.append(None)
            continue
        t = sum(v["amount"] for it in items for v in [it["values"][gi]]
                if it.get("row_type", "amount") == "amount"
                and not it.get("deduct") and not it.get("alternate")
                and v["status"] in ("priced", "included")
                and v["amount"] is not None)
        totals.append(t)
    order = sorted((t, gi) for gi, t in enumerate(totals) if t is not None)
    if blocked:
        return {"state": "blocked", "winner": None, "margin": None,
                "margin_pct": None, "totals": totals,
                "note": "No quantity/frequency basis — ranking withheld."}
    if len(order) < 2:
        return {"state": "ranked",
                "winner": order[0][1] if order else None,
                "margin": None, "margin_pct": None, "totals": totals, "note": ""}
    (t1, w1), (t2, w2) = order[0], order[1]
    margin = t2 - t1
    pct = (margin / t1) if t1 else 0
    low_top2 = (low_conf_counts[w1] if w1 < len(low_conf_counts) else 0) + \
               (low_conf_counts[w2] if w2 < len(low_conf_counts) else 0)
    if pct < 0.02 and (low_top2 > 0 or extraction_quality == "poor"):
        return {"state": "too_close", "winner": None, "margin": margin,
                "margin_pct": pct, "totals": totals,
                "note": (f"Top two bids differ by {pct * 100:.1f}% "
                         f"(${margin:,.0f}) — below extraction uncertainty. "
                         "Verify flagged cells before awarding.")}
    return {"state": "ranked", "winner": w1, "margin": margin,
            "margin_pct": pct, "totals": totals, "note": ""}


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
    # scan the whole row: tax rows can carry the rate ("8.5") in the item-number
    # column with the description elsewhere (Sim 15). Avoid matching item
    # descriptions that merely mention tax as a word — require tax + money/%.
    joined = " ".join(norm_desc(c) for c in row if c)
    if not any(w in joined for w in TAX_WORDS):
        return False
    raw = " ".join(str(c or "") for c in row)
    return bool(re.search(r"\d", raw))  # must carry a number (rate or amount)


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
    submitted, subtotals = [None] * n_bidders, [None] * n_bidders
    taxes = [0.0] * n_bidders
    has_tax_rows = False
    tax_rate_detected = None
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
            has_tax_rows = True
            rate = detect_tax_rate(r)
            if rate and tax_rate_detected is None:
                tax_rate_detected = rate
            for gi, (tc, _) in enumerate(groups):
                v, st = parse_money(r[tc] if tc < len(r) else "")
                if st == "priced":
                    taxes[gi] += v  # multiple tax rows (per schedule) accumulate
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
        row_type = detect_row_type(desc, " ".join(h for h in header if h))
        # substitution-alternate structure: deduct rows mirror base items, alternate
        # rows replace them. NEVER auto-combine: the gate lives in the frontend.
        dl = desc.lower()
        is_deduct = "deduct" in dl
        is_alternate = ("alternate" in dl or "in lieu of" in dl) and not is_deduct
        min_bid, max_bid = parse_minmax(desc)
        # comparable flag: qualifier words escalate to human review
        comparable = True
        qualifiers = ("only", "excluding", "not including", "per vendor", "if necessary",
                      "alternate", "deduct", "additive")
        if any(q in dl for q in qualifiers):
            comparable = False
        for gi, (tc, rc) in enumerate(groups):
            amt_raw = r[tc] if tc < len(r) else ""
            rate_raw = r[rc] if rc is not None and rc < len(r) else ""
            rate, _ = parse_money(rate_raw)
            verbatim = str(amt_raw or "").strip()
            # stacked as-bid vs corrected values in one cell (Sim 12): force human pick
            stacked = parse_money_all(amt_raw)
            if len(stacked) > 1:
                vals.append({"amount": None, "rate": rate, "status": "ambiguous",
                             "alternatives": [{"amount": v, "raw": t} for v, t in stacked],
                             "raw": verbatim, "verbatim": verbatim,
                             "confidence": "low", "verify": "unverified",
                             "min_violation": False, "max_violation": False})
                continue
            amt, st = parse_money(amt_raw)
            # confidence: high if verbatim parses cleanly, medium if repaired, low if ambiguous
            confidence = "high"
            if st == "text":
                confidence = "low"
            elif verbatim != clean_num_text(verbatim):
                confidence = "medium"
            # verification: check qty x rate = amount where possible
            verify = verify_cell(qty, rate, amt) if st == "priced" else "unverified"
            # item-level min/max bid constraints (Sim 15): flag, never silently carry
            min_viol = min_bid is not None and st == "priced" and amt is not None and amt < min_bid - 0.005
            max_viol = max_bid is not None and st == "priced" and amt is not None and amt > max_bid + 0.005
            if min_viol:
                warnings.append(
                    f"{bidder_names[gi]}: '{desc[:60]}' bid ${amt:,.2f} is below the "
                    f"stated minimum ${min_bid:,.2f} — responsiveness risk, verify with owner")
                confidence = "low"
            vals.append({"amount": amt, "rate": rate, "status": st,
                         "raw": verbatim, "verbatim": verbatim,
                         "confidence": confidence, "verify": verify,
                         "min_violation": min_viol, "max_violation": max_viol})
        # skip rows with no priced data at all (section headers etc.)
        if not any(v["status"] in ("priced", "included", "ambiguous") for v in vals):
            # keep rows where at least one bidder has not_priced but others priced
            if all(v["status"] == "not_priced" or v["status"] == "text" for v in vals):
                continue
        items.append({"description": desc, "qty": qty, "unit": unit_raw, "values": vals,
                      "row_type": row_type, "comparable": comparable,
                      "deduct": is_deduct, "alternate": is_alternate,
                      "min_bid": min_bid, "max_bid": max_bid,
                      "frequency": None,  # user-supplied visits-per-period (service bids)
                      "base_value": None})  # base_value set for percent rows when known

    has_substitution_alternates = (any(it.get("deduct") for it in items)
                                   and any(it.get("alternate") for it in items))
    if has_substitution_alternates:
        warnings.append(
            "Substitution alternate detected (deduct rows + alternate rows). "
            "Combined totals are WITHHELD until you confirm how each bidder's deduct "
            "maps to the replaced base items — adding the alternate without subtracting "
            "the per-bidder deducts overstates every total.")
    n_ambiguous = sum(1 for it in items for v in it["values"] if v["status"] == "ambiguous")
    if n_ambiguous:
        warnings.append(
            f"{n_ambiguous} cell(s) contain stacked as-bid AND corrected values — "
            "pick the authoritative value for each before leveling.")

    def _base_calc(gi):
        # NEVER sum rate/percent/fee_component rows; NEVER sum deduct/alternate
        # scenario rows into the base total.
        return sum(v["amount"] for it in items for v in [it["values"][gi]]
                   if it.get("row_type", "amount") == "amount"
                   and not it.get("deduct") and not it.get("alternate")
                   and v["status"] in ("priced", "included") and v["amount"] is not None)

    # math verification with three-state totals (verified / mismatch / unparseable).
    # Tax-aware: a tax-inclusive submitted total must match items + extracted tax
    # (Sim 15 selective-tax tabs). Try tax-inclusive first, then tax-exclusive.
    total_states = []
    tax_basis_detected = []
    for gi in range(n_bidders):
        calc = _base_calc(gi)
        tax = taxes[gi] if has_tax_rows else 0.0
        base = subtotals[gi] if subtotals[gi] is not None else submitted[gi]
        n_amount_rows = sum(1 for it in items
                            if it.get("row_type", "amount") == "amount"
                            and not it.get("deduct") and not it.get("alternate"))
        state, basis = "unparseable", None
        if base is not None and n_amount_rows > 0:
            tol = max(1.0, abs(base) * 0.002)
            if has_tax_rows and abs((calc + tax) - base) <= tol:
                state, basis = "verified", "tax_included"
            elif abs(calc - base) <= tol:
                state, basis = "verified", "tax_excluded"
            else:
                state = "mismatch"
                warnings.append(
                    f"{bidder_names[gi]}: line items sum to "
                    f"${calc:,.2f}" + (f" + ${tax:,.2f} tax" if has_tax_rows else "") +
                    f" but {'subtotal' if subtotals[gi] is not None else 'total'} "
                    f"shows ${base:,.2f} - showing BOTH, human must adjudicate")
        # award-basis suggestion comes from the SUBMITTED total (the actual bid):
        # if the bid tab's bottom line includes tax, awards are tax-inclusive.
        award_b = None
        sub = submitted[gi]
        if sub is not None and n_amount_rows > 0:
            tol_s = max(1.0, abs(sub) * 0.002)
            if has_tax_rows and abs((calc + tax) - sub) <= tol_s:
                award_b = "tax_included"
            elif abs(calc - sub) <= tol_s:
                award_b = "tax_excluded"
        total_states.append(state)
        tax_basis_detected.append(award_b)

    # disqualified / all-zero bidder detection
    bidder_excluded = []
    for gi in range(n_bidders):
        priced = sum(1 for it in items
                     if it["values"][gi]["status"] == "priced"
                     and it["values"][gi]["amount"] not in (None, 0))
        bidder_excluded.append(priced == 0)
        if priced == 0:
            warnings.append(
                f"{bidder_names[gi]}: no priced line items - excluded from "
                f"low-bidder calculation and statistics")

    # extraction quality: fraction of money-column cells that parse
    data_cells = parsed_cells = 0
    for r in data[hdr_idx + 1:]:
        for (tc, rc) in groups:
            for c in (tc, rc):
                if c is not None and c < len(r):
                    cell = (r[c] or "").strip()
                    if cell:
                        data_cells += 1
                        if parse_money(cell)[1] in ("priced", "not_priced", "included", "percent"):
                            parsed_cells += 1
    extraction_quality = "good" if (data_cells == 0 or parsed_cells / data_cells >= 0.6) else "poor"

    # basis state (math vs basis badges) + ranking honesty states
    basis_state, basis_notes = compute_basis_state(items)
    low_conf_counts = [sum(1 for it in items if it["values"][gi]["confidence"] == "low")
                       for gi in range(n_bidders)]
    ranking = compute_ranking(items, bidder_excluded, bidder_names,
                              extraction_quality, low_conf_counts,
                              blocked=(basis_state == "unknown"))
    if ranking["state"] == "blocked":
        warnings.append(
            "⛔ Cannot rank these bids: " + " ".join(basis_notes) +
            " Enter a visits-per-period frequency to see a PROVISIONAL comparison.")
    elif ranking["state"] == "too_close":
        warnings.append("⚠ " + ranking["note"])

    # value-weighted ranking: weight rows by dollar value, not count
    # score = sum over rows of (bidder_price / median_price * row_median_value)
    # lower score = better value
    value_scores = []
    for gi in range(n_bidders):
        if bidder_excluded[gi]:
            value_scores.append(None)
            continue
        score = 0.0
        for it in items:
            if it.get("row_type", "amount") != "amount":
                continue
            if it.get("deduct") or it.get("alternate"):
                continue
            row_vals = [it["values"][g]["amount"] for g in range(n_bidders)
                        if not bidder_excluded[g]
                        and it["values"][g]["status"] == "priced"
                        and it["values"][g]["amount"] not in (None, 0)]
            if len(row_vals) < 2:
                continue
            row_vals.sort()
            med = row_vals[len(row_vals) // 2]
            row_med_val = med  # weight by median dollar value of this row
            my_amt = it["values"][gi]["amount"]
            if my_amt is None or my_amt <= 0 or med <= 0:
                continue
            score += (my_amt / med) * row_med_val
        value_scores.append(score if score > 0 else None)

    return {
        "project_name": os.path.splitext(filename)[0].replace("_", " ").replace("-", " "),
        "mode": "tab",
        "bidders": [{"name": n, "location": l, "file": filename,
                     "excluded": ex, "total_state": ts}
                    for n, l, ex, ts in zip(bidder_names, bidder_locs,
                                           bidder_excluded, total_states)],
        "items": items,
        "submitted_totals": submitted,
        "subtotals": subtotals,
        "taxes": taxes,
        "has_tax_rows": has_tax_rows,
        "tax_rate_detected": tax_rate_detected,
        "tax_basis_detected": tax_basis_detected,
        "has_substitution_alternates": has_substitution_alternates,
        "basis_state": basis_state,
        "basis_notes": basis_notes,
        "ranking": ranking,
        "extraction_quality": extraction_quality,
        "warnings": warnings,
        "bidder_confidence": name_conf,
        "value_scores": value_scores,
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
                dl = (it["description"] or "").lower()
                mn, mx = parse_minmax(it["description"])
                hit = {"description": it["description"], "qty": None, "unit": "",
                       "row_type": "amount", "comparable": True, "base_value": None,
                       "deduct": "deduct" in dl,
                       "alternate": ("alternate" in dl or "in lieu of" in dl) and "deduct" not in dl,
                       "min_bid": mn, "max_bid": mx, "frequency": None,
                       "values": [{"amount": None, "rate": None, "status": "not_priced",
                                   "raw": "", "verbatim": "", "confidence": "low",
                                   "verify": "unverified",
                                   "min_violation": False, "max_violation": False}
                                  for _ in bidders]}
                items.append(hit)
            mn, mx = hit.get("min_bid"), hit.get("max_bid")
            amt = it["raw_amount"]
            min_viol = mn is not None and it["raw_status"] == "priced" and amt is not None and amt < mn - 0.005
            hit["values"][bi] = {"amount": amt, "rate": None,
                                 "status": it["raw_status"], "raw": "",
                                 "verbatim": "", "confidence": "medium",
                                 "verify": "unverified",
                                 "min_violation": min_viol, "max_violation": False}
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
        # bids mode: no submitted totals to verify against
        n_b = len(bidders)
        bidder_excluded = []
        for bi in range(n_b):
            priced = sum(1 for it in items
                         if it["values"][bi]["status"] == "priced"
                         and it["values"][bi]["amount"] not in (None, 0))
            bidder_excluded.append(priced == 0)
        basis_state, basis_notes = compute_basis_state(items)
        low_conf_counts = [sum(1 for it in items if it["values"][bi]["confidence"] == "low")
                           for bi in range(n_b)]
        ranking = compute_ranking(items, bidder_excluded,
                                  [b["name"] for b in bidders],
                                  "good", low_conf_counts,
                                  blocked=(basis_state == "unknown"))
        warnings_b = warnings + [
            "single-bid files: bidder names taken from filenames - please confirm"]
        if ranking["state"] == "blocked":
            warnings_b.append("⛔ Cannot rank these bids: " + " ".join(basis_notes))
        elif ranking["state"] == "too_close":
            warnings_b.append("⚠ " + ranking["note"])
        return jsonify({
            "project_name": "Bid package",
            "mode": "bids",
            "bidders": [{"name": b["name"], "location": b.get("location", ""),
                         "file": b["file"], "excluded": ex, "total_state": "unparseable"}
                        for b, ex in zip(bidders, bidder_excluded)],
            "items": items,
            "submitted_totals": [None] * len(bidders),
            "subtotals": [None] * len(bidders),
            "taxes": [0.0] * len(bidders),
            "has_tax_rows": False,
            "tax_rate_detected": None,
            "tax_basis_detected": [None] * len(bidders),
            "has_substitution_alternates": (
                any(it.get("deduct") for it in items)
                and any(it.get("alternate") for it in items)),
            "basis_state": basis_state,
            "basis_notes": basis_notes,
            "ranking": ranking,
            "extraction_quality": "good",
            "warnings": warnings_b,
            "bidder_confidence": "low",
            "value_scores": [None] * len(bidders),
        })
    finally:
        for path, _, _ in saved:
            try:
                os.unlink(path)
            except OSError:
                pass


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
