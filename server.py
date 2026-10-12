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
import secrets
import tempfile
from datetime import datetime, timezone

from flask import Flask, request, jsonify, send_from_directory, session, g

import auth_billing
from auth_billing import get_db, getpaid_required

app = Flask(__name__)
BASE = os.path.dirname(os.path.abspath(__file__))
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-only-change-me")
auth_billing.init_db()
auth_billing.register_routes(app)

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
    # alternate/deduct qualifiers: "ADD $8,400", "DEDUCT $2,000", "Alt. $1,200".
    # The qualifier word carries the row semantics (flagged elsewhere); the
    # amount itself parses as a positive magnitude. A leading alpha run is
    # tolerated for cell-boundary bleed ("sADD $12,750" — trailing letter of
    # an adjacent description glued to the amount), but approximation words
    # ("approx $500") must NOT parse: only qualifier-led strings strip.
    s = re.sub(r"^[A-Za-z]*?(add|deduct|alternate|alt\.?)\s+", "", s,
               flags=re.IGNORECASE)
    low = s.lower()
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
                 "per visit", "weekly", "monthly", "per season",
                 "service contract", "service agreement")
# NOTE: bare "service" intentionally excluded — it false-positives on
# "service upgrade", "electrical service", "main service panel", etc.


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
                    low_conf_counts, blocked=False, total_states=None):
    """Ranking with honesty states: ranked | too_close | blocked | needs_review.
    'blocked' = service bid with no quantity basis — refuse to rank.
    'too_close' = top-two margin below extraction uncertainty.
    'needs_review' = totals not verified or single valid bidder — refuse to
    crown a winner from garbage (Round 4 fix: Sim 22, 24, 25)."""
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
    # Round 4 fix: refuse to rank from unverified totals or a single bidder.
    # A lone "bidder" is usually a phantom from a failed parse, not a real winner.
    n_valid = len(order)
    if total_states:
        unverified = [bidder_names[gi] for gi, t in enumerate(totals)
                      if t is not None and total_states[gi] != "verified"]
    else:
        unverified = []
    if n_valid < 2:
        return {"state": "needs_review", "winner": None,
                "margin": None, "margin_pct": None, "totals": totals,
                "note": ("Only one bidder with priced items — cannot rank. "
                         "Verify the tab parsed correctly before awarding.")}
    if unverified:
        return {"state": "needs_review", "winner": None,
                "margin": None, "margin_pct": None, "totals": totals,
                "note": ("Bid totals not verified for: " + ", ".join(unverified) +
                         " — ranking withheld until totals are confirmed.")}
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


def desc_match(a, b, thresh=0.75):
    # Round 4 fix: raised from 0.55 to 0.75 — 0.55 collapsed 5 distinct
    # alternates into one row and could merge deduct+additive rows (Sim 16, 18).
    ta, tb = set(norm_desc(a).split()), set(norm_desc(b).split())
    if not ta or not tb:
        return False
    return len(ta & tb) / max(len(ta), len(tb)) >= thresh


def _has_deduct_lang(desc):
    """True if description carries deduct/substitution language."""
    dl = (desc or "").lower()
    return "deduct" in dl or "delete" in dl or "credit" in dl


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


def _row_labels(row, desc_idx=None):
    """Candidate label cells: column 0 plus the detected description column
    (total/subtotal labels sit in the description column when an item-number
    column leads)."""
    labs = [row_label(row)]
    if desc_idx is not None and desc_idx < len(row) and desc_idx != 0:
        labs.append(norm_desc(row[desc_idx]))
    return labs


def is_total_row(row, desc_idx=None):
    labs = _row_labels(row, desc_idx)
    return any(any(w == lab or lab.startswith(w + " ") for w in TOTAL_WORDS)
               or lab in TOTAL_WORDS for lab in labs)


def is_subtotal_row(row, desc_idx=None):
    return any(lab in SUBTOTAL_WORDS for lab in _row_labels(row, desc_idx))


def is_tax_row(row):
    # scan the whole row: tax rows can carry the rate ("8.5") in the item-number
    # column with the description elsewhere (Sim 15). Avoid matching item
    # descriptions that merely mention tax as a word — require tax + money/%.
    # Round 4 fix: use word boundaries (not substring) and explicitly exclude
    # "non-taxable" / "tax-exempt" / "no tax" phrases (Sim 20).
    joined = " ".join(norm_desc(c) for c in row if c)
    # explicit exclusions first: these are NOT tax rows
    if re.search(r"\bnon[\s-]?taxable\b|\btax[\s-]?exempt\b|\bno tax\b|\btax free\b", joined):
        return False
    # word-boundary match on "tax" or "sales tax" (not substring like "taxable")
    if not re.search(r"\bsales tax\b|\btax\b", joined):
        return False
    raw = " ".join(str(c or "") for c in row)
    return bool(re.search(r"\d", raw))  # must carry a number (rate or amount)


# ---------------------------------------------------------------- PDF (multi-bidder tab mode)

def _pdf_full_text(path):
    """Extract all text from a PDF, page by page."""
    import pdfplumber
    pages = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text() or "")
    return pages


def detect_verdantas_format(pages):
    """Detect Verdantas bid-platform summary: LIST OF BIDDERS + LIST OF TOTALS pages."""
    full = "\n".join(pages).upper()
    has_bidders = "LIST OF BIDDERS" in full
    has_totals = "LIST OF TOTALS" in full
    has_tab_summary = "BID TABULATION SUMMARY" in full
    return has_bidders and has_totals and has_tab_summary


def extract_verdantas_summary(path, filename):
    """Parse Verdantas bid tabulation summary: bidder list + totals pages.

    Returns (data, warnings) in the same shape as extract_pdf_tab, or (None, [errs]).
    """
    import re
    pages = _pdf_full_text(path)
    if not detect_verdantas_format(pages):
        return None, ["not a Verdantas summary format"]

    warnings = []
    # Find the LIST OF BIDDERS page and LIST OF TOTALS page
    bidder_names = []  # list of (number, name)
    totals = {}  # number -> (name, amount)

    for page_text in pages:
        upper = page_text.upper()
        if "LIST OF BIDDERS" in upper:
            # Pattern: number followed by company name, then address lines
            # e.g. "1 Perk Company Inc.\n3740 Carnegie Avenue..."
            lines = page_text.split("\n")
            i = 0
            while i < len(lines):
                m = re.match(r"^(\d+)\s+(.+?)\s*$", lines[i].strip())
                if m:
                    num = int(m.group(1))
                    name = m.group(2).strip()
                    # Company name is on this line; address follows on next lines
                    # Clean up: remove trailing stuff, keep the name
                    if name and len(name) > 2:
                        bidder_names.append((num, name))
                i += 1
        elif "LIST OF TOTALS" in upper:
            # Pattern: "1. Perk Company Inc. $557,644.50"
            # May have informal total: "3. A & J Cement $1,114,341.30 $1,114,336.30"
            # We take the FIRST (calculated) total.
            for line in page_text.split("\n"):
                m = re.match(r"^(\d+)\.\s+(.+?)\s+\$([\d,]+\.\d{2})(?:\s+\$[\d,]+\.\d{2})?\s*$", line.strip())
                if m:
                    num = int(m.group(1))
                    name = m.group(2).strip()
                    # Strip any trailing dollar amount that leaked into the name
                    name = re.sub(r"\s+\$[\d,]+\.\d{2}\s*$", "", name).strip()
                    val, status = parse_money(m.group(3))
                    if val is not None and status == "priced":
                        totals[num] = (name, val)

    if not bidder_names:
        return None, ["Verdantas format detected but no bidders found"]
    if not totals:
        return None, ["Verdantas format detected but no totals found"]

    # Build bidder data in the standard format
    bidders = []
    bidder_name_list = []
    for num, name in sorted(bidder_names):
        if num in totals:
            tname, amount = totals[num]
            # Use the totals-page name (more likely correct)
            final_name = tname if tname else name
            bidder_name_list.append(final_name)
            bidders.append({
                "name": final_name,
                "total": amount,
                "total_state": "verified",
                "items": [{
                    "description": "Total Bid",
                    "amount": amount,
                    "confidence": "high",
                }],
                "confidence": "high",
            })
        else:
            warnings.append(f"Bidder {name} has no total")

    if len(bidders) < 2:
        return None, ["Verdantas: fewer than 2 bidders with totals"]

    # Build the standard tab output format
    items = [{
        "description": "Total Bid",
        "bids": {b["name"]: {"amount": b["total"], "confidence": "high"} for b in bidders},
        "comparable": True,
    }]

    return {
        "bidders": bidder_name_list,
        "bidder_data": {b["name"]: b for b in bidders},
        "items": items,
        "extraction_quality": "good",
        "format": "verdantas_summary",
        "warnings": warnings,
    }, warnings


def detect_mass_housing(pages):
    """Detect MA Housing Authority bid tab: GENERAL BID TABULATION + HOUSING AUTHORITY."""
    full = "\n".join(pages[:2]).upper()
    return "GENERAL BID TABULATION" in full and "HOUSING AUTHORITY" in full


def extract_mass_housing(path, filename):
    """Parse MA Housing Authority bid tab: Contractor Name & Address | Bid Amount | ...

    Standardized form used across MA housing authorities. Rows have contractor
    names and bid amounts. Returns (data, warnings) or (None, [errs]).
    """
    import re
    import pdfplumber
    pages = _pdf_full_text(path)
    if not detect_mass_housing(pages):
        return None, ["not a MA housing authority tab"]

    warnings = []
    bidders = []

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            tables = page.find_tables()
            for t in tables:
                data = t.extract()
                if not data or len(data) < 2:
                    continue
                # Find the header row with "Contractor" and "Bid Amount"
                hdr_idx = None
                for i, row in enumerate(data):
                    row_text = " ".join([c or "" for c in row]).upper()
                    # Handle doubled-text rendering artifact
                    row_text = re.sub(r"(.)\1", r"\1", row_text)
                    if "CONTRACTOR" in row_text and "BID" in row_text:
                        hdr_idx = i
                        break
                if hdr_idx is None:
                    continue
                # Find bid amount column
                header = data[hdr_idx]
                amt_col = None
                name_col = 0
                for j, h in enumerate(header):
                    if not h:
                        continue
                    hn = re.sub(r"(.)\1", r"\1", h.upper())
                    if "BID" in hn and "AMOUNT" in hn:
                        amt_col = j
                    if "CONTRACTOR" in hn:
                        name_col = j
                if amt_col is None:
                    continue
                # Extract rows
                for row in data[hdr_idx + 1:]:
                    if len(row) <= max(name_col, amt_col):
                        continue
                    name = (row[name_col] or "").strip()
                    amt_raw = (row[amt_col] or "").strip()
                    # Clean doubled-text artifact
                    name = re.sub(r"(.)\1", r"\1", name)
                    if not name or len(name) < 3:
                        continue
                    val, status = parse_money(amt_raw)
                    if val is not None and status == "priced":
                        bidders.append((name, val))
                    elif amt_raw:
                        warnings.append(f"Could not parse amount for {name}: {amt_raw}")

    if len(bidders) < 2:
        return None, ["MA housing: fewer than 2 bidders with amounts"]

    # Deduplicate by name, keep first
    seen = set()
    unique = []
    for name, val in bidders:
        if name.lower() not in seen:
            seen.add(name.lower())
            unique.append((name, val))
    bidders = unique

    if len(bidders) < 2:
        return None, ["MA housing: fewer than 2 unique bidders"]

    bidder_names = [n for n, _ in bidders]
    items = [{
        "description": "Total Bid",
        "bids": {n: {"amount": v, "confidence": "medium"} for n, v in bidders},
        "comparable": True,
    }]

    return {
        "bidders": bidder_names,
        "bidder_data": {
            n: {
                "name": n, "total": v, "total_state": "unverified",
                "items": [{"description": "Total Bid", "amount": v, "confidence": "medium"}],
                "confidence": "medium",
            } for n, v in bidders
        },
        "items": items,
        "extraction_quality": "fair",
        "format": "mass_housing",
        "warnings": warnings,
    }, warnings


def detect_board_packet(pages):
    """Detect school board / council packet: ACTION REPORT, Board of Education, etc."""
    full = "\n".join(pages[:3]).upper()  # check first 3 pages
    indicators = [
        "ACTION REPORT",
        "BOARD OF EDUCATION",
        "CITY COUNCIL",
        "BOARD MEETING",
        "RECOMMENDATION",
        "AWARD A CONTRACT",
    ]
    score = sum(1 for ind in indicators if ind in full)
    # Also check for bid language
    has_bid_lang = any(
        phrase in full
        for phrase in ["SUBMITTED BIDS", "LOW BIDDER", "BID OPENING", "PUBLICLY BID"]
    )
    return score >= 2 and has_bid_lang


def extract_board_packet(path, filename):
    """Extract bid info from board/council packet narrative.

    These PDFs bury bid results in prose. We extract what's stated:
    bidder count, low bidder name/amount. Returns needs_review (never ranked).
    """
    import re
    pages = _pdf_full_text(path)
    if not detect_board_packet(pages):
        return None, ["not a board packet format"]

    full = "\n".join(pages)
    upper = full.upper()
    warnings = []

    # Extract number of bidders: "Nine contractors submitted bids" / "9 bidders"
    # NOTE: `upper` is uppercased text, so these patterns need re.IGNORECASE.
    bidder_count = None
    m = re.search(r"(\d+)\s+(contractors?|bidders?)\s+submitted\s+bids?", upper,
                  re.IGNORECASE)
    if m:
        bidder_count = int(m.group(1))
    else:
        # Try word numbers
        word_nums = {"ONE": 1, "TWO": 2, "THREE": 3, "FOUR": 4, "FIVE": 5,
                     "SIX": 6, "SEVEN": 7, "EIGHT": 8, "NINE": 9, "TEN": 10,
                     "ELEVEN": 11, "TWELVE": 12}
        m = re.search(r"\b(ONE|TWO|THREE|FOUR|FIVE|SIX|SEVEN|EIGHT|NINE|TEN|ELEVEN|TWELVE)\b\s+(contractors?|bidders?)\s+submitted", upper,
                      re.IGNORECASE)
        if m:
            bidder_count = word_nums[m.group(1)]

    # Extract low bidder: "The low bidder is X ... in the amount of $Y"
    # or "low bidder is X in the amount of Base Bid $Y"
    low_bidder = None
    low_amount = None
    m = re.search(
        r"low\s+bidder\s+is\s+(.+?)\s+(?:from\s+.+?\s+)?in\s+the\s+amount\s+of\s+(?:base\s+bid\s+)?\$([\d,]+)",
        full, re.IGNORECASE
    )
    if m:
        low_bidder = m.group(1).strip().rstrip(",")
        val, status = parse_money(m.group(2))
        low_amount = val if status == "priced" else None

    # Also try: "Award a contract to X for ... for $Y"
    if not low_bidder:
        m = re.search(
            r"award\s+a\s+contract\s+to\s+(.+?)\s+for\s+.+?\s+for\s+\$([\d,]+)",
            full, re.IGNORECASE
        )
        if m:
            low_bidder = m.group(1).strip().rstrip(",")
            val, status = parse_money(m.group(2))
            low_amount = val if status == "priced" else None

    # Also try: "lowest responsible bid was submitted by X, at $Y"
    if not low_bidder:
        m = re.search(
            r"lowest\s+(?:responsible\s+)?bid\s+was\s+submitted\s+by\s+(.+?)\s*,?\s+at\s+\$([\d,]+\.?\d*)",
            full, re.IGNORECASE
        )
        if m:
            low_bidder = m.group(1).strip().rstrip(",")
            val, status = parse_money(m.group(2))
            low_amount = val if status == "priced" else None

    # Also try: "low bidder for the project is X, from Y" + separate "low bid is $Z"
    if not low_bidder:
        m = re.search(
            r"low\s+bidder\s+(?:for\s+the\s+project\s+)?is\s+(.+?)\s*,?\s+from\s+",
            full, re.IGNORECASE
        )
        if m:
            low_bidder = m.group(1).strip().rstrip(",")
            # Find the low bid amount in nearby text
            m2 = re.search(
                r"low\s+bid\s+is\s+(?:approximately\s+)?\$([\d,]+\.?\d*)",
                full, re.IGNORECASE
            )
            if m2:
                val, status = parse_money(m2.group(1))
                low_amount = val if status == "priced" else None

    if not low_bidder or low_amount is None:
        return None, ["board packet detected but could not extract low bidder"]

    # Build a single-bidder informational result (never ranked)
    info = {
        "bidders": [low_bidder],
        "bidder_data": {
            low_bidder: {
                "name": low_bidder,
                "total": low_amount,
                "total_state": "unverified",
                "items": [{
                    "description": "Low Bid (from narrative)",
                    "amount": low_amount,
                    "confidence": "low",
                }],
                "confidence": "low",
            }
        },
        "items": [],
        "extraction_quality": "poor",
        "format": "board_packet",
        "bidder_count_stated": bidder_count,
        "warnings": [
            f"Board packet narrative only: {bidder_count or 'unknown number of'} bidders, "
            f"low bidder {low_bidder} at ${low_amount:,.2f}. "
            "Full tab not in document — verify before awarding.",
        ] + warnings,
    }
    return info, warnings


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
                            "bbox": t.bbox, "page_width": page.width,
                            "page": page.page_number, "text_fallback": False})
    return out


def _words_to_grid(words):
    """Reconstruct a table grid from pdfplumber words (borderless PDFs).

    Finds the header line (a line whose text reads like a bid-tab header),
    derives column intervals from the header words' x positions, then assigns
    every word to a column by its x-center. Returns (data, col_x) or None when
    no header line is found.
    """
    # group words into lines by top
    lines = {}
    for w in words:
        key = round(w["top"] / 3)  # ~3pt line tolerance
        lines.setdefault(key, []).append(w)
    ordered = [sorted(lines[k], key=lambda w: w["x0"]) for k in sorted(lines)]

    hdr_li = next(
        (i for i, lw in enumerate(ordered)
         if is_header_row([" ".join(w["text"] for w in lw)])),
        None,
    )
    if hdr_li is None:
        return None
    hdr_words = ordered[hdr_li]
    # column intervals from header words; merge horizontally overlapping words
    intervals = []
    for w in hdr_words:
        x0, x1 = w["x0"], w["x1"]
        if intervals and x0 <= intervals[-1][1] + 1:
            intervals[-1][1] = max(intervals[-1][1], x1)
        else:
            intervals.append([x0, x1])
    if len(intervals) < 4:
        return None
    # expand each interval to the midpoint gap with its neighbors
    bounds = []
    for i, (x0, x1) in enumerate(intervals):
        lo = intervals[i - 1][1] + (x0 - intervals[i - 1][1]) / 2 if i > 0 else 0
        hi = x1 + (intervals[i + 1][0] - x1) / 2 if i + 1 < len(intervals) else 1e9
        bounds.append((lo, hi))
    col_x = [(a + b) / 2 for a, b in intervals]

    def _col_of(w):
        cx = (w["x0"] + w["x1"]) / 2
        for i, (lo, hi) in enumerate(bounds):
            if lo <= cx < hi:
                return i
        return len(bounds) - 1 if cx >= bounds[-1][1] else 0

    data = []
    for lw in ordered:
        row = [""] * len(intervals)
        for w in lw:
            ci = _col_of(w)
            row[ci] = (row[ci] + " " + w["text"]).strip()
        if any(c.strip() for c in row):
            data.append(row)
    return data, col_x


def _pdf_tables_text_fallback(path):
    """Fallback for borderless PDFs: rebuild tables from word positions.

    Only pages whose reconstructed grid has a recognizable bid-tab header are
    kept, so prose pages can never hallucinate a table. Returns the same shape
    as _pdf_tables_with_boxes.
    """
    import pdfplumber
    out = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                continue
            grid = _words_to_grid(words)
            if grid is None:
                continue
            data, col_x = grid
            # plausibility gate: header within first rows, enough columns/rows
            if not any(is_header_row(r) for r in data[:6]):
                continue
            if len(data) < 3:
                continue
            out.append({"data": data, "col_x": col_x,
                        "bbox": None, "page_width": page.width,
                        "page": page.page_number, "text_fallback": True})
    return out


def _header_signature(table):
    """Normalized header-row signature for continuation-page matching."""
    for r in table["data"]:
        if is_header_row(r):
            return (len(r), tuple(norm_desc(c) for c in r))
    return None


def _merge_tab_tables(tables):
    """Merge continuation tables (same header on later pages) into one table.

    Returns (data, col_x, warnings). Tables with a different header signature
    are NOT merged; instead a warning names the dropped pages so data loss is
    never silent.
    """
    # primary: largest table (by cells) that has a recognizable header
    def _cells(t):
        return len(t["data"]) * max((len(r) for r in t["data"]), default=0)
    with_hdr = [t for t in tables if _header_signature(t) is not None]
    primary = max(with_hdr or tables, key=_cells)
    prim_sig = _header_signature(primary)
    prim_hdr_idx = next(
        (i for i, r in enumerate(primary["data"]) if is_header_row(r)), 0)

    data = [list(r) for r in primary["data"]]
    col_x = primary["col_x"]
    warnings = []
    merged_pages, dropped = [], []
    for t in tables:
        if t is primary:
            continue
        sig = _header_signature(t)
        if sig is not None and prim_sig is not None and sig == prim_sig:
            # continuation page: append rows after its header row
            hdr_i = next(i for i, r in enumerate(t["data"]) if is_header_row(r))
            for r in t["data"][hdr_i + 1:]:
                row = list(r)
                while len(row) < len(data[0]):
                    row.append("")
                data.append(row[:len(data[0])])
            merged_pages.append(t["page"])
        else:
            n_rows = sum(1 for r in t["data"] if any((c or "").strip() for c in r))
            dropped.append((t["page"], n_rows))
    if merged_pages:
        warnings.append(
            f"bid table continues on page(s) {', '.join(map(str, sorted(set(merged_pages))))} "
            "— continuation rows merged into one tabulation")
    for pg, n_rows in dropped:
        warnings.append(
            f"page {pg}: separate table detected with a different layout "
            f"({n_rows} non-empty rows) — only the main bid table was parsed; "
            "verify those rows are not part of this bid")
    return data, col_x, warnings


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
    """Parse a bid-tabulation PDF: one table, multiple bidder column groups.

    Continuation tables (same header on later pages) are merged; tables with a
    different layout produce an explicit warning so data loss is never silent.
    Borderless PDFs fall back to word-position grid reconstruction.
    """
    tables = _pdf_tables_with_boxes(path)
    if not tables:
        tables = _pdf_tables_text_fallback(path)
    if not tables:
        return None, ["no tables found in PDF"]
    data, col_x, merge_warnings = _merge_tab_tables(tables)
    result, warnings = extract_tab_from_rows(data, col_x, filename)
    return result, merge_warnings + warnings


def extract_tab_from_rows(data, col_x, filename):
    """Parse a multi-bidder bid tabulation from table rows.

    data: list of row-lists. col_x: per-column x-centers (PDF) or None
    (Excel/CSV — name mapping then uses column-index proximity).
    Returns (result_dict, warnings) or (None, [errors])."""
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

    def _split_name(txt):
        parts = [p.strip() for p in txt.split("\n") if p.strip()]
        return (parts[0] if parts else txt,
                parts[1] if len(parts) > 1 else "")

    if name_row_idx is not None and best_score > 0:
        name_row = data[name_row_idx]
        if col_x:
            # PDF: map each group's x-center to nearest name cell x-center.
            # Prefer the group's OWN columns first: borderless-PDF text grids can
            # split a name's words across the group's Rate/Total pair
            # ("ABC" | "Construction LLC"), so join the non-empty cells found
            # in the group's own columns (requiring at least one company hint
            # among them) before falling back to nearest-x.
            import re as _re
            _junk_cell = _re.compile(r"^[\d\s\-\.\(\)\+xX\$,\/]+$").match
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
            used_cols = set()
            used = set()
            for gi, gx in enumerate(group_x):
                tc, rc = groups[gi]
                own_cols = sorted(c for c in (tc, rc)
                                  if c is not None and c >= first_bid_col
                                  and c < len(name_row) and c not in used_cols)
                own_texts = [(name_row[c] or "").strip() for c in own_cols]
                own_texts = [t for t in own_texts if t and not _junk_cell(t)]
                if own_texts and any(any(h in t.lower() for h in COMPANY_HINTS)
                                     for t in own_texts):
                    # join split name words from the group's own columns
                    used_cols.update(own_cols)
                    for ni, (txt, nx) in enumerate(name_cells):
                        if ni not in used and any(
                                txt == (name_row[c] or "").strip() for c in own_cols):
                            used.add(ni)
                    bidder_names[gi], bidder_locs[gi] = _split_name(" ".join(own_texts))
                    continue
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
                    bidder_names[gi], bidder_locs[gi] = _split_name(name_cells[best][0])
        else:
            # Excel/CSV: map by column-index proximity — check the group's own
            # columns (and immediate neighbors) in the name row.
            used_cols = set()
            for gi, (tc, rc) in enumerate(groups):
                cand_cols = []
                for c in (tc, rc):
                    if c is not None:
                        cand_cols += [c - 1, c, c + 1]
                found = None
                # prefer cells with company hints, then any non-empty text
                for c in cand_cols:
                    if c in used_cols or c < first_bid_col or c >= len(name_row):
                        continue
                    cell = (name_row[c] or "").strip()
                    if cell and any(h in cell.lower() for h in COMPANY_HINTS):
                        found = (c, cell)
                        break
                if found is None:
                    for c in cand_cols:
                        if c in used_cols or c < first_bid_col or c >= len(name_row):
                            continue
                        cell = (name_row[c] or "").strip()
                        if cell and len(cell) > 2:
                            found = (c, cell)
                            break
                if found:
                    used_cols.add(found[0])
                    bidder_names[gi], bidder_locs[gi] = _split_name(found[1])
        if all(bidder_names):
            name_conf = "high"
        elif any(bidder_names):
            name_conf = "medium"
            warnings.append("some bidder names could not be mapped to columns - please confirm")
        else:
            warnings.append("bidder names not found - please enter them manually")
    else:
        # No name row above header — check if header itself has bidder names
        # (e.g. CSV with "Acme Construction" as column header).
        # Coe 2026-10-10: bidders-as-columns CSV was misread as single bidder.
        for gi, (tc, rc) in enumerate(groups):
            if tc < len(header):
                cell = (header[tc] or "").strip()
                # Skip generic headers like "Total", "Amount", "Bid"
                hn = norm_desc(cell)
                if cell and len(cell) > 2 and hn not in ("rate", "unit price", "price", "total", "amount", "extended", "bid"):
                    bidder_names[gi], bidder_locs[gi] = _split_name(cell)
        if all(bidder_names):
            name_conf = "high"
        elif any(bidder_names):
            name_conf = "medium"
        else:
            warnings.append("no bidder name row detected - please enter bidder names manually")

    for gi in range(n_bidders):
        if not bidder_names[gi]:
            bidder_names[gi] = f"Bidder {gi + 1}"

    # description/qty/unit columns: tabs commonly lead with an item-number
    # column, so detect from the header instead of assuming positions.
    tab_desc_col = _find_col(header, ("description", "work description",
                                      "scope", "scope of work",
                                      "description of work", "item description",
                                      "bid item", "work"))
    if tab_desc_col is None:
        for i, h in enumerate(header):
            hn = norm_desc(h)
            if "description" in hn or "scope" in hn:
                tab_desc_col = i
                break
    if tab_desc_col is None:
        tab_desc_col = 0
    tab_qty_col = _find_col(header, ("qty", "quantity", "qnty"))
    tab_unit_col = _find_col(header, ("unit", "uom", "u/m", "um"))
    # qty/unit must come before the first bidder column
    if tab_qty_col is not None and tab_qty_col >= first_bid_col:
        tab_qty_col = None
    if tab_unit_col is not None and tab_unit_col >= first_bid_col:
        tab_unit_col = None

    # item rows
    items = []
    submitted, subtotals = [None] * n_bidders, [None] * n_bidders
    taxes = [0.0] * n_bidders
    has_tax_rows = False
    tax_rate_detected = None
    for r in data[hdr_idx + 1:]:
        if not r or not any((c or "").strip() for c in r):
            continue
        if is_total_row(r, tab_desc_col):
            for gi, (tc, _) in enumerate(groups):
                v, st = parse_money(r[tc] if tc < len(r) else "")
                if st == "priced":
                    submitted[gi] = v
            continue
        if is_subtotal_row(r, tab_desc_col):
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
        desc = (r[tab_desc_col] if tab_desc_col < len(r) else "") or ""
        desc = desc.strip()
        if not desc or len(desc) < 2:
            continue
        qty_raw = ((r[tab_qty_col] if tab_qty_col is not None and tab_qty_col < len(r)
                    else "") or "").strip()
        unit_raw = ((r[tab_unit_col] if tab_unit_col is not None and tab_unit_col < len(r)
                     else "") or "").strip()
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
                              blocked=(basis_state == "unknown"),
                              total_states=total_states)
    if ranking["state"] == "blocked":
        warnings.append(
            "⛔ Cannot rank these bids: " + " ".join(basis_notes) +
            " Enter a visits-per-period frequency to see a PROVISIONAL comparison.")
    elif ranking["state"] == "too_close":
        warnings.append("⚠ " + ranking["note"])
    elif ranking["state"] == "needs_review":
        warnings.append("⛔ " + ranking["note"])

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

def _find_col(header, names):
    """Find first column whose normalized header matches one of names."""
    for i, h in enumerate(header):
        if norm_desc(h) in names:
            return i
    return None


def _rows_to_bidder(rows, filename):
    """rows: list of lists (strings). Find item/amount table -> one bidder's items.

    Detects the description column from the header (bid files commonly lead
    with an item-number column — description is NOT always column 0).
    Also extracts qty/unit when those columns exist."""
    # find header
    hdr = next((i for i, r in enumerate(rows) if is_header_row(r)), None)
    start = hdr + 1 if hdr is not None else 0
    header = rows[hdr] if hdr is not None else []
    # description column: prefer explicit description/scope headers; bare
    # "item"/"#" usually means the item NUMBER column, not the description.
    desc_col = None
    if header:
        desc_col = _find_col(header, ("description", "work description",
                                      "scope", "scope of work", "description of work",
                                      "item description", "bid item", "work"))
        if desc_col is None:
            for i, h in enumerate(header):
                hn = norm_desc(h)
                if "description" in hn or "scope" in hn:
                    desc_col = i
                    break
        if desc_col is None:
            # "item" alone is ambiguous — use it only if no item-number-ish
            # column exists elsewhere
            item_col = _find_col(header, ("item", "item no", "item #", "#", "no",
                                          "line", "line item", "ref"))
            desc_col = 0 if item_col in (None, 0) else 1
            if desc_col >= len(header):
                desc_col = 0
    if desc_col is None:
        desc_col = 0
    # qty / unit columns
    qty_col = _find_col(header, ("qty", "quantity", "qnty", "q ty")) if header else None
    unit_col = _find_col(header, ("unit", "uom", "u/m", "um")) if header else None
    # amount column: prefer TOTAL/AMOUNT, else last numeric-ish column
    amt_col = None
    if header:
        for i, h in enumerate(header):
            if norm_desc(h) in ("total", "amount", "extended", "extended total",
                                "bid", "price", "total price", "extended price",
                                "line total"):
                amt_col = i
                break
    skip_cols = {c for c in (desc_col, qty_col, unit_col) if c is not None}
    items = []
    stated_total = None  # bidder's own stated total, captured from total rows
    tax_total = 0.0  # accumulated tax rows (not line items, never summed)
    has_tax = False
    for r in rows[start:]:
        if not r or not any(str(c or "").strip() for c in r):
            continue
        if is_total_row(r, desc_col) or is_subtotal_row(r, desc_col) or is_tax_row(r):
            # capture the stated total before skipping
            if stated_total is None and is_total_row(r, desc_col):
                for i in range(len(r) - 1, -1, -1):
                    v, st = parse_money(r[i])
                    if st == "priced" and v:
                        stated_total = v
                        break
            if is_tax_row(r):
                # tax is jurisdiction math, not a bid line — capture separately
                for i in ([amt_col] if amt_col is not None else []) + \
                         [j for j in range(len(r) - 1, -1, -1) if j != amt_col]:
                    if i is None or i >= len(r):
                        continue
                    v, st = parse_money(r[i])
                    if st == "priced" and v:
                        tax_total += v
                        has_tax = True
                        break
            continue
        desc = (r[desc_col] if desc_col < len(r) else "") or ""
        desc = str(desc).strip()
        if not desc or len(desc) < 2:
            continue
        # qty / unit
        qty = None
        if qty_col is not None and qty_col < len(r):
            try:
                qraw = clean_num_text(str(r[qty_col] or "").strip())
                qty = float(qraw) if qraw else None
            except ValueError:
                qty = None
        unit = str(r[unit_col] if unit_col is not None and unit_col < len(r)
                    else "" or "").strip()
        # amount: chosen col or scan for last money-like cell (skip desc/qty/unit)
        order = [amt_col] if amt_col is not None else []
        order += [i for i in range(len(r) - 1, -1, -1)
                  if i != amt_col and i not in skip_cols]
        val, status, rate = None, "not_priced", None
        for i in order:
            if i is None or i >= len(r):
                continue
            v, st = parse_money(r[i])
            # "percent" is a real value (fee as %); "text" means keep scanning
            # for a money cell (e.g. item-number column holds "4")
            if st in ("priced", "included", "not_priced", "percent"):
                val, status = v, st
                break
        if status == "text":
            continue
        items.append({"description": desc, "qty": qty, "unit": unit,
                      "raw_amount": val, "raw_status": status})
    # bidder name from filename
    name = os.path.splitext(filename)[0].replace("_", " ").replace("-", " ")
    return {"name": name, "location": "", "file": filename,
            "items": items, "stated_total": stated_total,
            "tax_total": tax_total, "has_tax": has_tax}


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
    # try multi-bidder tab first (one tabulation with all bidders)
    data, errs = extract_tab_from_rows(best, None, filename)
    if (data and len(data.get("bidders", [])) >= 2 and data.get("items")
            and data.get("bidder_confidence") != "low"):
        data["_is_tab"] = True
        return data, []
    # fall back to single-bidder mode
    return _rows_to_bidder(best, filename), []


def extract_csv_file(path, filename):
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = [r for r in csv.reader(f)]
    if not rows:
        return None, ["empty csv"]
    # try multi-bidder tab first
    data, errs = extract_tab_from_rows(rows, None, filename)
    if (data and len(data.get("bidders", [])) >= 2 and data.get("items")
            and data.get("bidder_confidence") != "low"):
        data["_is_tab"] = True
        return data, []
    return _rows_to_bidder(rows, filename), []


def _pdf_prose_total(pages):
    """Find a stated lump-sum total in prose (no-table proposal letters).

    Returns the amount or None. Anchored to explicit total language
    ("lump sum", "total bid", "bid amount") so unrelated dollar figures
    never fabricate a total. Ambiguous (multiple distinct values) -> None.
    """
    import re
    full = "\n".join(pages)
    patterns = [
        r"lump\s*sum\s*[:\-]?\s*\$([\d,]+(?:\.\d{1,2})?)",
        r"total\s+(?:bid|proposal|contract)\s*(?:amount|price)?\s*[:\-]?\s*\$([\d,]+(?:\.\d{1,2})?)",
        r"\bbid\s+amount\s*[:\-]?\s*\$([\d,]+(?:\.\d{1,2})?)",
        r"propos(?:al|ed)\s+(?:price|amount)\s*[:\-]?\s*\$([\d,]+(?:\.\d{1,2})?)",
        r"contract\s+sum\s*[:\-]?\s*\$([\d,]+(?:\.\d{1,2})?)",
    ]
    found = []
    for pat in patterns:
        for m in re.finditer(pat, full, re.IGNORECASE):
            try:
                found.append(float(m.group(1).replace(",", "")))
            except ValueError:
                pass
    distinct = sorted(set(found))
    if len(distinct) == 1:
        return distinct[0]
    return None


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
        # lump-sum proposal letter: no tables — look for a stated total
        # in the prose so the bidder isn't silently dropped from a merge.
        pages = _pdf_full_text(path)
        total = _pdf_prose_total(pages)
        if total is not None:
            name = os.path.splitext(filename)[0].replace("_", " ").replace("-", " ")
            return {"name": name, "location": "", "file": filename,
                    "items": [], "stated_total": total,
                    "tax_total": 0.0, "has_tax": False,
                    "prose_total": True}, [
            f"no line-item table found in {filename} — lump-sum total "
            f"${total:,.2f} taken from proposal text; no line items to level"]
        return None, ["no tables found in PDF"]
    return _rows_to_bidder(rows, filename), []


def _apply_scope_comparability(items, bidders):
    """Round 4 fix (Sim 25): detect when bidders priced disjoint scopes.
    If any pair of bidders shares < 50% of priced items, mark those items
    comparable=False so the UI warns instead of ranking across scopes."""
    n_b = len(bidders)
    if n_b < 2 or not items:
        return []
    # per-bidder set of priced item indices
    priced_sets = []
    for bi in range(n_b):
        s = set(i for i, it in enumerate(items)
                if it["values"][bi]["status"] == "priced"
                and it["values"][bi]["amount"] not in (None, 0))
        priced_sets.append(s)
    notes = []
    for a in range(n_b):
        for b in range(a + 1, n_b):
            sa, sb = priced_sets[a], priced_sets[b]
            union = sa | sb
            if not union:
                continue
            overlap = len(sa & sb) / len(union)
            if overlap < 0.5:
                notes.append(
                    f"{bidders[a]['name']} vs {bidders[b]['name']}: only "
                    f"{overlap * 100:.0f}% scope overlap — likely different scopes")
    if notes:
        for it in items:
            it["comparable"] = False
    return notes


def merge_single_bids(bidders):
    """Merge per-bidder item lists into shared normalized items (fuzzy desc match)."""
    items = []  # {description, values:[per bidder {amount,status}]}
    for bi, b in enumerate(bidders):
        for it in b["items"]:
            # Round 4 fix: NEVER merge a deduct row with a non-deduct row —
            # doing so strips the deduct flag and defeats the substitution gate.
            it_deduct = _has_deduct_lang(it["description"])
            hit = next((m for m in items
                        if desc_match(m["description"], it["description"])
                        and _has_deduct_lang(m["description"]) == it_deduct), None)
            if hit is None:
                dl = (it["description"] or "").lower()
                mn, mx = parse_minmax(it["description"])
                hit = {"description": it["description"], "qty": it.get("qty"),
                       "unit": it.get("unit") or "",
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


@app.route("/w/<token>")
def waiver_sign_page(token):
    """Public waiver signing page — served by the SPA (client renders from /api/waiver/<token>)."""
    return send_from_directory(BASE, "index.html")


@app.route("/api/health")
def health():
    return jsonify(ok=True)


@app.route("/api/track", methods=["POST"])
def track_page_view():
    """Record a page view. No auth required. Bots are flagged, not blocked."""
    from auth_billing import get_db
    import hashlib, time
    body = request.get_json(force=True, silent=True) or {}
    path = (body.get("path") or "/")[:200]
    referrer = (body.get("referrer") or "")[:500]
    ua = request.headers.get("User-Agent", "")[:500]
    # Simple bot detection
    ua_lower = ua.lower()
    is_bot = any(b in ua_lower for b in ("bot", "crawler", "spider", "headless", "selenium", "phantom"))
    # Daily salted visitor hash (no raw IP stored)
    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "")
    day = time.strftime("%Y-%m-%d")
    visitor_hash = hashlib.sha256(f"{ip}|{ua}|{day}|tabulator-salt".encode()).hexdigest()[:16]
    if is_bot:
        return jsonify(ok=True, bot=True)
    db = get_db()
    try:
        db.execute(
            "INSERT INTO page_views (path, referrer, user_agent, visitor_hash, created_at) VALUES (?, ?, ?, ?, ?)",
            (path, referrer, ua, visitor_hash, int(time.time()))
        )
        db.commit()
    except Exception:
        pass
    finally:
        db.close()
    return jsonify(ok=True)


@app.route("/api/admin/tracking")
def admin_tracking():
    """Admin-only: page view stats. Requires admin auth."""
    from auth_billing import get_db, login_required, admin_required
    # Use the existing admin decorator pattern
    import functools
    from flask import g
    # Check admin via the existing mechanism
    auth_header = request.headers.get("Authorization", "")
    # Simplified: reuse the admin check from auth_billing
    try:
        from auth_billing import get_current_user
        user = get_current_user()
        if not user or not user.get("is_admin"):
            return jsonify(error="Admin required"), 403
    except Exception:
        return jsonify(error="Admin required"), 403
    import time
    now = int(time.time())
    db = get_db()
    try:
        def count_since(seconds):
            row = db.execute(
                "SELECT COUNT(*) as c, COUNT(DISTINCT visitor_hash) as u FROM page_views WHERE created_at > ?",
                (now - seconds,)
            ).fetchone()
            return {"views": row["c"], "visitors": row["u"]}
        stats = {
            "last_hour": count_since(3600),
            "today": count_since(86400),
            "last_7d": count_since(7 * 86400),
            "last_30d": count_since(30 * 86400),
            "all_time": count_since(now),
        }
        # Top pages (last 30d)
        top_pages = [
            {"path": r["path"], "views": r["c"]}
            for r in db.execute(
                "SELECT path, COUNT(*) as c FROM page_views WHERE created_at > ? GROUP BY path ORDER BY c DESC LIMIT 10",
                (now - 30 * 86400,)
            ).fetchall()
        ]
        # Top referrers (last 30d)
        top_refs = [
            {"referrer": r["referrer"] or "(direct)", "views": r["c"]}
            for r in db.execute(
                "SELECT referrer, COUNT(*) as c FROM page_views WHERE created_at > ? GROUP BY referrer ORDER BY c DESC LIMIT 10",
                (now - 30 * 86400,)
            ).fetchall()
        ]
        # Daily views (last 14 days) for a simple chart
        daily = []
        for i in range(13, -1, -1):
            day_start = now - (i + 1) * 86400
            day_end = now - i * 86400
            row = db.execute(
                "SELECT COUNT(*) as c, COUNT(DISTINCT visitor_hash) as u FROM page_views WHERE created_at > ? AND created_at <= ?",
                (day_start, day_end)
            ).fetchone()
            daily.append({
                "date": time.strftime("%m/%d", time.localtime(day_end)),
                "views": row["c"],
                "visitors": row["u"],
            })
        stats["top_pages"] = top_pages
        stats["top_referrers"] = top_refs
        stats["daily"] = daily
        return jsonify(stats)
    finally:
        db.close()


@app.route("/api/admin/tracking/reset", methods=["POST"])
def admin_tracking_reset():
    """Admin-only: delete all tracked page views (reset counters to zero)."""
    try:
        from auth_billing import get_current_user, get_db
        user = get_current_user()
        if not user or not user.get("is_admin"):
            return jsonify(error="Admin required"), 403
    except Exception:
        return jsonify(error="Admin required"), 403
    db = get_db()
    try:
        db.execute("DELETE FROM page_views")
        db.commit()
        return jsonify(ok=True)
    finally:
        db.close()


def _do_upload():
    """Core parsing logic. Returns a Flask response. Wrapped by upload() for gating."""
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
            path, filename = saved[0][0], saved[0][1]
            # Format-specific parsers first (cheaper and more precise than table extraction)
            pages = _pdf_full_text(path)
            data, errs = None, []
            if detect_verdantas_format(pages):
                data, errs = extract_verdantas_summary(path, filename)
            elif detect_mass_housing(pages):
                data, errs = extract_mass_housing(path, filename)
            elif detect_board_packet(pages):
                data, errs = extract_board_packet(path, filename)
            if data and len(data.get("bidders", [])) >= 2 and data.get("items"):
                return jsonify(data)
            if data and data.get("format") == "board_packet":
                # Board packet: informational only, never ranked
                return jsonify(data)
            # fall through to standard table extraction
            data, errs2 = extract_pdf_tab(path, filename)
            if data and len(data["bidders"]) >= 2 and data["items"]:
                return jsonify(data)
            # fall through to single-bidder mode. Tab-attempt warnings describe
            # a discarded parse — keep them only if single-bidder also fails.
            b, errs3 = extract_pdf_single(path, filename)
            bidders = [b] if b else []
            warnings = (errs + errs2 + errs3) if not b else errs3
        else:
            bidders, warnings = [], []
            for path, filename, ext in saved:
                if ext == ".pdf":
                    # try tab first
                    data, errs = extract_pdf_tab(path, filename)
                    if data and len(data["bidders"]) >= 2 and data["items"]:
                        return jsonify(data)
                    b, errs2 = extract_pdf_single(path, filename)
                    # tab-attempt warnings describe a discarded parse — keep
                    # them only if single-bidder extraction also fails.
                    warnings += errs2 if b else (errs + errs2)
                elif ext in (".xlsx", ".xls"):
                    b, errs2 = extract_excel(path, filename)
                    if isinstance(b, dict) and b.pop("_is_tab", False):
                        # multi-bidder tab found in this workbook
                        return jsonify(b)
                    warnings += errs2
                else:
                    b, errs2 = extract_csv_file(path, filename)
                    if isinstance(b, dict) and b.pop("_is_tab", False):
                        # multi-bidder tab found in this CSV
                        return jsonify(b)
                    warnings += errs2
                if b:
                    bidders.append(b)
            if not bidders:
                return jsonify(error="could not extract bid data",
                               warnings=warnings), 422

        items = merge_single_bids(bidders)
        # Round 4 fix (Sim 25): flag cross-scope comparisons
        scope_notes = _apply_scope_comparability(items, bidders)
        n_b = len(bidders)
        bidder_excluded = []
        for bi in range(n_b):
            priced = sum(1 for it in items
                         if it["values"][bi]["status"] == "priced"
                         and it["values"][bi]["amount"] not in (None, 0))
            bidder_excluded.append(priced == 0)
        # bids mode: verify against each file's own stated total when present.
        # Tax-aware: try tax-inclusive first, then tax-exclusive (same as tab mode).
        def _bids_calc(bi):
            return sum(it["values"][bi]["amount"] for it in items
                       if it.get("row_type", "amount") == "amount"
                       and not it.get("deduct") and not it.get("alternate")
                       and it["values"][bi]["status"] in ("priced", "included")
                       and it["values"][bi]["amount"] is not None)

        bids_total_states = []
        bids_submitted = []
        for bi, b in enumerate(bidders):
            calc = _bids_calc(bi)
            tax = b.get("tax_total", 0.0) or 0.0
            stated = b.get("stated_total")
            bids_submitted.append(stated)
            state = "unparseable"
            if stated is not None and calc > 0:
                tol = max(1.0, abs(stated) * 0.002)
                if abs((calc + tax) - stated) <= tol:
                    state = "verified"
                elif abs(calc - stated) <= tol:
                    state = "verified"
                else:
                    state = "mismatch"
            bids_total_states.append(state)
            if state == "mismatch":
                warnings.append(
                    f"{b['name']}: line items sum to ${calc:,.2f}"
                    + (f" + ${tax:,.2f} tax" if tax else "") +
                    f" but the file states ${stated:,.2f} — showing both, "
                    f"verify which is correct")
        basis_state, basis_notes = compute_basis_state(items)
        low_conf_counts = [sum(1 for it in items if it["values"][bi]["confidence"] == "low")
                           for bi in range(n_b)]
        # Round 4 fix: extraction quality from actual data, not hardcoded "good".
        # A single phantom bidder from a failed tab parse must never look confident.
        total_vals = sum(1 for it in items for bi in range(n_b)
                         if it["values"][bi]["status"] in ("priced", "included"))
        low_vals = sum(1 for it in items for bi in range(n_b)
                       if it["values"][bi]["confidence"] == "low")
        bids_extraction_quality = (
            "poor" if n_b < 2 or (total_vals and low_vals / total_vals > 0.4)
            else "good")
        ranking = compute_ranking(items, bidder_excluded,
                                  [b["name"] for b in bidders],
                                  bids_extraction_quality, low_conf_counts,
                                  blocked=(basis_state == "unknown"),
                                  total_states=bids_total_states)
        warnings_b = warnings + [
            "single-bid files: bidder names taken from filenames - please confirm"]
        if ranking["state"] == "blocked":
            warnings_b.append("⛔ Cannot rank these bids: " + " ".join(basis_notes))
        elif ranking["state"] == "too_close":
            warnings_b.append("⚠ " + ranking["note"])
        elif ranking["state"] == "needs_review":
            warnings_b.append("⛔ " + ranking["note"])
        if scope_notes:
            warnings_b.append("⛔ Cross-scope warning: " + "; ".join(scope_notes))
        # bids-mode taxes: captured per-file, never mixed into line items
        bids_taxes = [b.get("tax_total", 0.0) or 0.0 for b in bidders]
        bids_has_tax = any(b.get("has_tax") for b in bidders)
        return jsonify({
            "project_name": "Bid package",
            "mode": "bids",
            "bidders": [{"name": b["name"], "location": b.get("location", ""),
                         "file": b["file"], "excluded": ex, "total_state": ts}
                        for b, ex, ts in zip(bidders, bidder_excluded,
                                             bids_total_states)],
            "items": items,
            "submitted_totals": bids_submitted,
            "subtotals": [None] * len(bidders),
            "taxes": bids_taxes,
            "has_tax_rows": bids_has_tax,
            "tax_rate_detected": None,
            "tax_basis_detected": [None] * len(bidders),
            "has_substitution_alternates": (
                any(it.get("deduct") for it in items)
                and any(it.get("alternate") for it in items)),
            "basis_state": basis_state,
            "basis_notes": basis_notes,
            "ranking": ranking,
            "extraction_quality": bids_extraction_quality,
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


@app.route("/api/upload", methods=["POST"])
def upload():
    """Gated wrapper: enforces free-tier + subscription before parsing."""
    user = auth_billing.current_user()
    allowed, reason = auth_billing.can_upload(user)
    if not allowed:
        if reason == "needs_account":
            return jsonify(error="free_limit",
                           message="Your free tabulation is used. Create a free account to save it — "
                                   "then subscribe for $129/mo for unlimited tabulations."), 402
        return jsonify(error="subscription_required",
                       message="You've used your free tabulation. Subscribe for $129/mo "
                               "for unlimited bid tabulations."), 402
    resp = _do_upload()
    # On success, persist the tabulation and mark free use.
    try:
        status = resp.status_code if hasattr(resp, "status_code") else 200
        if status == 200:
            data = resp.get_json() if hasattr(resp, "get_json") else None
            if data and not data.get("error"):
                project = data.get("project_name") or "Bid package"
                uid = user["id"] if user else None
                tab_id = auth_billing.save_tabulation(uid, project, data)
                # attach tabulation id to response
                data["tabulation_id"] = tab_id
                data["saved"] = bool(uid)
                data["account_status"] = reason
                resp = jsonify(data)
                if not user:
                    session["free_used"] = True
    except Exception:
        pass  # never break parsing because of persistence
    return resp


# ---------------------------------------------------------------- Get-Paid suite
#
# Progress-billing + lien-waiver tracking for GCs, billed as an add-on service
# ($99/mo). Every endpoint here (except the two PUBLIC waiver endpoints) is
# gated by getpaid_required: logged in AND getpaid service active, else 402
# {error:"subscription_required"}.
#
# LEGAL LINE (load-bearing): these endpoints store whatever waiver_kind /
# waiver_scope the USER selects. Nothing here suggests which waiver type to
# use, sets defaults for them, validates them, or interprets state law.
# Keep it that way.
#
# company_id = users.id: this app has no companies table; one account = one
# company. If a companies table ever appears, swap g.user["id"] for the
# user's company id here.

def _gp_now():
    return datetime.now(timezone.utc).isoformat()


def _gp_project(company_id, pid):
    db = get_db()
    try:
        row = db.execute(
            "SELECT id, company_id, name, client_name, created_at FROM getpaid_projects"
            " WHERE id=? AND company_id=?", (pid, company_id)).fetchone()
        return dict(row) if row else None
    finally:
        db.close()


def _gp_draw(company_id, did):
    db = get_db()
    try:
        row = db.execute(
            "SELECT id, company_id, project_id, draw_number, period_start, period_end,"
            " status, retainage_pct, notes, created_at FROM pay_draws"
            " WHERE id=? AND company_id=?", (did, company_id)).fetchone()
        return dict(row) if row else None
    finally:
        db.close()


def _gp_draw_lines(did):
    db = get_db()
    try:
        rows = db.execute(
            "SELECT id, draw_id, description, scheduled_value, pct_complete, created_at"
            " FROM pay_draw_lines WHERE draw_id=? ORDER BY id", (did,)).fetchall()
        return [dict(r) for r in rows]
    finally:
        db.close()


def _gp_draw_waivers(did):
    db = get_db()
    try:
        rows = db.execute(
            "SELECT id, company_id, draw_id, sub_name, sub_email, waiver_kind,"
            " waiver_scope, amount, status, token, signed_at, signed_name, created_at"
            " FROM lien_waivers WHERE draw_id=? ORDER BY id", (did,)).fetchall()
        return [dict(r) for r in rows]
    finally:
        db.close()


def _gp_delete_draw(did):
    """Delete a draw and its lines + waivers. Caller must own the draw."""
    db = get_db()
    try:
        db.execute("DELETE FROM pay_draw_lines WHERE draw_id=?", (did,))
        db.execute("DELETE FROM lien_waivers WHERE draw_id=?", (did,))
        db.execute("DELETE FROM pay_draws WHERE id=?", (did,))
        db.commit()
    finally:
        db.close()


def _num(v, name):
    try:
        return float(v)
    except (TypeError, ValueError):
        raise ValueError(f"{name} must be a number")


def register_getpaid_routes(app):
    # ------------------------------------------------------------ projects

    @app.get("/api/getpaid/projects")
    @getpaid_required
    def gp_projects_list():
        db = get_db()
        try:
            rows = db.execute(
                "SELECT id, company_id, name, client_name, created_at"
                " FROM getpaid_projects WHERE company_id=? ORDER BY id DESC",
                (g.user["id"],)).fetchall()
            return jsonify(projects=[dict(r) for r in rows])
        finally:
            db.close()

    @app.post("/api/getpaid/projects")
    @getpaid_required
    def gp_projects_create():
        body = request.get_json(force=True, silent=True) or {}
        name = (body.get("name") or "").strip()
        if not name:
            return jsonify(error="name is required"), 400
        db = get_db()
        try:
            cur = db.execute(
                "INSERT INTO getpaid_projects (company_id, name, client_name, created_at)"
                " VALUES (?,?,?,?)",
                (g.user["id"], name, (body.get("client_name") or "").strip(),
                 _gp_now()))
            db.commit()
            pid = cur.lastrowid
        finally:
            db.close()
        return jsonify(project=_gp_project(g.user["id"], pid)), 201

    @app.get("/api/getpaid/projects/<int:pid>")
    @getpaid_required
    def gp_project_get(pid):
        proj = _gp_project(g.user["id"], pid)
        if not proj:
            return jsonify(error="Not found."), 404
        db = get_db()
        try:
            draws = db.execute(
                "SELECT id, company_id, project_id, draw_number, period_start,"
                " period_end, status, retainage_pct, notes, created_at"
                " FROM pay_draws WHERE project_id=? AND company_id=?"
                " ORDER BY draw_number, id",
                (pid, g.user["id"])).fetchall()
            return jsonify(project=proj, draws=[dict(r) for r in draws])
        finally:
            db.close()

    @app.put("/api/getpaid/projects/<int:pid>")
    @getpaid_required
    def gp_project_put(pid):
        if not _gp_project(g.user["id"], pid):
            return jsonify(error="Not found."), 404
        body = request.get_json(force=True, silent=True) or {}
        name = (body.get("name") or "").strip()
        if not name:
            return jsonify(error="name is required"), 400
        db = get_db()
        try:
            db.execute(
                "UPDATE getpaid_projects SET name=?, client_name=? WHERE id=? AND company_id=?",
                (name, (body.get("client_name") or "").strip(), pid, g.user["id"]))
            db.commit()
        finally:
            db.close()
        return jsonify(project=_gp_project(g.user["id"], pid))

    @app.delete("/api/getpaid/projects/<int:pid>")
    @getpaid_required
    def gp_project_delete(pid):
        if not _gp_project(g.user["id"], pid):
            return jsonify(error="Not found."), 404
        db = get_db()
        try:
            draw_ids = [r["id"] for r in db.execute(
                "SELECT id FROM pay_draws WHERE project_id=? AND company_id=?",
                (pid, g.user["id"])).fetchall()]
            for did in draw_ids:
                db.execute("DELETE FROM pay_draw_lines WHERE draw_id=?", (did,))
                db.execute("DELETE FROM lien_waivers WHERE draw_id=?", (did,))
            db.execute("DELETE FROM pay_draws WHERE project_id=? AND company_id=?",
                       (pid, g.user["id"]))
            db.execute("DELETE FROM getpaid_projects WHERE id=? AND company_id=?",
                       (pid, g.user["id"]))
            db.commit()
        finally:
            db.close()
        return jsonify(ok=True)

    # ------------------------------------------------------------ draws

    @app.get("/api/getpaid/projects/<int:pid>/draws")
    @getpaid_required
    def gp_draws_list(pid):
        if not _gp_project(g.user["id"], pid):
            return jsonify(error="Not found."), 404
        db = get_db()
        try:
            rows = db.execute(
                "SELECT id, company_id, project_id, draw_number, period_start,"
                " period_end, status, retainage_pct, notes, created_at"
                " FROM pay_draws WHERE project_id=? AND company_id=?"
                " ORDER BY draw_number, id",
                (pid, g.user["id"])).fetchall()
            return jsonify(draws=[dict(r) for r in rows])
        finally:
            db.close()

    @app.post("/api/getpaid/projects/<int:pid>/draws")
    @getpaid_required
    def gp_draws_create(pid):
        if not _gp_project(g.user["id"], pid):
            return jsonify(error="Not found."), 404
        body = request.get_json(force=True, silent=True) or {}
        try:
            draw_number = int(body.get("draw_number", 1))
            retainage_pct = _num(body.get("retainage_pct", 10), "retainage_pct")
        except (TypeError, ValueError) as e:
            return jsonify(error=str(e) if isinstance(e, ValueError) else
                           "draw_number must be an integer"), 400
        db = get_db()
        try:
            cur = db.execute(
                """INSERT INTO pay_draws
                   (company_id, project_id, draw_number, period_start, period_end,
                    status, retainage_pct, notes, created_at)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                (g.user["id"], pid, draw_number,
                 body.get("period_start") or None, body.get("period_end") or None,
                 body.get("status") or "draft", retainage_pct,
                 body.get("notes") or "", _gp_now()))
            db.commit()
            did = cur.lastrowid
        finally:
            db.close()
        return jsonify(draw=_gp_draw(g.user["id"], did)), 201

    @app.get("/api/getpaid/draws/<int:did>")
    @getpaid_required
    def gp_draw_get(did):
        draw = _gp_draw(g.user["id"], did)
        if not draw:
            return jsonify(error="Not found."), 404
        lines = _gp_draw_lines(did)
        waivers = _gp_draw_waivers(did)
        scheduled_total = sum((l.get("scheduled_value") or 0) for l in lines)
        earned_to_date = sum((l.get("scheduled_value") or 0) *
                             (l.get("pct_complete") or 0) / 100.0 for l in lines)
        retained_to_date = earned_to_date * (draw.get("retainage_pct") or 0) / 100.0
        due_this_draw = earned_to_date - retained_to_date
        return jsonify(draw=draw, lines=lines, waivers=waivers, totals={
            "scheduled_total": round(scheduled_total, 2),
            "earned_to_date": round(earned_to_date, 2),
            "retained_to_date": round(retained_to_date, 2),
            "due_this_draw": round(due_this_draw, 2),
        })

    @app.put("/api/getpaid/draws/<int:did>")
    @getpaid_required
    def gp_draw_put(did):
        draw = _gp_draw(g.user["id"], did)
        if not draw:
            return jsonify(error="Not found."), 404
        body = request.get_json(force=True, silent=True) or {}
        fields = {}
        if "draw_number" in body:
            try:
                fields["draw_number"] = int(body["draw_number"])
            except (TypeError, ValueError):
                return jsonify(error="draw_number must be an integer"), 400
        for k in ("period_start", "period_end", "status", "notes"):
            if k in body:
                fields[k] = (body[k] or "").strip() if isinstance(body[k], str) else body[k]
        if "retainage_pct" in body:
            try:
                fields["retainage_pct"] = _num(body["retainage_pct"], "retainage_pct")
            except ValueError as e:
                return jsonify(error=str(e)), 400
        if not fields:
            return jsonify(draw=draw)
        sets = ", ".join(f"{k}=?" for k in fields)
        db = get_db()
        try:
            db.execute(f"UPDATE pay_draws SET {sets} WHERE id=? AND company_id=?",
                       (*fields.values(), did, g.user["id"]))
            db.commit()
        finally:
            db.close()
        return jsonify(draw=_gp_draw(g.user["id"], did))

    # ------------------------------------------------------------ draw lines

    @app.post("/api/getpaid/draws/<int:did>/lines")
    @getpaid_required
    def gp_lines_create(did):
        if not _gp_draw(g.user["id"], did):
            return jsonify(error="Not found."), 404
        body = request.get_json(force=True, silent=True) or {}
        description = (body.get("description") or "").strip()
        if not description:
            return jsonify(error="description is required"), 400
        try:
            scheduled_value = _num(body.get("scheduled_value", 0), "scheduled_value")
            pct_complete = _num(body.get("pct_complete", 0), "pct_complete")
        except ValueError as e:
            return jsonify(error=str(e)), 400
        db = get_db()
        try:
            cur = db.execute(
                "INSERT INTO pay_draw_lines (draw_id, description, scheduled_value,"
                " pct_complete, created_at) VALUES (?,?,?,?,?)",
                (did, description, scheduled_value, pct_complete, _gp_now()))
            db.commit()
            lid = cur.lastrowid
            row = db.execute(
                "SELECT id, draw_id, description, scheduled_value, pct_complete, created_at"
                " FROM pay_draw_lines WHERE id=?", (lid,)).fetchone()
        finally:
            db.close()
        return jsonify(line=dict(row)), 201

    @app.put("/api/getpaid/lines/<int:lid>")
    @getpaid_required
    def gp_line_put(lid):
        db = get_db()
        try:
            row = db.execute(
                "SELECT l.id, l.description, l.scheduled_value, l.pct_complete,"
                " l.created_at, d.company_id"
                " FROM pay_draw_lines l JOIN pay_draws d ON d.id=l.draw_id"
                " WHERE l.id=?", (lid,)).fetchone()
            if not row or row["company_id"] != g.user["id"]:
                return jsonify(error="Not found."), 404
            body = request.get_json(force=True, silent=True) or {}
            fields = {}
            if "description" in body:
                d = (body["description"] or "").strip()
                if not d:
                    return jsonify(error="description cannot be empty"), 400
                fields["description"] = d
            for k in ("scheduled_value", "pct_complete"):
                if k in body:
                    try:
                        fields[k] = _num(body[k], k)
                    except ValueError as e:
                        return jsonify(error=str(e)), 400
            if fields:
                sets = ", ".join(f"{k}=?" for k in fields)
                db.execute(f"UPDATE pay_draw_lines SET {sets} WHERE id=?",
                           (*fields.values(), lid))
                db.commit()
            row = db.execute(
                "SELECT id, draw_id, description, scheduled_value, pct_complete, created_at"
                " FROM pay_draw_lines WHERE id=?", (lid,)).fetchone()
            return jsonify(line=dict(row))
        finally:
            db.close()

    @app.delete("/api/getpaid/lines/<int:lid>")
    @getpaid_required
    def gp_line_delete(lid):
        db = get_db()
        try:
            row = db.execute(
                "SELECT l.id, d.company_id FROM pay_draw_lines l"
                " JOIN pay_draws d ON d.id=l.draw_id WHERE l.id=?",
                (lid,)).fetchone()
            if not row or row["company_id"] != g.user["id"]:
                return jsonify(error="Not found."), 404
            db.execute("DELETE FROM pay_draw_lines WHERE id=?", (lid,))
            db.commit()
        finally:
            db.close()
        return jsonify(ok=True)

    # ------------------------------------------------------------ lien waivers

    @app.post("/api/getpaid/draws/<int:did>/waivers")
    @getpaid_required
    def gp_waivers_create(did):
        if not _gp_draw(g.user["id"], did):
            return jsonify(error="Not found."), 404
        body = request.get_json(force=True, silent=True) or {}
        sub_name = (body.get("sub_name") or "").strip()
        if not sub_name:
            return jsonify(error="sub_name is required"), 400
        try:
            amount = _num(body.get("amount", 0), "amount")
        except ValueError as e:
            return jsonify(error=str(e)), 400
        # NOTE: waiver_kind / waiver_scope are stored exactly as the user
        # selects. No defaults, no suggestions, no legal interpretation.
        db = get_db()
        try:
            cur = db.execute(
                """INSERT INTO lien_waivers
                   (company_id, draw_id, sub_name, sub_email, waiver_kind,
                    waiver_scope, amount, status, created_at)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                (g.user["id"], did, sub_name, (body.get("sub_email") or "").strip(),
                 (body.get("waiver_kind") or "").strip(),
                 (body.get("waiver_scope") or "").strip(),
                 amount, "pending", _gp_now()))
            db.commit()
            wid = cur.lastrowid
            row = db.execute(
                "SELECT id, company_id, draw_id, sub_name, sub_email, waiver_kind,"
                " waiver_scope, amount, status, token, signed_at, signed_name, created_at"
                " FROM lien_waivers WHERE id=?", (wid,)).fetchone()
        finally:
            db.close()
        return jsonify(waiver=dict(row)), 201

    @app.put("/api/getpaid/waivers/<int:wid>")
    @getpaid_required
    def gp_waiver_put(wid):
        db = get_db()
        try:
            row = db.execute(
                "SELECT id FROM lien_waivers WHERE id=? AND company_id=?",
                (wid, g.user["id"])).fetchone()
            if not row:
                return jsonify(error="Not found."), 404
            body = request.get_json(force=True, silent=True) or {}
            fields = {}
            for k in ("sub_name", "sub_email", "waiver_kind", "waiver_scope"):
                # stored as the user provides; never interpreted here.
                if k in body:
                    fields[k] = (body[k] or "").strip() if isinstance(body[k], str) else body[k]
            if "sub_name" in fields and not fields["sub_name"]:
                return jsonify(error="sub_name cannot be empty"), 400
            for k in ("amount",):
                if k in body:
                    try:
                        fields[k] = _num(body[k], k)
                    except ValueError as e:
                        return jsonify(error=str(e)), 400
            if "status" in body and isinstance(body["status"], str):
                fields["status"] = body["status"].strip() or "pending"
            if fields:
                sets = ", ".join(f"{k}=?" for k in fields)
                db.execute(f"UPDATE lien_waivers SET {sets} WHERE id=?",
                           (*fields.values(), wid))
                db.commit()
            row = db.execute(
                "SELECT id, company_id, draw_id, sub_name, sub_email, waiver_kind,"
                " waiver_scope, amount, status, token, signed_at, signed_name, created_at"
                " FROM lien_waivers WHERE id=?", (wid,)).fetchone()
            return jsonify(waiver=dict(row))
        finally:
            db.close()

    @app.delete("/api/getpaid/waivers/<int:wid>")
    @getpaid_required
    def gp_waiver_delete(wid):
        db = get_db()
        try:
            row = db.execute(
                "SELECT id FROM lien_waivers WHERE id=? AND company_id=?",
                (wid, g.user["id"])).fetchone()
            if not row:
                return jsonify(error="Not found."), 404
            db.execute("DELETE FROM lien_waivers WHERE id=?", (wid,))
            db.commit()
        finally:
            db.close()
        return jsonify(ok=True)

    @app.post("/api/getpaid/waivers/<int:wid>/send-link")
    @getpaid_required
    def gp_waiver_send_link(wid):
        db = get_db()
        try:
            row = db.execute(
                "SELECT id, token FROM lien_waivers WHERE id=? AND company_id=?",
                (wid, g.user["id"])).fetchone()
            if not row:
                return jsonify(error="Not found."), 404
            token = row["token"]
            if not token:
                # generate a secure random token, retry on the rare collision
                for _ in range(5):
                    token = secrets.token_urlsafe(24)
                    try:
                        db.execute("UPDATE lien_waivers SET token=? WHERE id=?",
                                   (token, wid))
                        db.commit()
                        break
                    except Exception:
                        db.rollback()
                        token = None
                if not token:
                    return jsonify(error="Could not generate link, try again."), 500
            return jsonify(url=f"/w/{token}")
        finally:
            db.close()

    # ------------------------------------------- PUBLIC waiver endpoints

    @app.get("/api/waiver/<token>")
    def waiver_public_get(token):
        """Public: show the waiver details to the signer. No auth."""
        db = get_db()
        try:
            row = db.execute(
                """SELECT w.sub_name, w.amount, w.waiver_kind, w.waiver_scope,
                          w.status, w.signed_at, w.token,
                          p.name AS project_name, d.draw_number
                   FROM lien_waivers w
                   JOIN pay_draws d ON d.id=w.draw_id
                   JOIN getpaid_projects p ON p.id=d.project_id
                   WHERE w.token=?""", (token,)).fetchone()
        finally:
            db.close()
        if not row:
            return jsonify(error="Link not found or expired."), 404
        return jsonify(sub_name=row["sub_name"], amount=row["amount"],
                       waiver_kind=row["waiver_kind"],
                       waiver_scope=row["waiver_scope"],
                       project_name=row["project_name"],
                       draw_number=row["draw_number"],
                       status=row["status"])

    @app.post("/api/waiver/<token>/sign")
    def waiver_public_sign(token):
        """Public: record the sub's signature. Single-use."""
        body = request.get_json(force=True, silent=True) or {}
        signed_name = (body.get("signed_name") or "").strip()
        if not signed_name:
            return jsonify(error="signed_name is required"), 400
        db = get_db()
        try:
            row = db.execute(
                "SELECT id, status FROM lien_waivers WHERE token=?",
                (token,)).fetchone()
            if not row:
                return jsonify(error="Link not found or expired."), 404
            if row["status"] == "signed":
                return jsonify(error="already_signed"), 409
            signed_at = _gp_now()
            db.execute("UPDATE lien_waivers SET status='signed', signed_at=?, signed_name=?"
                       " WHERE id=?", (signed_at, signed_name, row["id"]))
            db.commit()
        finally:
            db.close()
        return jsonify(ok=True, signed_at=signed_at, signed_name=signed_name)


register_getpaid_routes(app)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
