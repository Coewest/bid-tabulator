#!/usr/bin/env python3
"""Bid Tabulator — accounts, tabulation history, usage tracking, Stripe billing.

Tables (SQLite):
  users(id, email UNIQUE, password_hash, created_at)
  tabulations(id, user_id NULL, project_name, data_json, created_at, claimed)
  subscriptions(user_id UNIQUE, stripe_customer_id, stripe_subscription_id,
                status, current_period_end, updated_at)

Flow:
  - Anonymous: first /api/upload works, tabulation saved with user_id=NULL.
    Response includes tabulation_id; frontend prompts "create account to save".
  - Signup: POST /api/signup links any unclaimed tabulation ids passed in.
  - Logged in, no subscription, tabulations >= 1: /api/upload -> 402 needs_subscription.
  - Stripe Checkout -> webhook marks subscription active.
  - Active subscription: unlimited uploads.
"""
import hashlib
import hmac
import json
import os
import secrets
import sqlite3
import time
import urllib.error
import urllib.request
from html import escape as html_escape
from datetime import datetime, timezone
from functools import wraps

from flask import g, jsonify, request, session

try:
    import stripe
except ImportError:
    stripe = None

DB_PATH = os.environ.get("TABULATOR_DB", os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "tabulator.db"))

STRIPE_SECRET_KEY = os.environ.get("STRIPE_SECRET_KEY", "")
STRIPE_PUBLISHABLE_KEY = os.environ.get("STRIPE_PUBLISHABLE_KEY", "")
STRIPE_WEBHOOK_SECRET = os.environ.get("STRIPE_WEBHOOK_SECRET", "")
STRIPE_PRICE_ID = os.environ.get("STRIPE_PRICE_ID", "")  # $129/mo recurring price
STRIPE_PRICE_ID_ANNUAL = os.environ.get("STRIPE_PRICE_ID_ANNUAL", "")  # $1,290/yr recurring price
# Get-Paid suite ($99/mo, $990/yr). Env vars override the baked-in fallbacks.
STRIPE_PRICE_ID_GETPAID = os.environ.get("STRIPE_PRICE_ID_GETPAID",
                                         "price_1UPN87GjNNTRCkRobuKW4BCa")
STRIPE_PRICE_ID_GETPAID_ANNUAL = os.environ.get("STRIPE_PRICE_ID_GETPAID_ANNUAL",
                                                "price_1UPNZrGjNNTRCkRol3G98ZIF")
# GC Suite bundle: Tabulator + Get-Paid ($179/mo, $1,750/yr). Env vars override fallbacks.
STRIPE_PRICE_ID_SUITE = os.environ.get("STRIPE_PRICE_ID_SUITE",
                                       "price_1UPNZxGjNNTRCkRolvnQypky")
STRIPE_PRICE_ID_SUITE_ANNUAL = os.environ.get("STRIPE_PRICE_ID_SUITE_ANNUAL",
                                              "price_1UPNZyGjNNTRCkRo51lP9ZEd")
APP_URL = os.environ.get("APP_URL", "https://bid-tabulator-production.up.railway.app").rstrip("/")

FREE_TABULATIONS = 1


def stripe_configured():
    return bool(stripe and STRIPE_SECRET_KEY and STRIPE_PRICE_ID)


def getpaid_configured():
    return bool(stripe and STRIPE_SECRET_KEY and STRIPE_PRICE_ID_GETPAID)


def suite_configured():
    return bool(stripe and STRIPE_SECRET_KEY and STRIPE_PRICE_ID_SUITE)


def get_db():
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    return db


def _add_column(db, table, col, ddl):
    """Idempotent ADD COLUMN. Gunicorn runs 2 workers and both call
    init_db() at boot, so check-then-ALTER races: the loser must ignore
    the duplicate-column error instead of crashing the worker."""
    cols = [r[1] for r in db.execute(f"PRAGMA table_info({table})")]
    if col in cols:
        return
    try:
        db.execute(f"ALTER TABLE {table} ADD COLUMN {ddl}")
    except sqlite3.OperationalError as e:
        if "duplicate column" not in str(e).lower():
            raise


def init_db():
    db = get_db()
    db.executescript("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        created_at INTEGER NOT NULL
    );
    CREATE TABLE IF NOT EXISTS tabulations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NULL REFERENCES users(id),
        project_name TEXT NOT NULL DEFAULT 'Bid package',
        data_json TEXT NOT NULL,
        created_at INTEGER NOT NULL,
        claimed INTEGER NOT NULL DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS subscriptions (
        user_id INTEGER PRIMARY KEY REFERENCES users(id),
        stripe_customer_id TEXT,
        stripe_subscription_id TEXT,
        status TEXT NOT NULL DEFAULT 'none',
        current_period_end INTEGER,
        updated_at INTEGER NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_tab_user ON tabulations(user_id);
    CREATE TABLE IF NOT EXISTS page_views (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        path TEXT NOT NULL,
        referrer TEXT,
        user_agent TEXT,
        visitor_hash TEXT NOT NULL,
        created_at INTEGER NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_pv_created ON page_views(created_at);
    CREATE INDEX IF NOT EXISTS idx_pv_visitor ON page_views(visitor_hash);
    -- Get-Paid suite tables (2026-10-11). company_id = users.id: this app has
    -- no companies table; one account = one company.
    CREATE TABLE IF NOT EXISTS getpaid_projects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        client_name TEXT NOT NULL DEFAULT '',
        created_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_gp_projects_company ON getpaid_projects(company_id);
    CREATE TABLE IF NOT EXISTS pay_draws (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        project_id INTEGER NOT NULL REFERENCES getpaid_projects(id),
        draw_number INTEGER NOT NULL DEFAULT 1,
        period_start TEXT,
        period_end TEXT,
        status TEXT NOT NULL DEFAULT 'draft',
        retainage_pct REAL NOT NULL DEFAULT 10,
        notes TEXT NOT NULL DEFAULT '',
        created_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_pay_draws_project ON pay_draws(project_id);
    CREATE INDEX IF NOT EXISTS idx_pay_draws_company ON pay_draws(company_id);
    CREATE TABLE IF NOT EXISTS pay_draw_lines (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        draw_id INTEGER NOT NULL REFERENCES pay_draws(id),
        description TEXT NOT NULL,
        scheduled_value REAL NOT NULL DEFAULT 0,
        pct_complete REAL NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_pay_draw_lines_draw ON pay_draw_lines(draw_id);
    CREATE TABLE IF NOT EXISTS lien_waivers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        draw_id INTEGER NOT NULL REFERENCES pay_draws(id),
        sub_name TEXT NOT NULL,
        sub_email TEXT NOT NULL DEFAULT '',
        waiver_kind TEXT NOT NULL DEFAULT '',
        waiver_scope TEXT NOT NULL DEFAULT '',
        amount REAL NOT NULL DEFAULT 0,
        status TEXT NOT NULL DEFAULT 'pending',
        token TEXT UNIQUE,
        signed_at TEXT,
        signed_name TEXT NOT NULL DEFAULT '',
        created_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_lien_waivers_draw ON lien_waivers(draw_id);
    CREATE INDEX IF NOT EXISTS idx_lien_waivers_company ON lien_waivers(company_id);
    CREATE INDEX IF NOT EXISTS idx_lien_waivers_token ON lien_waivers(token);
    CREATE TABLE IF NOT EXISTS company_services (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        service_key TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'active',
        stripe_subscription_id TEXT,
        stripe_price_id TEXT,
        current_period_end TEXT,
        created_at TEXT NOT NULL,
        UNIQUE (company_id, service_key)
    );
    CREATE INDEX IF NOT EXISTS idx_company_services_company ON company_services(company_id);
    """)
    # Migration: Get-Paid lien_waivers signed_name (DBs created before 2026-10-11)
    _add_column(db, "lien_waivers", "signed_name", "signed_name TEXT NOT NULL DEFAULT ''")
    # Migration: company profile fields
    for col in ("company_name", "contact_name", "phone"):
        _add_column(db, "users", col, f"{col} TEXT NOT NULL DEFAULT ''")
    # Migration: subscription billing interval (month/year)
    _add_column(db, "subscriptions", "billing_interval", "billing_interval TEXT NOT NULL DEFAULT 'month'")
    # Migration: email-first save (2026-10-12): anonymous tabulations carry a
    # random claim token; emailed set-password links use one-time tokens.
    _add_column(db, "tabulations", "claim_token", "claim_token TEXT NOT NULL DEFAULT ''")
    db.execute(
        "CREATE TABLE IF NOT EXISTS password_tokens ("
        " id INTEGER PRIMARY KEY AUTOINCREMENT,"
        " user_id INTEGER NOT NULL,"
        " token_hash TEXT NOT NULL,"
        " purpose TEXT NOT NULL DEFAULT 'set_password',"
        " expires_at INTEGER NOT NULL,"
        " used_at INTEGER NOT NULL DEFAULT 0,"
        " created_at INTEGER NOT NULL)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_password_tokens_hash ON password_tokens(token_hash)")
    db.commit()
    db.close()


def hash_password(password: str) -> str:
    salt = secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 200_000)
    return f"pbkdf2$200000${salt}${h.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        _, iters, salt, hexh = stored.split("$")
        h = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), int(iters))
        return secrets.compare_digest(h.hex(), hexh)
    except Exception:
        return False


def current_user():
    uid = session.get("user_id")
    if not uid:
        return None
    db = get_db()
    try:
        row = db.execute(
            "SELECT id, email, created_at, company_name, contact_name, phone"
            " FROM users WHERE id=?", (uid,)).fetchone()
        return dict(row) if row else None
    finally:
        db.close()


ADMIN_EMAIL = "coe@clearscopebid.com"


def is_admin(user):
    return bool(user) and user.get("email") == ADMIN_EMAIL


def admin_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        u = current_user()
        if not u or not is_admin(u):
            return jsonify(error="admin required"), 403
        g.user = u
        return fn(*a, **kw)
    return wrapper


def login_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        u = current_user()
        if not u:
            return jsonify(error="login required"), 401
        g.user = u
        return fn(*a, **kw)
    return wrapper


def getpaid_required(fn):
    """Get-Paid suite gate: logged in AND the getpaid service active.
    A GC Suite subscription also unlocks Get-Paid.
    Tabulator tabulations use the existing can_upload paywall, unchanged."""
    @wraps(fn)
    def wrapper(*a, **kw):
        u = current_user()
        if not u:
            return jsonify(error="login required"), 401
        g.user = u
        if not (service_active(u["id"], "getpaid") or service_active(u["id"], "suite")):
            return jsonify(error="subscription_required"), 402
        return fn(*a, **kw)
    return wrapper


def subscription_status(user_id: int) -> str:
    db = get_db()
    try:
        row = db.execute(
            "SELECT status, current_period_end FROM subscriptions WHERE user_id=?",
            (user_id,)).fetchone()
        if not row:
            return "none"
        status = row["status"]
        # treat past_due generously until period end
        if status in ("active", "trialing", "past_due"):
            return "active"
        return status
    finally:
        db.close()


def tabulation_count(user_id: int) -> int:
    db = get_db()
    try:
        row = db.execute(
            "SELECT COUNT(*) c FROM tabulations WHERE user_id=?", (user_id,)).fetchone()
        return row["c"]
    finally:
        db.close()


def _utcnow():
    return datetime.now(timezone.utc).isoformat()


def service_active(company_id: int, service_key: str) -> bool:
    """True when the company has an active row for the given service key
    in company_services. company_id = users.id (one account = one company)."""
    db = get_db()
    try:
        row = db.execute(
            "SELECT status FROM company_services WHERE company_id=? AND service_key=?",
            (company_id, service_key)).fetchone()
        return bool(row) and row["status"] == "active"
    finally:
        db.close()


def upsert_company_service(company_id, service_key, stripe_subscription_id=None,
                           stripe_price_id=None, status="active", current_period_end=None):
    """Create or update a company_services row. None arguments leave the
    existing column value untouched."""
    db = get_db()
    try:
        row = db.execute(
            "SELECT id FROM company_services WHERE company_id=? AND service_key=?",
            (company_id, service_key)).fetchone()
        if row:
            db.execute(
                """UPDATE company_services SET
                     stripe_subscription_id=COALESCE(?, stripe_subscription_id),
                     stripe_price_id=COALESCE(?, stripe_price_id),
                     status=COALESCE(?, status),
                     current_period_end=COALESCE(?, current_period_end)
                   WHERE id=?""",
                (stripe_subscription_id, stripe_price_id, status,
                 current_period_end, row["id"]))
        else:
            db.execute(
                """INSERT INTO company_services
                   (company_id, service_key, status, stripe_subscription_id,
                    stripe_price_id, current_period_end, created_at)
                   VALUES (?,?,?,?,?,?,?)""",
                (company_id, service_key, status or "active", stripe_subscription_id,
                 stripe_price_id, current_period_end, _utcnow()))
        db.commit()
    finally:
        db.close()


def save_tabulation(user_id, project_name: str, data: dict) -> int:
    db = get_db()
    try:
        cur = db.execute(
            "INSERT INTO tabulations (user_id, project_name, data_json, created_at, claimed)"
            " VALUES (?,?,?,?,?)",
            (user_id, project_name, json.dumps(data), int(time.time()),
             1 if user_id else 0))
        db.commit()
        return cur.lastrowid
    finally:
        db.close()


def claim_tabulations(user_id: int, tab_ids):
    if not tab_ids:
        return 0
    db = get_db()
    try:
        n = 0
        for tid in tab_ids:
            cur = db.execute(
                "UPDATE tabulations SET user_id=?, claimed=1 WHERE id=? AND user_id IS NULL",
                (user_id, int(tid)))
            n += cur.rowcount
        db.commit()
        return n
    finally:
        db.close()


def create_password_token(user_id, purpose="set_password", ttl_seconds=72 * 3600):
    """Mint a one-time token; only its hash is stored. Returns the raw token."""
    raw = secrets.token_urlsafe(32)
    th = hashlib.sha256(raw.encode()).hexdigest()
    now = int(time.time())
    db = get_db()
    try:
        db.execute(
            "INSERT INTO password_tokens (user_id, token_hash, purpose, expires_at, used_at, created_at)"
            " VALUES (?,?,?,?,0,?)",
            (user_id, th, purpose, now + ttl_seconds, now))
        db.commit()
    finally:
        db.close()
    return raw


def consume_password_token(raw_token, purpose="set_password"):
    """Validate + burn a token. Returns user_id, or None if invalid/expired/used."""
    if not raw_token:
        return None
    th = hashlib.sha256(raw_token.encode()).hexdigest()
    now = int(time.time())
    db = get_db()
    try:
        row = db.execute(
            "SELECT id, user_id, expires_at, used_at FROM password_tokens"
            " WHERE token_hash=? AND purpose=?", (th, purpose)).fetchone()
        if not row or row["used_at"] or row["expires_at"] < now:
            return None
        db.execute("UPDATE password_tokens SET used_at=? WHERE id=?", (now, row["id"]))
        db.commit()
        return row["user_id"]
    finally:
        db.close()


def send_email(to, subject, body_html, text=None):
    """Send one email via Resend. Returns True if accepted. Without a
    RESEND_API_KEY configured, returns False so callers can degrade."""
    key = os.environ.get("RESEND_API_KEY")
    if not key:
        return False
    payload = {"from": "ClearScope Tabulator <coe@tabulator.clearscopebid.com>",
               "to": [to], "subject": subject, "html": body_html}
    if text:
        payload["text"] = text
    req = urllib.request.Request(
        "https://api.resend.com/emails",
        data=json.dumps(payload).encode(),
        headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json",
                 # Resend sits behind Cloudflare, which rejects the default
                 # Python-urllib user-agent with error 1010 (diagnosed
                 # 2026-10-12). A named UA is required.
                 "User-Agent": "ClearScope-Tabulator/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return 200 <= resp.status < 300
    except urllib.error.HTTPError as e:
        # Surface Resend's own error name/message so a rejection (e.g. a
        # suppressed address) is self-diagnosing in the logs.
        detail = ""
        try:
            err_body = json.loads(e.read().decode("utf-8", "replace") or "{}")
            parts = [str(err_body[k]) for k in ("name", "message") if err_body.get(k)]
            if parts:
                detail = " | resend_error: " + ": ".join(parts)
        except Exception:
            pass
        print(f"[send_email] FAILED to={to}: {type(e).__name__}: {e}{detail}", flush=True)
        return False
    except Exception as e:
        print(f"[send_email] FAILED to={to}: {type(e).__name__}: {e}", flush=True)
        return False


def list_tabulations(user_id: int):
    db = get_db()
    try:
        rows = db.execute(
            "SELECT id, project_name, created_at FROM tabulations"
            " WHERE user_id=? ORDER BY id DESC", (user_id,)).fetchall()
        return [dict(r) for r in rows]
    finally:
        db.close()


def get_tabulation(user_id: int, tab_id: int):
    db = get_db()
    try:
        row = db.execute(
            "SELECT id, project_name, data_json, created_at FROM tabulations"
            " WHERE id=? AND user_id=?", (tab_id, user_id)).fetchone()
        if not row:
            return None
        d = dict(row)
        d["data"] = json.loads(d.pop("data_json"))
        return d
    finally:
        db.close()


def can_upload(user):
    """Returns (allowed: bool, reason: str)."""
    if user:
        if subscription_status(user["id"]) == "active":
            return True, "subscriber"
        if service_active(user["id"], "suite"):
            return True, "suite"
        if tabulation_count(user["id"]) < FREE_TABULATIONS:
            return True, "free"
        return False, "needs_subscription"
    # anonymous: one free tabulation per session
    if session.get("free_used"):
        return False, "needs_account"
    return True, "free_anonymous"


def register_routes(app):
    _SET_PASSWORD_HTML = """<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Set your password — ClearScope Tabulator</title>
<style>
 body{font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;background:#F5F3EF;margin:0;padding:40px 16px}
 .card{max-width:420px;margin:0 auto;background:#fff;border-radius:18px;padding:28px;box-shadow:0 8px 30px rgba(16,42,67,.12)}
 h1{color:#102A43;font-size:22px;margin:0 0 8px}
 p{color:#40566b;font-size:14px;line-height:1.5}
 label{display:block;font-size:13px;color:#102A43;margin:14px 0 4px;font-weight:600}
 input{width:100%;box-sizing:border-box;padding:10px;border:1px solid #cbd5e1;border-radius:8px;font-size:15px}
 button{margin-top:18px;width:100%;background:#E56A2B;color:#fff;border:0;border-radius:8px;padding:12px;font-size:15px;font-weight:700;cursor:pointer}
 button:disabled{opacity:.6}
 .err{color:#b91c1c;font-size:13px;margin-top:10px}
</style></head><body>
<div class="card">
 <h1>Set your password</h1>
 <p>Your tabulation is saved to your free ClearScope Tabulator account. Pick a password to sign in and open it anytime.</p>
 <label for="pw">New password (8+ characters)</label>
 <input id="pw" type="password" autocomplete="new-password">
 <label for="pw2">Confirm password</label>
 <input id="pw2" type="password" autocomplete="new-password">
 <div id="msg" class="err"></div>
 <button id="go">Save password &amp; open my tabulation</button>
</div>
<script>
var TOKEN = "__TOKEN__";
document.getElementById("go").onclick = function(){
  var pw = document.getElementById("pw").value, pw2 = document.getElementById("pw2").value;
  var msg = document.getElementById("msg");
  if (pw.length < 8){ msg.textContent = "Password must be at least 8 characters."; return; }
  if (pw !== pw2){ msg.textContent = "Passwords do not match."; return; }
  var btn = this; btn.disabled = true; msg.textContent = "";
  fetch("/api/set-password", {method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({token: TOKEN, password: pw})})
  .then(function(r){ return r.json().then(function(j){ return {ok:r.ok, body:j}; }); })
  .then(function(res){
    if (res.ok && res.body.ok){ location.href = "/"; }
    else { btn.disabled = false; msg.textContent = (res.body && res.body.error) || "Something went wrong."; }
  }).catch(function(){ btn.disabled = false; msg.textContent = "Network error — try again."; });
};
</script></body></html>"""

    @app.post("/api/save-by-email")
    def save_by_email():
        """Email-first save: attach an anonymous tabulation to the account for
        `email` (created on the spot with an unusable random password) and
        email a one-time set-password link. Possession of the tabulation's
        claim token (issued only to the browser that ran it) is the proof."""
        body = request.get_json(force=True, silent=True) or {}
        email = (body.get("email") or "").strip().lower()
        token = (body.get("claim_token") or "").strip()
        if "@" not in email or "." not in email.split("@")[-1]:
            return jsonify(error="Enter a valid email address."), 400
        if not os.environ.get("RESEND_API_KEY"):
            return jsonify(error="email_unavailable",
                           message="Email save isn't available right now — use the password option below instead."), 503
        try:
            tid = int(body.get("tabulation_id") or 0)
        except (TypeError, ValueError):
            tid = 0
        db = get_db()
        try:
            tab = db.execute(
                "SELECT id, user_id, claim_token, project_name FROM tabulations WHERE id=?",
                (tid,)).fetchone()
            if not tab:
                return jsonify(error="Tabulation not found."), 400
            row = db.execute("SELECT id FROM users WHERE email=?", (email,)).fetchone()
            uid = row["id"] if row else None
            if tab["user_id"] is None:
                # Unclaimed: the claim token from the uploading browser is the proof.
                if not tab["claim_token"] or not hmac.compare_digest(tab["claim_token"], token):
                    return jsonify(error="This tabulation can't be saved that way."), 400
                if uid is None:
                    cur = db.execute(
                        "INSERT INTO users (email, password_hash, created_at) VALUES (?,?,?)",
                        (email, hash_password(secrets.token_urlsafe(24)), int(time.time())))
                    db.commit()
                    uid = cur.lastrowid
                if not claim_tabulations(uid, [tid]):
                    return jsonify(error="This tabulation was already saved."), 400
            elif tab["user_id"] != uid:
                return jsonify(error="This tabulation is saved to a different account."), 400
            # else: already this email's tabulation -> fall through and re-send
            # a fresh link (recovery if the first email never arrived).
            project = tab["project_name"]
        finally:
            db.close()
        raw = create_password_token(uid)
        base = os.environ.get("APP_URL", "https://tabulator.clearscopebid.com").rstrip("/")
        link = f"{base}/set-password?token={raw}"
        body_html = (
            f"<p>Your bid tabulation for <strong>{html_escape(project or 'your bid package')}</strong> "
            "is saved to your free ClearScope Tabulator account.</p>"
            f"<p><a href=\"{link}\">Set your password and open your tabulation</a></p>"
            "<p>The link works once and expires in 72 hours. If it expires, run a new "
            "tabulation and use the save step again — or just reply to this email.</p>"
            "<p>— ClearScope Tabulator</p>")
        ok = send_email(email, "Your bid tabulation is saved", body_html)
        # Also hand the SAME one-time link to this browser session: possession
        # of the claim token already proved this browser owns the save, so
        # in-session delivery is no weaker than the email — and an address
        # Resend permanently refuses is no longer locked out. Only this
        # response carries the path; the token is never listed elsewhere.
        return jsonify(ok=True, emailed=ok,
                       set_password_path=f"/set-password?token={raw}")

    @app.get("/set-password")
    def set_password_page():
        token = (request.args.get("token") or "").strip()
        return _SET_PASSWORD_HTML.replace("__TOKEN__", token)

    @app.post("/api/set-password")
    def set_password():
        body = request.get_json(force=True, silent=True) or {}
        password = body.get("password") or ""
        if len(password) < 8:
            return jsonify(error="Password must be at least 8 characters."), 400
        uid = consume_password_token(body.get("token") or "")
        if not uid:
            return jsonify(error="This link is invalid or has expired."), 400
        db = get_db()
        try:
            db.execute("UPDATE users SET password_hash=? WHERE id=?",
                       (hash_password(password), uid))
            db.commit()
        finally:
            db.close()
        session["user_id"] = uid
        session.pop("free_used", None)
        return jsonify(ok=True)

    @app.post("/api/signup")
    def signup():
        body = request.get_json(force=True, silent=True) or {}
        email = (body.get("email") or "").strip().lower()
        password = body.get("password") or ""
        if "@" not in email or "." not in email.split("@")[-1]:
            return jsonify(error="Enter a valid email address."), 400
        if len(password) < 8:
            return jsonify(error="Password must be at least 8 characters."), 400
        db = get_db()
        try:
            exists = db.execute("SELECT id FROM users WHERE email=?", (email,)).fetchone()
            if exists:
                return jsonify(error="That email is already registered. Try logging in."), 409
            cur = db.execute(
                "INSERT INTO users (email, password_hash, created_at) VALUES (?,?,?)",
                (email, hash_password(password), int(time.time())))
            db.commit()
            uid = cur.lastrowid
        finally:
            db.close()
        session["user_id"] = uid
        claimed = claim_tabulations(uid, body.get("claim_tabulations") or [])
        # anonymous free flag consumed by signup
        session.pop("free_used", None)
        return jsonify(ok=True, user={"id": uid, "email": email}, claimed=claimed)

    @app.post("/api/login")
    def login():
        body = request.get_json(force=True, silent=True) or {}
        email = (body.get("email") or "").strip().lower()
        password = body.get("password") or ""
        db = get_db()
        try:
            row = db.execute("SELECT id, email, password_hash FROM users WHERE email=?",
                             (email,)).fetchone()
        finally:
            db.close()
        if not row or not verify_password(password, row["password_hash"]):
            return jsonify(error="Invalid email or password."), 401
        session["user_id"] = row["id"]
        session.pop("free_used", None)
        return jsonify(ok=True, user={"id": row["id"], "email": row["email"]})

    @app.post("/api/logout")
    def logout():
        session.clear()
        return jsonify(ok=True)

    @app.get("/api/me")
    def me():
        u = current_user()
        if not u:
            return jsonify(user=None)
        sub = subscription_status(u["id"])
        return jsonify(user={
            "id": u["id"], "email": u["email"],
            "company_name": u.get("company_name") or "",
            "contact_name": u.get("contact_name") or "",
            "phone": u.get("phone") or "",
            "is_admin": is_admin(u),
            "subscription": sub,
            "tabulations": tabulation_count(u["id"]),
            "free_remaining": max(0, FREE_TABULATIONS - tabulation_count(u["id"])),
            "stripe_configured": stripe_configured(),
        })

    @app.get("/api/settings")
    @login_required
    def settings_get():
        u = g.user
        return jsonify(settings={
            "company_name": u.get("company_name") or "",
            "contact_name": u.get("contact_name") or "",
            "phone": u.get("phone") or "",
            "email": u["email"],
        })

    @app.post("/api/settings")
    @login_required
    def settings_post():
        body = request.get_json(force=True, silent=True) or {}
        db = get_db()
        try:
            db.execute(
                "UPDATE users SET company_name=?, contact_name=?, phone=? WHERE id=?",
                ((body.get("company_name") or "").strip(),
                 (body.get("contact_name") or "").strip(),
                 (body.get("phone") or "").strip(),
                 g.user["id"]))
            db.commit()
        finally:
            db.close()
        return jsonify(ok=True)

    @app.post("/api/change-password")
    @login_required
    def change_password():
        body = request.get_json(force=True, silent=True) or {}
        old_pw = body.get("current_password") or ""
        new_pw = body.get("new_password") or ""
        if len(new_pw) < 8:
            return jsonify(error="New password must be at least 8 characters."), 400
        db = get_db()
        try:
            row = db.execute("SELECT password_hash FROM users WHERE id=?",
                             (g.user["id"],)).fetchone()
            if not row or not verify_password(old_pw, row["password_hash"]):
                return jsonify(error="Current password is incorrect."), 401
            db.execute("UPDATE users SET password_hash=? WHERE id=?",
                       (hash_password(new_pw), g.user["id"]))
            db.commit()
        finally:
            db.close()
        return jsonify(ok=True)

    @app.get("/api/admin/users")
    @admin_required
    def admin_users():
        db = get_db()
        try:
            rows = db.execute(
                """SELECT u.id, u.email, u.created_at, u.company_name,
                          (SELECT COUNT(*) FROM tabulations t WHERE t.user_id=u.id) AS tab_count,
                          COALESCE(s.status, 'none') AS sub_status
                   FROM users u
                   LEFT JOIN subscriptions s ON s.user_id=u.id
                   ORDER BY u.id DESC""").fetchall()
            return jsonify(users=[dict(r) for r in rows])
        finally:
            db.close()

    @app.get("/api/admin/stats")
    @admin_required
    def admin_stats():
        db = get_db()
        try:
            total_users = db.execute("SELECT COUNT(*) c FROM users").fetchone()["c"]
            total_tabs = db.execute("SELECT COUNT(*) c FROM tabulations").fetchone()["c"]
            paying = db.execute(
                "SELECT COUNT(*) c FROM subscriptions WHERE status IN ('active','trialing','past_due')"
            ).fetchone()["c"]
            annual = db.execute(
                "SELECT COUNT(*) c FROM subscriptions WHERE status IN ('active','trialing','past_due') AND billing_interval='year'"
            ).fetchone()["c"]
            monthly = paying - annual
            return jsonify(stats={
                "total_users": total_users,
                "total_tabulations": total_tabs,
                "paying_customers": paying,
                "mrr": monthly * 129 + annual * 1290 // 12,
            })
        finally:
            db.close()

    @app.get("/api/tabulations")
    @login_required
    def tab_list():
        return jsonify(tabulations=list_tabulations(g.user["id"]))

    @app.get("/api/tabulations/<int:tab_id>")
    @login_required
    def tab_get(tab_id):
        t = get_tabulation(g.user["id"], tab_id)
        if not t:
            return jsonify(error="Not found."), 404
        return jsonify(t)

    @app.post("/api/stripe/checkout")
    @login_required
    def stripe_checkout():
        if not stripe_configured():
            return jsonify(error="Billing is not configured yet. Please check back soon."), 503
        body = request.get_json(force=True, silent=True) or {}
        interval = (body.get("interval") or "month").lower()
        price_id = STRIPE_PRICE_ID_ANNUAL if interval == "year" and STRIPE_PRICE_ID_ANNUAL else STRIPE_PRICE_ID
        stripe.api_key = STRIPE_SECRET_KEY
        db = get_db()
        try:
            sub = db.execute("SELECT stripe_customer_id FROM subscriptions WHERE user_id=?",
                             (g.user["id"],)).fetchone()
            customer_id = sub["stripe_customer_id"] if sub else None
        finally:
            db.close()
        try:
            kwargs = {
                "mode": "subscription",
                "line_items": [{"price": price_id, "quantity": 1}],
                "success_url": f"{APP_URL}/?billing=success",
                "cancel_url": f"{APP_URL}/?billing=cancelled",
                "metadata": {"user_id": str(g.user["id"]), "interval": interval},
            }
            if customer_id:
                kwargs["customer"] = customer_id
            else:
                kwargs["customer_email"] = g.user["email"]
            sess = stripe.checkout.Session.create(**kwargs)
            return jsonify(url=sess.url)
        except Exception as e:
            return jsonify(error=f"Could not start checkout: {e}"), 500

    @app.post("/api/stripe/portal")
    @login_required
    def stripe_portal():
        if not stripe_configured():
            return jsonify(error="Billing is not configured yet."), 503
        stripe.api_key = STRIPE_SECRET_KEY
        db = get_db()
        try:
            sub = db.execute("SELECT stripe_customer_id FROM subscriptions WHERE user_id=?",
                             (g.user["id"],)).fetchone()
        finally:
            db.close()
        if not sub or not sub["stripe_customer_id"]:
            return jsonify(error="No billing account found."), 404
        try:
            ps = stripe.billing_portal.Session.create(
                customer=sub["stripe_customer_id"],
                return_url=f"{APP_URL}/")
            return jsonify(url=ps.url)
        except Exception as e:
            return jsonify(error=f"Could not open billing portal: {e}"), 500

    @app.post("/api/stripe/webhook")
    def stripe_webhook():
        if not stripe or not STRIPE_WEBHOOK_SECRET:
            return jsonify(error="webhook not configured"), 503
        payload = request.get_data()
        sig = request.headers.get("Stripe-Signature", "")
        try:
            event = stripe.Webhook.construct_event(payload, sig, STRIPE_WEBHOOK_SECRET)
        except Exception:
            return jsonify(error="bad signature"), 400
        stripe.api_key = STRIPE_SECRET_KEY
        etype = event["type"]
        obj = event["data"]["object"]
        # stripe-python >= 12 resource objects are not dicts (no .get());
        # normalize to a plain dict so every branch below behaves the same.
        if hasattr(obj, "to_dict"):
            obj = obj.to_dict()

        def upsert(user_id, customer_id, sub_id, status, period_end, billing_interval="month"):
            db = get_db()
            try:
                db.execute(
                    """INSERT INTO subscriptions
                       (user_id, stripe_customer_id, stripe_subscription_id, status,
                        current_period_end, billing_interval, updated_at)
                       VALUES (?,?,?,?,?,?,?)
                       ON CONFLICT(user_id) DO UPDATE SET
                         stripe_customer_id=excluded.stripe_customer_id,
                         stripe_subscription_id=excluded.stripe_subscription_id,
                         status=excluded.status,
                         current_period_end=excluded.current_period_end,
                         billing_interval=excluded.billing_interval,
                         updated_at=excluded.updated_at""",
                    (user_id, customer_id, sub_id, status, period_end, billing_interval, int(time.time())))
                db.commit()
            finally:
                db.close()

        if etype == "checkout.session.completed":
            sess = obj
            uid = int((sess.get("metadata") or {}).get("user_id") or 0)
            customer_id = sess.get("customer")
            sub_id = sess.get("subscription")
            interval = (sess.get("metadata") or {}).get("interval") or "month"
            service_key = (sess.get("metadata") or {}).get("service_key") or ""
            if uid and sub_id:
                sub = stripe.Subscription.retrieve(sub_id)
                if hasattr(sub, "to_dict"):
                    sub = sub.to_dict()
                if service_key in ("getpaid", "suite"):
                    # Service checkout (Get-Paid or GC Suite): record the service row.
                    price_id = None
                    try:
                        items = (sub.get("items") or {}).get("data") or []
                        if items:
                            price_id = (items[0].get("price") or {}).get("id")
                    except Exception:
                        pass
                    cpe = sub.get("current_period_end")
                    upsert_company_service(
                        uid, service_key,
                        stripe_subscription_id=sub_id,
                        stripe_price_id=price_id,
                        status=sub["status"],
                        current_period_end=(_utcnow() if not cpe else
                                           datetime.fromtimestamp(cpe, timezone.utc).isoformat()))
                else:
                    upsert(uid, customer_id, sub_id, sub["status"],
                           sub.get("current_period_end"), interval)
        elif etype in ("customer.subscription.updated", "customer.subscription.deleted"):
            sub = obj
            customer_id = sub.get("customer")
            # Get-Paid service rows are keyed by stripe_subscription_id.
            cpe = sub.get("current_period_end")
            gp_row = None
            db = get_db()
            try:
                gp_row = db.execute(
                    "SELECT company_id, service_key FROM company_services"
                    " WHERE stripe_subscription_id=?", (sub["id"],)).fetchone()
            finally:
                db.close()
            if gp_row:
                upsert_company_service(
                    gp_row["company_id"], gp_row["service_key"],
                    status=sub["status"],
                    current_period_end=(_utcnow() if not cpe else
                                       datetime.fromtimestamp(cpe, timezone.utc).isoformat()))
            # derive interval from the subscription's price
            interval = "month"
            try:
                items = (sub.get("items") or {}).get("data") or []
                if items:
                    interval = ((items[0].get("price") or {}).get("recurring") or {}).get("interval") or "month"
            except Exception:
                pass
            # find user by customer id
            db = get_db()
            try:
                row = db.execute("SELECT user_id FROM subscriptions WHERE stripe_customer_id=?",
                                 (customer_id,)).fetchone()
            finally:
                db.close()
            if row:
                upsert(row["user_id"], customer_id, sub["id"], sub["status"],
                       sub.get("current_period_end"), interval)
        elif etype == "invoice.payment_failed":
            # keep status as-is; Stripe retries. Mark past_due via subscription.updated.
            pass
        return jsonify(received=True)

    @app.get("/api/billing/config")
    def billing_config():
        return jsonify(publishable_key=STRIPE_PUBLISHABLE_KEY or None,
                       configured=stripe_configured())

    # ------------------------------------------------- services (multi-product)

    @app.get("/api/billing/services")
    @login_required
    def billing_services():
        """Per-service subscription status for the Billing settings page."""
        tab_active = subscription_status(g.user["id"]) == "active"
        gp_active = service_active(g.user["id"], "getpaid")
        suite_active = service_active(g.user["id"], "suite")
        return jsonify(services=[
            {"key": "tabulator", "name": "Bid Tabulator",
             "price": "$129/mo", "price_monthly": "$129/mo", "price_annual": "$1,290/yr",
             "active": tab_active},
            {"key": "getpaid", "name": "Get-Paid suite",
             "price": "$99/mo", "price_monthly": "$99/mo", "price_annual": "$990/yr",
             "active": gp_active},
            {"key": "suite", "name": "GC Suite — Tabulator + Get-Paid",
             "price": "$179/mo", "price_monthly": "$179/mo", "price_annual": "$1,750/yr",
             "save_note": "Save $49/mo vs separate",
             "active": suite_active},
        ])

    @app.post("/api/billing/services/getpaid/subscribe")
    @login_required
    def getpaid_subscribe():
        """Stripe Checkout for the Get-Paid suite ($99/mo or $990/yr). Mirrors the
        /api/stripe/checkout pattern; the webhook upserts company_services."""
        body = request.get_json(force=True, silent=True) or {}
        interval = body.get("interval") or "month"
        price_id = STRIPE_PRICE_ID_GETPAID_ANNUAL if interval == "year" else STRIPE_PRICE_ID_GETPAID
        if not (stripe and STRIPE_SECRET_KEY and price_id):
            return jsonify(error="billing_not_configured"), 503
        stripe.api_key = STRIPE_SECRET_KEY
        db = get_db()
        try:
            # One account = one Stripe customer; reuse the tabulator customer.
            sub = db.execute("SELECT stripe_customer_id FROM subscriptions WHERE user_id=?",
                             (g.user["id"],)).fetchone()
            customer_id = sub["stripe_customer_id"] if sub else None
        finally:
            db.close()
        try:
            kwargs = {
                "mode": "subscription",
                "line_items": [{"price": price_id, "quantity": 1}],
                "success_url": f"{APP_URL}/?billing=getpaid_success",
                "cancel_url": f"{APP_URL}/?billing=getpaid_cancelled",
                "metadata": {"user_id": str(g.user["id"]), "service_key": "getpaid",
                             "interval": interval},
            }
            if customer_id:
                kwargs["customer"] = customer_id
            else:
                kwargs["customer_email"] = g.user["email"]
            sess = stripe.checkout.Session.create(**kwargs)
            return jsonify(url=sess.url)
        except Exception as e:
            return jsonify(error=f"Could not start checkout: {e}"), 500

    @app.post("/api/billing/services/getpaid/cancel")
    @login_required
    def getpaid_cancel():
        """Cancel the Get-Paid subscription at period end via Stripe, then mark
        the company_services row canceled."""
        if not getpaid_configured():
            return jsonify(error="billing_not_configured"), 503
        db = get_db()
        try:
            row = db.execute(
                "SELECT stripe_subscription_id FROM company_services"
                " WHERE company_id=? AND service_key='getpaid'",
                (g.user["id"],)).fetchone()
        finally:
            db.close()
        if not row or not row["stripe_subscription_id"]:
            return jsonify(error="No Get-Paid subscription found."), 404
        stripe.api_key = STRIPE_SECRET_KEY
        try:
            stripe.Subscription.modify(
                row["stripe_subscription_id"], cancel_at_period_end=True)
        except Exception as e:
            return jsonify(error=f"Could not cancel: {e}"), 500
        upsert_company_service(g.user["id"], "getpaid", status="canceled")
        return jsonify(ok=True)

    @app.post("/api/billing/services/suite/subscribe")
    @login_required
    def suite_subscribe():
        """Stripe Checkout for the GC Suite bundle ($179/mo or $1,750/yr).
        Unlocks BOTH Bid Tabulator and Get-Paid features."""
        body = request.get_json(force=True, silent=True) or {}
        interval = body.get("interval") or "month"
        price_id = STRIPE_PRICE_ID_SUITE_ANNUAL if interval == "year" else STRIPE_PRICE_ID_SUITE
        if not (stripe and STRIPE_SECRET_KEY and price_id):
            return jsonify(error="billing_not_configured"), 503
        stripe.api_key = STRIPE_SECRET_KEY
        db = get_db()
        try:
            sub = db.execute("SELECT stripe_customer_id FROM subscriptions WHERE user_id=?",
                             (g.user["id"],)).fetchone()
            customer_id = sub["stripe_customer_id"] if sub else None
        finally:
            db.close()
        try:
            kwargs = {
                "mode": "subscription",
                "line_items": [{"price": price_id, "quantity": 1}],
                "success_url": f"{APP_URL}/?billing=suite_success",
                "cancel_url": f"{APP_URL}/?billing=suite_cancelled",
                "metadata": {"user_id": str(g.user["id"]), "service_key": "suite",
                             "interval": interval},
            }
            if customer_id:
                kwargs["customer"] = customer_id
            else:
                kwargs["customer_email"] = g.user["email"]
            sess = stripe.checkout.Session.create(**kwargs)
            return jsonify(url=sess.url)
        except Exception as e:
            return jsonify(error=f"Could not start checkout: {e}"), 500

    @app.post("/api/billing/services/suite/cancel")
    @login_required
    def suite_cancel():
        """Cancel the GC Suite subscription at period end via Stripe, then mark
        the company_services row canceled."""
        if not suite_configured():
            return jsonify(error="billing_not_configured"), 503
        db = get_db()
        try:
            row = db.execute(
                "SELECT stripe_subscription_id FROM company_services"
                " WHERE company_id=? AND service_key='suite'",
                (g.user["id"],)).fetchone()
        finally:
            db.close()
        if not row or not row["stripe_subscription_id"]:
            return jsonify(error="No GC Suite subscription found."), 404
        stripe.api_key = STRIPE_SECRET_KEY
        try:
            stripe.Subscription.modify(
                row["stripe_subscription_id"], cancel_at_period_end=True)
        except Exception as e:
            return jsonify(error=f"Could not cancel: {e}"), 500
        upsert_company_service(g.user["id"], "suite", status="canceled")
        return jsonify(ok=True)


# --- TEMPORARY one-shot founder-comp probe (removed in the restore commit) ---
def _register_founder_comp_probe(app):
    @app.post("/api/internal/founder-comp")
    def founder_comp():
        expected = os.environ.get("COMP_TOKEN", "")
        provided = request.headers.get("X-Comp-Token", "")
        if not expected or not provided or not hmac.compare_digest(provided, expected):
            return jsonify(error="not found"), 404

        def snap(db):
            subs = [dict(r) for r in db.execute("SELECT * FROM subscriptions ORDER BY user_id")]
            svcs = [dict(r) for r in db.execute("SELECT * FROM company_services ORDER BY id")]
            return subs, svcs

        db = get_db()
        try:
            before_subs, before_svcs = snap(db)
            now = int(time.time())
            cpe = now + 30 * 24 * 3600
            # Founder comp: same upsert the Stripe webhook runs for a real
            # Tabulator monthly checkout, except the Stripe ids stay NULL —
            # deliberately no Stripe data fabricated.
            db.execute(
                """INSERT INTO subscriptions
                   (user_id, stripe_customer_id, stripe_subscription_id, status,
                    current_period_end, billing_interval, updated_at)
                   VALUES (?,?,?,?,?,?,?)
                   ON CONFLICT(user_id) DO UPDATE SET
                     stripe_customer_id=excluded.stripe_customer_id,
                     stripe_subscription_id=excluded.stripe_subscription_id,
                     status=excluded.status,
                     current_period_end=excluded.current_period_end,
                     billing_interval=excluded.billing_interval,
                     updated_at=excluded.updated_at""",
                (1, None, None, "active", cpe, "month", now))
            db.commit()
            after_subs, after_svcs = snap(db)
        finally:
            db.close()
        return jsonify(
            before_subscriptions=before_subs,
            before_company_services=before_svcs,
            after_subscriptions=after_subs,
            after_company_services=after_svcs,
            subscription_status_1=subscription_status(1),
            can_upload_1=can_upload({"id": 1}),
            billing_services_1={
                "tabulator_active": subscription_status(1) == "active",
                "getpaid_active": service_active(1, "getpaid"),
                "suite_active": service_active(1, "suite"),
            })


try:
    _register_founder_comp_probe(app)
except NameError:
    pass
