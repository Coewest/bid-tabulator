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
import json
import os
import secrets
import sqlite3
import time
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
APP_URL = os.environ.get("APP_URL", "https://bid-tabulator-production.up.railway.app").rstrip("/")

FREE_TABULATIONS = 1


def stripe_configured():
    return bool(stripe and STRIPE_SECRET_KEY and STRIPE_PRICE_ID)


def get_db():
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    return db


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
    """)
    # Migration: company profile fields
    cols = [r[1] for r in db.execute("PRAGMA table_info(users)")]
    for col in ("company_name", "contact_name", "phone"):
        if col not in cols:
            db.execute(f"ALTER TABLE users ADD COLUMN {col} TEXT NOT NULL DEFAULT ''")
    # Migration: subscription billing interval (month/year)
    scols = [r[1] for r in db.execute("PRAGMA table_info(subscriptions)")]
    if "billing_interval" not in scols:
        db.execute("ALTER TABLE subscriptions ADD COLUMN billing_interval TEXT NOT NULL DEFAULT 'month'")
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
        if tabulation_count(user["id"]) < FREE_TABULATIONS:
            return True, "free"
        return False, "needs_subscription"
    # anonymous: one free tabulation per session
    if session.get("free_used"):
        return False, "needs_account"
    return True, "free_anonymous"


def register_routes(app):
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
            if uid and sub_id:
                sub = stripe.Subscription.retrieve(sub_id)
                upsert(uid, customer_id, sub_id, sub["status"],
                       sub.get("current_period_end"), interval)
        elif etype in ("customer.subscription.updated", "customer.subscription.deleted"):
            sub = obj
            customer_id = sub.get("customer")
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
