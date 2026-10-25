"""
Excuse Letter Board - standalone prototype (Flask + SQLite)

Flow:
  Viewer  -> fills in a request form -> gets an anonymous code (e.g. EX-7K3QA)
  Board   -> shows ONLY the code, category, dates and status (no names)
  Admin   -> logs in, reviews, approves/rejects, downloads .docx, archives, deletes
  Approve -> posts a public announcement on the board (the "notify everyone" part)
"""
import io
import os
import re
import secrets
import sqlite3
import hmac
import threading
import time
from datetime import date, datetime, timedelta, timezone
from functools import wraps

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from flask import (Flask, abort, flash, g, jsonify, redirect, render_template,
                   request, send_file, send_from_directory, session, url_for)
from werkzeug.middleware.proxy_fix import ProxyFix

app = Flask(__name__)

# ---------------------------------------------------------------- CONFIG ----
# Values come from environment variables so secrets never live in the code.
# Render sets RENDER=true automatically, which is how we know we're "live".
IS_PROD = os.environ.get("RENDER") == "true" or os.environ.get("PRODUCTION") == "1"

DEV_SECRET = "dev-only-change-me"
DEV_PASSWORD = "admin123"
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", DEV_SECRET)
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", DEV_PASSWORD)

# DATABASE_URL set  -> use PostgreSQL (data survives redeploys; used on Render)
# DATABASE_URL empty -> use a local SQLite file (easy testing on your laptop)
DATABASE_URL = os.environ.get("DATABASE_URL")
DB_PATH = os.environ.get("DATABASE_PATH", "excuse.db")  # SQLite file name
USE_PG = bool(DATABASE_URL)

# FAIL FAST: when live, refuse to start with unsafe settings instead of
# quietly running with a known password or an ephemeral database.
if IS_PROD:
    problems = []
    if app.config["SECRET_KEY"] == DEV_SECRET:
        problems.append("SECRET_KEY is not set")
    if ADMIN_PASSWORD == DEV_PASSWORD or len(ADMIN_PASSWORD) < 8:
        problems.append("ADMIN_PASSWORD is missing, default, or shorter than 8 chars")
    if not USE_PG:
        problems.append("DATABASE_URL is not set (data would be lost on redeploy)")
    if problems:
        raise RuntimeError("Unsafe production config: " + "; ".join(problems))
    # Render sits behind a proxy; this makes request.remote_addr the REAL visitor IP
    # (needed for rate limiting) and makes https detection work.
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1)

# Cookie safety: JS can't read it, only sent to our site, only over https when live.
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=IS_PROD,
    PERMANENT_SESSION_LIFETIME=timedelta(hours=8),
    MAX_CONTENT_LENGTH=64 * 1024,   # reject huge request bodies (64 KB is plenty)
)

CATEGORIES = ["Illness", "Medical appointment", "Family emergency",
              "School activity", "Bereavement", "Other"]

# Rate limits: (max attempts, per how many seconds). Tune these for your campus.
# Careful: many students may share one wifi/IP, so don't make them too strict.
LIMITS = {
    "login":   (10, 15 * 60),
    "request": (15, 10 * 60),
    "track":   (30, 10 * 60),
}

# -------------------------------------------------------------- DATABASE ----
class DB:
    """A tiny wrapper so the rest of the app doesn't care which database it
    is talking to. SQLite writes placeholders as ?  while PostgreSQL uses %s,
    so execute() swaps them automatically."""

    def __init__(self):
        if USE_PG:
            import psycopg2                       # only needed on PostgreSQL
            from psycopg2.extras import RealDictCursor
            # RealDictCursor = rows behave like dicts: row["code"]
            self.conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
        else:
            self.conn = sqlite3.connect(DB_PATH)
            self.conn.row_factory = sqlite3.Row   # rows behave like dicts too

    def execute(self, sql, params=()):
        if USE_PG:
            cur = self.conn.cursor()
            cur.execute(sql.replace("?", "%s"), params)
            return cur
        return self.conn.execute(sql, params)

    def commit(self):
        self.conn.commit()

    def close(self):
        self.conn.close()


def get_db():
    """Open one DB connection per request and reuse it."""
    if "db" not in g:
        g.db = DB()
    return g.db


@app.teardown_appcontext
def close_db(_exc):
    """Runs after every request: close the connection."""
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    """Create the tables if they don't exist yet."""
    # The only difference between the two databases: how auto-numbering is written.
    pk = "SERIAL PRIMARY KEY" if USE_PG else "INTEGER PRIMARY KEY AUTOINCREMENT"
    db = DB()
    db.execute(f"""
        CREATE TABLE IF NOT EXISTS letters (
            id            {pk},
            code          TEXT UNIQUE NOT NULL,   -- anonymous public ID
            student_name  TEXT NOT NULL,          -- PRIVATE (admin only)
            guardian_name TEXT NOT NULL,          -- PRIVATE (admin only)
            addressed_to  TEXT NOT NULL,
            category      TEXT NOT NULL,
            reason        TEXT NOT NULL,          -- PRIVATE (admin only)
            absent_from   TEXT NOT NULL,
            absent_to     TEXT NOT NULL,
            return_date   TEXT NOT NULL,
            status        TEXT NOT NULL DEFAULT 'pending',  -- pending/approved/rejected
            archived      INTEGER NOT NULL DEFAULT 0,       -- 0 = active, 1 = archived
            created_at    TEXT NOT NULL
        )""")
    db.execute(f"""
        CREATE TABLE IF NOT EXISTS announcements (
            id         {pk},
            message    TEXT NOT NULL,
            created_at TEXT NOT NULL
        )""")
    db.commit()
    db.close()


# Philippines has no daylight saving, so a fixed UTC+8 is always correct
# (and needs no extra package). Render servers run on UTC, hence this.
MANILA = timezone(timedelta(hours=8))


def now():
    """Current time in Manila, as text."""
    return datetime.now(MANILA).strftime("%Y-%m-%d %H:%M")


def today():
    return datetime.now(MANILA).date()


# Control characters (like \x00 or \x0b, which sneak in when pasting from Word/PDF)
# crash the .docx builder. We strip them but keep normal newlines.
_CONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def clean_text(value, maxlen=100, multiline=False):
    """Make user text safe: remove control chars, trim, cap the length."""
    value = (value or "").replace("\r\n", "\n").replace("\r", "\n")
    value = _CONTROL.sub("", value)
    if not multiline:
        value = " ".join(value.split())          # newlines/tabs -> single spaces
    return value.strip()[:maxlen]


def parse_date(value):
    """'2026-10-26' -> date object, or None if it isn't a REAL date."""
    try:
        return date.fromisoformat(value)
    except (ValueError, TypeError):
        return None


def safe_equal(a, b):
    """Constant-time compare that doesn't crash on non-ASCII text
    (hmac.compare_digest raises TypeError on str with é/ñ; bytes are fine)."""
    return hmac.compare_digest((a or "").encode("utf-8"), (b or "").encode("utf-8"))


def make_code(db):
    """Make a random anonymous code like EX-7K3QA (no confusing 0/O/1/I)."""
    alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
    while True:
        code = "EX-" + "".join(secrets.choice(alphabet) for _ in range(5))
        if not db.execute("SELECT 1 FROM letters WHERE code=?", (code,)).fetchone():
            return code


# -------------------------------------------------------------- SECURITY ----
def csrf_token():
    """One random token per session; every POST form must send it back.
    This stops other websites from tricking an admin into clicking 'delete'."""
    if "csrf" not in session:
        session["csrf"] = secrets.token_hex(16)
    return session["csrf"]


# Make csrf_token() and today's date usable inside every template.
app.jinja_env.globals["csrf_token"] = csrf_token


@app.before_request
def check_csrf():
    if request.method == "POST":
        expected = session.get("csrf")
        if not expected or not safe_equal(request.form.get("csrf", ""), expected):
            abort(400, "Invalid form token. Go back, refresh and try again.")


# --- tiny in-memory rate limiter (no extra package needed) ---
# Limitation: it lives in each server process and resets on restart. That's fine
# for stopping guessing/spam; for the super-app use Redis-backed limiting.
_hits = {}
_hits_lock = threading.Lock()


def too_many(bucket):
    """Record this visit; return True if the visitor went over the limit."""
    limit, window = LIMITS[bucket]
    key = (bucket, request.remote_addr)
    t = time.monotonic()
    with _hits_lock:
        recent = [x for x in _hits.get(key, []) if t - x < window]
        over = len(recent) >= limit
        if not over:
            recent.append(t)
        _hits[key] = recent
        if len(_hits) > 5000:  # housekeeping so memory can't grow forever
            for k in [k for k, v in _hits.items() if not v or t - v[-1] > 3600]:
                del _hits[k]
    return over


def reset_rate_limits():
    """Used by the tests."""
    with _hits_lock:
        _hits.clear()


def rate_limit(bucket):
    """Decorator: put under @app.route to limit POSTs on that route."""
    def deco(view):
        @wraps(view)
        def wrapper(*args, **kwargs):
            if request.method == "POST" and too_many(bucket):
                abort(429, "Too many attempts. Please wait a few minutes and try again.")
            return view(*args, **kwargs)
        return wrapper
    return deco


@app.after_request
def security_headers(resp):
    """Extra browser protections added to every response."""
    resp.headers["X-Content-Type-Options"] = "nosniff"        # don't guess file types
    resp.headers["X-Frame-Options"] = "DENY"                  # can't be put in an <iframe>
    resp.headers["Referrer-Policy"] = "same-origin"
    # Only our own scripts/styles may run. (That's why there are no inline scripts.)
    resp.headers["Content-Security-Policy"] = (
        "default-src 'self'; img-src 'self' data:; style-src 'self'; "
        "script-src 'self'; base-uri 'self'; form-action 'self'; frame-ancestors 'none'")
    if IS_PROD:
        resp.headers["Strict-Transport-Security"] = "max-age=31536000"
    if request.path.startswith("/admin"):                     # never store admin pages
        resp.headers["Cache-Control"] = "no-store"
    return resp


@app.errorhandler(400)
@app.errorhandler(404)
@app.errorhandler(429)
@app.errorhandler(500)
def show_error(err):
    """Friendly error pages instead of the ugly default."""
    code = getattr(err, "code", 500)
    msg = getattr(err, "description", "Something went wrong.") if code != 500 \
        else "Something went wrong on our side. Please try again."
    return render_template("error.html", code=code, message=msg), code


def admin_required(view):
    """Decorator: put @admin_required above any route that only admins may open."""
    @wraps(view)
    def wrapper(*args, **kwargs):
        if not session.get("is_admin"):
            flash("Please log in as admin first.", "error")
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapper


# ---------------------------------------------------------------- HELPERS ----
@app.template_filter("pretty_date")
def pretty_date(value):
    """'2026-10-25' -> 'Oct 25, 2026' (used in templates)."""
    try:
        return datetime.strptime(value[:10], "%Y-%m-%d").strftime("%b %d, %Y")
    except (ValueError, TypeError):
        return value


def build_letter(row):
    """Builds the letter text from the template. Used by BOTH the on-screen
    preview and the .docx download, so they always match."""
    # Clean again here so OLD rows saved before the fix can't crash the download.
    row = {**dict(row),
           "student_name": clean_text(row["student_name"]),
           "guardian_name": clean_text(row["guardian_name"]),
           "addressed_to": clean_text(row["addressed_to"]),
           "reason": clean_text(row["reason"], 1000, multiline=True)}
    if row["absent_from"] == row["absent_to"]:
        when = pretty_date(row["absent_from"])
    else:
        when = f'{pretty_date(row["absent_from"])} to {pretty_date(row["absent_to"])}'
    return {
        "date": pretty_date(row["created_at"]),
        "greeting": f'Dear {row["addressed_to"]},',
        "paragraphs": [
            f'I am writing to formally excuse {row["student_name"]} from class '
            f'on {when}.',
            f'The reason for the absence is as follows: {row["reason"]}',
            f'{row["student_name"]} is expected to return to school on '
            f'{pretty_date(row["return_date"])}. Thank you for your kind '
            f'understanding.',
        ],
        "farewell": "Respectfully yours,",
        "student": row["student_name"],
        "guardian": row["guardian_name"],
    }


# ---------------------------------------------------------- PUBLIC ROUTES ----
@app.route("/")
def board():
    """The public board: anonymous cards + latest announcements."""
    db = get_db()
    # NOTE: we only SELECT the safe columns. Names/reasons never reach the page.
    cards = db.execute(
        "SELECT code, category, absent_from, absent_to, status, created_at "
        "FROM letters WHERE archived=0 ORDER BY id DESC").fetchall()
    news = db.execute(
        "SELECT message, created_at FROM announcements "
        "ORDER BY id DESC LIMIT 5").fetchall()
    return render_template("board.html", cards=cards, news=news)


@app.route("/request", methods=["GET", "POST"])
@rate_limit("request")
def request_letter():
    """Form where a viewer asks for an excuse letter."""
    if request.method == "POST":
        f = request.form
        # Clean every field (strip control chars, trim, cap length).
        data = {
            "student_name": clean_text(f.get("student_name"), 100),
            "guardian_name": clean_text(f.get("guardian_name"), 100),
            "addressed_to": clean_text(f.get("addressed_to"), 100),
            "category": clean_text(f.get("category"), 50),
            "reason": clean_text(f.get("reason"), 1000, multiline=True),
            "absent_from": clean_text(f.get("absent_from"), 10),
            "absent_to": clean_text(f.get("absent_to"), 10),
            "return_date": clean_text(f.get("return_date"), 10),
        }

        # --- validation (never trust the browser) ---
        if not all(data.values()):
            flash("Please fill in every field.", "error")
            return render_template("request.html", categories=CATEGORIES, form=data)
        if data["category"] not in CATEGORIES:
            abort(400)

        d_from = parse_date(data["absent_from"])
        d_to = parse_date(data["absent_to"])
        d_back = parse_date(data["return_date"])
        if not (d_from and d_to and d_back):
            flash("One of the dates isn't a valid date.", "error")
            return render_template("request.html", categories=CATEGORIES, form=data)
        if d_to < d_from or d_back <= d_to:
            flash("Check your dates: the return date must be after the last "
                  "day of absence.", "error")
            return render_template("request.html", categories=CATEGORIES, form=data)
        if (d_to - d_from).days > 60 or not (
                today() - timedelta(days=90) <= d_from <= today() + timedelta(days=180)):
            flash("Those dates look unusual. Please double-check them.", "error")
            return render_template("request.html", categories=CATEGORIES, form=data)

        # Store the cleaned, normalized (YYYY-MM-DD) dates.
        data["absent_from"], data["absent_to"], data["return_date"] = (
            d_from.isoformat(), d_to.isoformat(), d_back.isoformat())

        db = get_db()
        code = make_code(db)
        db.execute(
            "INSERT INTO letters (code, student_name, guardian_name, addressed_to,"
            " category, reason, absent_from, absent_to, return_date, created_at)"
            " VALUES (?,?,?,?,?,?,?,?,?,?)",
            (code, data["student_name"], data["guardian_name"],
             data["addressed_to"], data["category"], data["reason"],
             data["absent_from"], data["absent_to"], data["return_date"], now()))
        db.commit()
        return redirect(url_for("submitted", code=code))
    return render_template("request.html", categories=CATEGORIES, form={})


@app.route("/submitted/<code>")
def submitted(code):
    """Shows the student their anonymous code - they must save it."""
    if not re.fullmatch(r"EX-[A-Z2-9]{5}", code):   # only real-looking codes
        abort(404)
    return render_template("submitted.html", code=code)


@app.route("/track", methods=["GET", "POST"])
@rate_limit("track")
def track():
    """Student types their code to check the status."""
    result, searched = None, None
    if request.method == "POST":
        searched = clean_text(request.form.get("code"), 12).upper()
        result = get_db().execute(
            "SELECT code, category, status, absent_from, absent_to, return_date "
            "FROM letters WHERE code=?", (searched,)).fetchone()
        if not result:
            flash("No letter found with that code.", "error")
    return render_template("track.html", result=result, searched=searched)


# ----------------------------------------------------------- ADMIN ROUTES ----
@app.route("/admin/login", methods=["GET", "POST"])
@rate_limit("login")
def login():
    if request.method == "POST":
        user = request.form.get("username", "")
        pw = request.form.get("password", "")
        # safe_equal = timing-attack-resistant compare that handles any characters.
        # Both checks always run (no early exit) so timing reveals nothing.
        user_ok = safe_equal(user, ADMIN_USERNAME)
        pw_ok = safe_equal(pw, ADMIN_PASSWORD)
        if user_ok and pw_ok:
            session.clear()           # fresh session after login
            session.permanent = True  # expires after PERMANENT_SESSION_LIFETIME (8h)
            session["is_admin"] = True
            return redirect(url_for("admin_panel"))
        app.logger.warning("Failed admin login from %s", request.remote_addr)
        flash("Wrong username or password.", "error")
    return render_template("login.html")


@app.route("/admin/logout", methods=["POST"])
def logout():
    session.clear()
    flash("Logged out.", "ok")
    return redirect(url_for("board"))


@app.route("/admin")
@admin_required
def admin_panel():
    """Admin dashboard. ?view=archived shows the archive instead."""
    view = request.args.get("view", "active")
    archived = 1 if view == "archived" else 0
    db = get_db()
    letters = db.execute(
        "SELECT * FROM letters WHERE archived=? ORDER BY id DESC",
        (archived,)).fetchall()
    # Little counters for the stat blocks at the top.
    counts = {s: db.execute(
        "SELECT COUNT(*) AS n FROM letters WHERE archived=0 AND status=?", (s,)
    ).fetchone()["n"] for s in ("pending", "approved", "rejected")}
    news = db.execute("SELECT * FROM announcements ORDER BY id DESC LIMIT 10").fetchall()
    return render_template("admin.html", letters=letters, counts=counts,
                           view=view, news=news)


def get_letter_or_404(letter_id):
    row = get_db().execute("SELECT * FROM letters WHERE id=?", (letter_id,)).fetchone()
    if row is None:
        abort(404)
    return row


@app.route("/admin/letter/<int:letter_id>")
@admin_required
def view_letter(letter_id):
    """Preview of the letter (has a print button + download button)."""
    row = get_letter_or_404(letter_id)
    return render_template("letter.html", row=row, letter=build_letter(row))


@app.route("/admin/letter/<int:letter_id>/download")
@admin_required
def download_letter(letter_id):
    """Builds a Word (.docx) file in memory and sends it."""
    row = get_letter_or_404(letter_id)
    L = build_letter(row)

    doc = Document()
    p = doc.add_paragraph(L["date"])
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph(L["greeting"])
    for text in L["paragraphs"]:
        doc.add_paragraph(text)
    doc.add_paragraph(L["farewell"])
    doc.add_paragraph("")
    doc.add_paragraph(f'{L["student"]}\nStudent')
    doc.add_paragraph(f'{L["guardian"]}\nParent / Guardian')

    buffer = io.BytesIO()          # a "fake file" in memory, nothing saved to disk
    doc.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True,
                     download_name=f'excuse-letter-{row["code"]}.docx',
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/admin/letter/<int:letter_id>/<action>", methods=["POST"])
@admin_required
def letter_action(letter_id, action):
    """One route for all buttons: approve / reject / archive / restore / delete."""
    row = get_letter_or_404(letter_id)
    db = get_db()
    # The exact text of this letter's public announcement (used to find/remove it).
    notice = f'Excuse letter {row["code"]} has been APPROVED.'

    if action == "approve":
        if row["status"] == "approved":
            flash(f'{row["code"]} is already approved.', "ok")   # double-click safe
        else:
            db.execute("UPDATE letters SET status='approved' WHERE id=?", (letter_id,))
            db.execute("INSERT INTO announcements (message, created_at) VALUES (?,?)",
                       (notice, now()))
            flash(f'{row["code"]} approved and announced on the board.', "ok")
    elif action == "reject":
        db.execute("UPDATE letters SET status='rejected' WHERE id=?", (letter_id,))
        # If it was announced as approved before, take that notice down.
        db.execute("DELETE FROM announcements WHERE message=?", (notice,))
        flash(f'{row["code"]} rejected.', "ok")
    elif action == "archive":
        db.execute("UPDATE letters SET archived=1 WHERE id=?", (letter_id,))
        flash(f'{row["code"]} archived (removed from the public board).', "ok")
    elif action == "restore":
        db.execute("UPDATE letters SET archived=0 WHERE id=?", (letter_id,))
        flash(f'{row["code"]} restored.', "ok")
    elif action == "delete":
        db.execute("DELETE FROM letters WHERE id=?", (letter_id,))
        db.execute("DELETE FROM announcements WHERE message=?", (notice,))
        flash(f'{row["code"]} permanently deleted.', "ok")
    else:
        abort(404)
    db.commit()
    # Go back to the tab the admin was on (a fixed list, never a user-supplied URL).
    view = "archived" if request.form.get("view") == "archived" else "active"
    return redirect(url_for("admin_panel", view=view))


@app.route("/admin/announcement/<int:news_id>/delete", methods=["POST"])
@admin_required
def delete_announcement(news_id):
    db = get_db()
    db.execute("DELETE FROM announcements WHERE id=?", (news_id,))
    db.commit()
    return redirect(url_for("admin_panel"))


# ------------------------------------------------------------ PWA ROUTES ----
# These make the site installable ("Add to Home Screen").
@app.route("/manifest.webmanifest")
def manifest():
    """The app's ID card: name, colors, icons. Browsers read this to offer install."""
    return jsonify({
        "name": "Excuse Letter Board",
        "short_name": "Excuse Board",
        "description": "Request and track excuse letters anonymously.",
        "id": "/",
        "start_url": "/",
        "scope": "/",
        "display": "standalone",          # opens without the browser address bar
        "background_color": "#1f2321",    # splash screen color (dark gray)
        "theme_color": "#228b22",         # forest green
        "icons": [
            {"src": "/static/icons/icon-192.png", "sizes": "192x192", "type": "image/png"},
            {"src": "/static/icons/icon-512.png", "sizes": "512x512", "type": "image/png"},
            {"src": "/static/icons/maskable-512.png", "sizes": "512x512",
             "type": "image/png", "purpose": "maskable"},
        ],
    })


@app.route("/sw.js")
def service_worker():
    """Service worker must be served from the site root so it covers every page.
    no-cache makes sure browsers always check for a newer version."""
    resp = send_from_directory(app.static_folder, "sw.js", mimetype="text/javascript")
    resp.headers["Cache-Control"] = "no-cache"
    return resp


@app.route("/offline")
def offline():
    """Page the service worker shows when there is no internet."""
    return render_template("offline.html")


# Create the tables when the app starts (also works under gunicorn on Render).
init_db()

if __name__ == "__main__":
    app.run(debug=True)  # local only; Render uses gunicorn instead
