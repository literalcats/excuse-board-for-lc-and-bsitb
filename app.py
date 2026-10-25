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
import secrets
import sqlite3
import hmac
from datetime import datetime, date
from functools import wraps

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from flask import (Flask, abort, flash, g, jsonify, redirect, render_template,
                   request, send_file, send_from_directory, session, url_for)

app = Flask(__name__)

# ---------------------------------------------------------------- CONFIG ----
# Values come from environment variables so secrets never live in the code.
# The fallbacks are ONLY for running on your own laptop.
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-only-change-me")
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin123")
# DATABASE_URL set  -> use PostgreSQL (data survives redeploys; used on Render)
# DATABASE_URL empty -> use a local SQLite file (easy testing on your laptop)
DATABASE_URL = os.environ.get("DATABASE_URL")
DB_PATH = os.environ.get("DATABASE_PATH", "excuse.db")  # SQLite file name
USE_PG = bool(DATABASE_URL)

CATEGORIES = ["Illness", "Medical appointment", "Family emergency",
              "School activity", "Bereavement", "Other"]


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


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


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
        sent = request.form.get("csrf", "")
        if not hmac.compare_digest(sent, session.get("csrf", "x")):
            abort(400, "Invalid form token. Go back, refresh and try again.")


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
def request_letter():
    """Form where a viewer asks for an excuse letter."""
    if request.method == "POST":
        f = request.form
        # Pull every field out and strip stray spaces.
        data = {k: f.get(k, "").strip() for k in (
            "student_name", "guardian_name", "addressed_to", "category",
            "reason", "absent_from", "absent_to", "return_date")}

        # --- validation (never trust the browser) ---
        if not all(data.values()):
            flash("Please fill in every field.", "error")
            return render_template("request.html", categories=CATEGORIES, form=data)
        if data["category"] not in CATEGORIES:
            abort(400)
        if data["absent_to"] < data["absent_from"] or data["return_date"] <= data["absent_to"]:
            flash("Check your dates: the return date must be after the last "
                  "day of absence.", "error")
            return render_template("request.html", categories=CATEGORIES, form=data)

        db = get_db()
        code = make_code(db)
        db.execute(
            "INSERT INTO letters (code, student_name, guardian_name, addressed_to,"
            " category, reason, absent_from, absent_to, return_date, created_at)"
            " VALUES (?,?,?,?,?,?,?,?,?,?)",
            (code, data["student_name"][:100], data["guardian_name"][:100],
             data["addressed_to"][:100], data["category"], data["reason"][:1000],
             data["absent_from"], data["absent_to"], data["return_date"], now()))
        db.commit()
        return redirect(url_for("submitted", code=code))
    return render_template("request.html", categories=CATEGORIES, form={})


@app.route("/submitted/<code>")
def submitted(code):
    """Shows the student their anonymous code - they must save it."""
    return render_template("submitted.html", code=code)


@app.route("/track", methods=["GET", "POST"])
def track():
    """Student types their code to check the status."""
    result, searched = None, None
    if request.method == "POST":
        searched = request.form.get("code", "").strip().upper()
        result = get_db().execute(
            "SELECT code, category, status, absent_from, absent_to, return_date "
            "FROM letters WHERE code=?", (searched,)).fetchone()
        if not result:
            flash("No letter found with that code.", "error")
    return render_template("track.html", result=result, searched=searched)


# ----------------------------------------------------------- ADMIN ROUTES ----
@app.route("/admin/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = request.form.get("username", "")
        pw = request.form.get("password", "")
        # compare_digest = safe comparison that resists timing attacks
        ok = (hmac.compare_digest(user, ADMIN_USERNAME)
              and hmac.compare_digest(pw, ADMIN_PASSWORD))
        if ok:
            session.clear()           # fresh session after login
            session["is_admin"] = True
            return redirect(url_for("admin_panel"))
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

    if action == "approve":
        db.execute("UPDATE letters SET status='approved' WHERE id=?", (letter_id,))
        # This is the "notify the student and the whole website" part:
        db.execute("INSERT INTO announcements (message, created_at) VALUES (?,?)",
                   (f'Excuse letter {row["code"]} has been APPROVED.', now()))
        flash(f'{row["code"]} approved and announced on the board.', "ok")
    elif action == "reject":
        db.execute("UPDATE letters SET status='rejected' WHERE id=?", (letter_id,))
        flash(f'{row["code"]} rejected.', "ok")
    elif action == "archive":
        db.execute("UPDATE letters SET archived=1 WHERE id=?", (letter_id,))
        flash(f'{row["code"]} archived (removed from the public board).', "ok")
    elif action == "restore":
        db.execute("UPDATE letters SET archived=0 WHERE id=?", (letter_id,))
        flash(f'{row["code"]} restored.', "ok")
    elif action == "delete":
        db.execute("DELETE FROM letters WHERE id=?", (letter_id,))
        flash(f'{row["code"]} permanently deleted.', "ok")
    else:
        abort(404)
    db.commit()
    return redirect(request.referrer or url_for("admin_panel"))


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
