# Excuse Letter Board (prototype)

A small standalone Flask app. Part of a bigger super-app, built here as a test version.

## What it does
- **Viewers** request an excuse letter and get an anonymous code like `EX-7K3QA`.
- **Public board** shows only the code, category, dates and status. Names and reasons stay private.
- **Admin** (login required) can view, approve, reject, download as `.docx`, archive, restore and delete.
- **Approving** posts an announcement on the board, so students and everyone else see it.
- Dark mode by default (dark gray + forest green). Light mode uses Google-style colors. A button toggles it.

## Project layout
```
app.py            all the Python: routes, database, letter template, login
templates/        HTML pages (base.html is the shared layout)
static/           style.css (colors) and theme.js (dark/light toggle)
requirements.txt  Python packages Render installs
render.yaml       Render deploy settings
```

## Run on your laptop
```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```
Open http://127.0.0.1:5000. Local admin login is `admin` / `admin123` (local only).

## Deploy on Render
1. Push this folder to a GitHub repo.
2. On Render: **New > Blueprint**, pick the repo. It reads `render.yaml`.
3. After deploy, open your service > **Environment** and copy `ADMIN_PASSWORD` (Render generated it). Username is `admin`.
4. Visit `https://your-app.onrender.com/admin/login`.

(Manual alternative: New > Web Service, build `pip install -r requirements.txt`, start `gunicorn app:app`, and add env vars `SECRET_KEY`, `ADMIN_USERNAME`, `ADMIN_PASSWORD`.)

## Things to know
- **Database:** on Render the Blueprint creates a PostgreSQL database and sets `DATABASE_URL` for you, so data survives redeploys. On your laptop (no `DATABASE_URL`) it falls back to a local SQLite file. Check Render's current free-tier limits for Postgres; free databases may expire.
- **The free plan sleeps** after inactivity, so the first visit can take ~50 seconds.
- **Change the admin password** by editing the `ADMIN_PASSWORD` env var on Render.
- **Edit the letter wording** in `build_letter()` in `app.py`. Preview and `.docx` both update.
- **Edit colors** at the top of `static/style.css`.

## Ideas for phase 2
Student accounts instead of codes, real file/PDF export, email notifications, rate limiting on the request form, 
## Security notes (added in the hardening pass)
- On Render the app **refuses to start** if `SECRET_KEY`, a strong `ADMIN_PASSWORD`, or `DATABASE_URL` is missing.
- Login, request and track forms are rate limited (edit `LIMITS` in `app.py`). The limiter is in-memory per server process.
- Browser protections: Content-Security-Policy (no inline scripts), clickjacking block, secure cookies, no-store on admin pages.
- Times are saved in Manila time (UTC+8). Rows saved before this update are in UTC.

## Run the tests
```bash
python -m unittest -v
```
