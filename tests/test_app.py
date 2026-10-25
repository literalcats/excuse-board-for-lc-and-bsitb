"""
Tests for the Excuse Letter Board. Run with:   python -m unittest -v
Each test proves one thing we fixed. Uses a throwaway SQLite file, never your real DB.
"""
import os, re, subprocess, sys, tempfile, unittest
from datetime import datetime, timedelta, timezone

# Must be set BEFORE importing the app.
_tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
os.environ["DATABASE_PATH"] = _tmp.name
os.environ.pop("DATABASE_URL", None)
os.environ.pop("RENDER", None)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as A  # noqa: E402

A.app.config["TESTING"] = False   # let our real error pages run (like in production)


def token(html):
    return re.search(r'name="csrf" value="(\w+)"', html).group(1)


def future(days):
    return (A.today() + timedelta(days=days)).isoformat()


class Base(unittest.TestCase):
    def setUp(self):
        A.reset_rate_limits()
        db = A.DB()
        db.execute("DELETE FROM letters"); db.execute("DELETE FROM announcements")
        db.execute("DELETE FROM sqlite_sequence")   # restart IDs at 1 so tests can use /letter/1
        db.commit(); db.close()
        self.c = A.app.test_client()

    def submit(self, **over):
        t = token(self.c.get("/request").text)
        form = dict(csrf=t, student_name="Juan Dela Cruz", guardian_name="Maria Dela Cruz",
                    addressed_to="Prof. Santos", category="Illness", reason="Fever",
                    absent_from=future(1), absent_to=future(2), return_date=future(3))
        form.update(over)
        return self.c.post("/request", data=form)

    def login(self):
        t = token(self.c.get("/admin/login").text)
        r = self.c.post("/admin/login", data=dict(csrf=t, username="admin", password="admin123"))
        self.assertEqual(r.status_code, 302)
        return token(self.c.get("/admin").text)


class SecurityTests(Base):
    def test_non_ascii_login_does_not_crash(self):
        t = token(self.c.get("/admin/login").text)
        r = self.c.post("/admin/login", data=dict(csrf=t, username="admin", password="pässwörd ñ"))
        self.assertEqual(r.status_code, 200)            # normal "wrong password" page, not 500

    def test_non_ascii_csrf_is_400_not_500(self):
        self.c.get("/admin/login")
        r = self.c.post("/admin/login", data=dict(csrf="é", username="a", password="b"))
        self.assertEqual(r.status_code, 400)

    def test_post_without_csrf_rejected(self):
        self.login()
        self.assertEqual(self.c.post("/admin/letter/1/delete", data={}).status_code, 400)

    def test_admin_requires_login(self):
        for path in ["/admin", "/admin/letter/1", "/admin/letter/1/download"]:
            self.assertEqual(self.c.get(path).status_code, 302, path)

    def test_security_headers(self):
        h = self.c.get("/").headers
        self.assertIn("script-src 'self'", h["Content-Security-Policy"])
        self.assertEqual(h["X-Frame-Options"], "DENY")
        self.assertEqual(h["X-Content-Type-Options"], "nosniff")
        self.login()
        self.assertEqual(self.c.get("/admin").headers["Cache-Control"], "no-store")

    def test_cookie_flags(self):
        self.assertTrue(A.app.config["SESSION_COOKIE_HTTPONLY"])
        self.assertEqual(A.app.config["SESSION_COOKIE_SAMESITE"], "Lax")

    def test_no_inline_scripts_so_csp_does_not_break_the_site(self):
        self.login()
        for path in ["/", "/request", "/track", "/admin/login", "/admin", "/offline"]:
            html = self.c.get(path).text
            for m in re.finditer(r"<script(?![^>]*\bsrc=)[^>]*>", html):
                self.fail(f"inline <script> on {path}")
            self.assertNotRegex(html, r"\son(click|submit|load)=", path)
            self.assertNotIn(' style="', html, path)

    def test_login_rate_limit(self):
        t = token(self.c.get("/admin/login").text)
        codes = [self.c.post("/admin/login", data=dict(csrf=t, username="admin", password=f"x{i}")).status_code
                 for i in range(12)]
        self.assertIn(429, codes)
        self.assertEqual(codes[0], 200)

    def test_prod_refuses_unsafe_config(self):
        env = {k: v for k, v in os.environ.items() if k not in ("SECRET_KEY", "ADMIN_PASSWORD", "DATABASE_URL")}
        env["RENDER"] = "true"
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        p = subprocess.run([sys.executable, "-c", "import app"], cwd=root, env=env,
                           capture_output=True, text=True)
        self.assertNotEqual(p.returncode, 0)
        self.assertIn("Unsafe production config", p.stderr)


class PrivacyTests(Base):
    def test_board_never_shows_names_or_reason(self):
        self.submit(student_name="Secretname Person", guardian_name="Hiddenguardian", reason="Privatereason")
        html = self.c.get("/").text
        for secret in ("Secretname", "Hiddenguardian", "Privatereason"):
            self.assertNotIn(secret, html)
        self.assertRegex(html, r"EX-[A-Z2-9]{5}")

    def test_track_never_shows_names_or_reason(self):
        r = self.submit(student_name="Secretname Person", reason="Privatereason")
        code = r.headers["Location"].split("/")[-1]
        t = token(self.c.get("/track").text)
        html = self.c.post("/track", data=dict(csrf=t, code=code)).text
        self.assertIn(code, html)
        self.assertNotIn("Secretname", html); self.assertNotIn("Privatereason", html)

    def test_submitted_page_rejects_arbitrary_text(self):
        self.assertEqual(self.c.get("/submitted/<b>hi</b>").status_code, 404)


class InputTests(Base):
    def test_junk_dates_rejected(self):
        r = self.submit(absent_from="zzzz", absent_to="zzzz", return_date="zzzzz")
        self.assertEqual(r.status_code, 200)            # form shown again, nothing saved
        self.assertIn("valid date", r.text)

    def test_impossible_date_rejected(self):
        r = self.submit(absent_from="2026-02-31")
        self.assertEqual(r.status_code, 200)

    def test_return_before_absence_rejected(self):
        r = self.submit(return_date=future(1))
        self.assertEqual(r.status_code, 200)

    def test_control_characters_do_not_break_docx(self):
        self.submit(reason="fever\x0bcough\x00 end", student_name="Ju\x07an")
        self.login()
        r = self.c.get("/admin/letter/1/download")
        self.assertEqual(r.status_code, 200)
        self.assertGreater(len(r.data), 1000)

    def test_old_dirty_rows_still_download(self):
        """Rows saved BEFORE the fix may hold control characters already."""
        db = A.DB()
        db.execute("INSERT INTO letters (code,student_name,guardian_name,addressed_to,category,reason,"
                   "absent_from,absent_to,return_date,created_at) VALUES (?,?,?,?,?,?,?,?,?,?)",
                   ("EX-AAAAA", "Old\x0bName", "G", "P", "Illness", "bad\x00reason",
                    future(1), future(2), future(3), A.now()))
        db.commit(); db.close()
        self.login()
        self.assertEqual(self.c.get("/admin/letter/1/download").status_code, 200)

    def test_html_in_input_is_escaped(self):
        self.submit(student_name="<script>alert(1)</script>")
        self.login()
        html = self.c.get("/admin").text
        self.assertNotIn("<script>alert(1)</script>", html)
        self.assertIn("&lt;script&gt;", html)


class LogicTests(Base):
    def make_and_login(self):
        self.submit()
        return self.login()

    def count_notices(self):
        return self.c.get("/").text.count("has been APPROVED")

    def test_double_approve_makes_one_announcement(self):
        t = self.make_and_login()
        for _ in range(3):
            self.c.post("/admin/letter/1/approve", data=dict(csrf=t))
        self.assertEqual(self.count_notices(), 1)

    def test_reject_removes_approved_notice(self):
        t = self.make_and_login()
        self.c.post("/admin/letter/1/approve", data=dict(csrf=t))
        self.assertEqual(self.count_notices(), 1)
        self.c.post("/admin/letter/1/reject", data=dict(csrf=t))
        self.assertEqual(self.count_notices(), 0)

    def test_delete_removes_notice_and_letter(self):
        t = self.make_and_login()
        self.c.post("/admin/letter/1/approve", data=dict(csrf=t))
        self.c.post("/admin/letter/1/delete", data=dict(csrf=t))
        self.assertEqual(self.count_notices(), 0)
        self.assertEqual(self.c.get("/admin/letter/1").status_code, 404)

    def test_archive_hides_from_board_and_restore_returns(self):
        t = self.make_and_login()
        code = A.DB().execute("SELECT code FROM letters").fetchone()["code"]
        self.c.post("/admin/letter/1/archive", data=dict(csrf=t, view="active"))
        card = f'class="code">{code}'                 # the board card (not a flash message)
        self.assertNotIn(card, self.c.get("/").text)
        self.assertIn(code, self.c.get("/admin?view=archived").text)
        r = self.c.post("/admin/letter/1/restore", data=dict(csrf=t, view="archived"))
        self.assertIn("view=archived", r.headers["Location"])
        self.assertIn(card, self.c.get("/").text)

    def test_timestamps_are_manila_time(self):
        manila = datetime.now(timezone(timedelta(hours=8))).replace(tzinfo=None)
        stamped = datetime.strptime(A.now(), "%Y-%m-%d %H:%M")
        self.assertLess(abs((manila - stamped).total_seconds()), 120)


class PwaTests(Base):
    def test_offline_page_has_no_session_data(self):
        self.login()                                   # even while an admin is logged in...
        html = self.c.get("/offline").text
        self.assertNotIn("Log out", html)
        self.assertNotIn('name="csrf"', html)
        self.assertNotIn("/admin", html)

    def test_manifest_and_sw(self):
        m = self.c.get("/manifest.webmanifest").get_json()
        self.assertEqual(m["display"], "standalone"); self.assertEqual(m["id"], "/")
        self.assertEqual(self.c.get("/sw.js").status_code, 200)
        self.assertIn("res.ok", self.c.get("/sw.js").text)

    def test_friendly_404(self):
        r = self.c.get("/nope")
        self.assertEqual(r.status_code, 404); self.assertIn("Back to the board", r.text)


if __name__ == "__main__":
    unittest.main()
