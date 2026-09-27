"""Site privat pentru geocodarea fișierelor Excel Origine–Destinație."""

import json
import shutil
import time
import uuid
from pathlib import Path

import openpyxl
from flask import (
    Flask, abort, flash, g, jsonify, redirect, render_template, request,
    send_file, session, url_for,
)
from werkzeug.middleware.proxy_fix import ProxyFix

from . import auth, config, db, geocoding, worker

STATUS_LABELS = {
    "new": "De configurat",
    "queued": "În așteptare",
    "running": "În lucru",
    "done": "Gata",
    "failed": "Eroare",
    "cancelled": "Oprit",
}


def create_app() -> Flask:
    app = Flask(__name__)
    app.config.update(
        SECRET_KEY=config.secret_key(),
        MAX_CONTENT_LENGTH=config.MAX_UPLOAD_MB * 1024 * 1024,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_COOKIE_SECURE=config.COOKIE_SECURE,
        PERMANENT_SESSION_LIFETIME=30 * 86400,
    )

    if config.TRUST_PROXY:
        # Pe Railway/Fly/Render cererile vin printr-un proxy; IP-ul real e în X-Forwarded-For
        app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

    db.init()
    auth.ensure_admin()
    worker.start()

    app.before_request(auth.load_current_user)
    app.before_request(auth.check_csrf)

    @app.context_processor
    def inject():
        return {
            "csrf_token": auth.csrf_token,
            "status_labels": STATUS_LABELS,
            "providers": geocoding.PROVIDERS,
            "max_upload_mb": config.MAX_UPLOAD_MB,
        }

    @app.template_filter("datetime")
    def fmt_datetime(ts):
        return time.strftime("%d.%m.%Y %H:%M", time.localtime(ts)) if ts else ""

    @app.after_request
    def security_headers(resp):
        resp.headers.setdefault("X-Frame-Options", "DENY")
        resp.headers.setdefault("X-Content-Type-Options", "nosniff")
        resp.headers.setdefault("Referrer-Policy", "same-origin")
        return resp

    @app.errorhandler(413)
    def too_large(_):
        flash(f"Fișierul depășește {config.MAX_UPLOAD_MB} MB.", "error")
        return redirect(url_for("index"))

    register_routes(app)
    return app


def get_job(job_id: str):
    job = db.query("SELECT * FROM jobs WHERE id = ?", (job_id,), one=True)
    if job is None or (job["user_id"] != g.user["id"] and not g.user["is_admin"]):
        abort(404)
    return job


def job_json(job) -> dict:
    return {
        "id": job["id"],
        "status": job["status"],
        "status_label": STATUS_LABELS[job["status"]],
        "phase": job["phase"],
        "rows_count": job["rows_count"],
        "already_done": job["already_done"],
        "unique_count": job["unique_count"],
        "cache_hits": job["cache_hits"],
        "total": job["total"],
        "done": job["done"],
        "errors": job["errors"],
        "addresses": job["addresses"],
        "found": job["found"],
        "message": job["message"],
    }


def describe(job) -> str:
    """Ce se geocodează, pe scurt: „origine + destinație · 37 grupuri de coloane”."""
    parts = []
    if job["col_origin"]:
        parts.append("origine + destinație" if job["col_dest"] else "o coloană de adrese")
    groups = len(json.loads(job["groups"] or "[]"))
    if groups:
        parts.append(f"{groups} grupuri de coloane")
    return " · ".join(parts)


def parse_config(form) -> dict:
    """Setările din pagina de configurare. ValueError cu mesaj pentru utilizator."""
    def opt_int(name):
        v = form.get(name, "").strip()
        return int(v) if v else None

    try:
        cfg = {
            "sheet": form["sheet"],
            "header_row": int(form["header_row"]),
            "col_origin": opt_int("col_origin"),
            "col_dest": opt_int("col_dest"),
            "max_rows": opt_int("max_rows"),
        }
    except (KeyError, ValueError):
        raise ValueError("Setări invalide — reîncărcați pagina.")
    try:
        cfg["groups"] = geocoding.parse_groups(form.get("groups", ""))
    except ValueError:
        raise ValueError("Coloanele grupurilor trebuie scrise ca litere Excel, ex.: GY, HO, ID.")
    cfg["lookup_sheet"] = form.get("lookup_sheet") or None
    cfg["skip_done"] = bool(form.get("skip_done"))
    if cfg["col_dest"] and not cfg["col_origin"]:
        cfg["col_origin"], cfg["col_dest"] = cfg["col_dest"], None
    if not cfg["col_origin"] and not cfg["groups"]:
        raise ValueError("Alegeți cel puțin o coloană de adrese sau un grup de coloane.")
    if cfg["col_origin"] and cfg["col_origin"] == cfg["col_dest"]:
        raise ValueError("Coloanele de origine și destinație trebuie să fie diferite.")
    return cfg


def register_routes(app: Flask):

    # ------------------------------------------------------------------ login

    @app.route("/login", methods=["GET", "POST"])
    def login():
        if request.method == "POST":
            ip = request.remote_addr or "?"
            username = request.form.get("username", "").strip()
            wait = auth.login_wait_seconds(ip, username)
            if wait:
                flash(f"Prea multe parole greșite. Poți încerca din nou peste {-(-wait // 60)} minute.", "error")
                return render_template("login.html"), 429
            user = auth.authenticate(username, request.form.get("password", ""))
            if user is None:
                left = auth.record_failed_login(ip, username)
                if left:
                    flash(f"Utilizator sau parolă greșită. Mai ai {left} "
                          f"{'încercare' if left == 1 else 'încercări'}, apoi accesul se blochează "
                          f"{config.LOGIN_LOCKOUT_MINUTES} minute.", "error")
                else:
                    flash(f"Prea multe parole greșite. Accesul e blocat {config.LOGIN_LOCKOUT_MINUTES} minute.", "error")
                return render_template("login.html"), 401
            auth.clear_failed_logins(ip, username)
            session.clear()
            session.permanent = True
            session["user_id"] = user["id"]
            nxt = request.args.get("next", "")
            return redirect(nxt if nxt.startswith("/") and not nxt.startswith("//") else url_for("index"))
        return render_template("login.html")

    @app.post("/logout")
    def logout():
        session.clear()
        return redirect(url_for("login"))

    @app.route("/account", methods=["GET", "POST"])
    @auth.login_required
    def account():
        if request.method == "POST":
            current = request.form.get("current", "")
            new = request.form.get("new", "")
            if not auth.check_password(current, g.user["pw_hash"]):
                flash("Parola actuală este greșită.", "error")
            elif len(new) < 8:
                flash("Parola nouă trebuie să aibă cel puțin 8 caractere.", "error")
            else:
                auth.set_password(g.user["id"], new)
                flash("Parola a fost schimbată.", "ok")
                return redirect(url_for("index"))
        return render_template("account.html")

    # ------------------------------------------------------------------ joburi

    @app.get("/")
    @auth.login_required
    def index():
        if g.user["is_admin"]:
            jobs = db.query("""SELECT jobs.*, users.username FROM jobs JOIN users ON users.id = jobs.user_id
                               ORDER BY created_at DESC LIMIT 100""")
        else:
            jobs = db.query("""SELECT jobs.*, users.username FROM jobs JOIN users ON users.id = jobs.user_id
                               WHERE user_id = ? ORDER BY created_at DESC LIMIT 100""", (g.user["id"],))
        return render_template("index.html", jobs=jobs)

    @app.post("/upload")
    @auth.login_required
    def upload():
        f = request.files.get("file")
        if not f or not f.filename:
            flash("Alegeți un fișier .xlsx.", "error")
            return redirect(url_for("index"))
        filename = Path(f.filename.replace("\\", "/")).name
        if not filename.lower().endswith(".xlsx"):
            flash("Sunt acceptate doar fișiere .xlsx.", "error")
            return redirect(url_for("index"))

        job_id = uuid.uuid4().hex[:12]
        path = worker.input_path(job_id)
        path.parent.mkdir(parents=True)
        f.save(path)
        try:
            openpyxl.load_workbook(path, read_only=True).close()
        except Exception:
            shutil.rmtree(path.parent, ignore_errors=True)
            flash("Fișierul nu a putut fi citit ca Excel (.xlsx).", "error")
            return redirect(url_for("index"))

        now = time.time()
        db.execute(
            """INSERT INTO jobs (id, user_id, filename, status, created_at, updated_at, output_name)
               VALUES (?, ?, ?, 'new', ?, ?, ?)""",
            (job_id, g.user["id"], filename, now, now, f"{Path(filename).stem}Codat.xlsx"),
        )
        return redirect(url_for("configure", job_id=job_id))

    @app.route("/jobs/<job_id>/configure", methods=["GET", "POST"])
    @auth.login_required
    def configure(job_id):
        job = get_job(job_id)
        if job["status"] not in ("new", "done", "failed", "cancelled"):
            return redirect(url_for("job_page", job_id=job_id))
        providers = geocoding.available_providers()

        if request.method == "POST":
            form = request.form
            try:
                cfg = parse_config(form)
            except ValueError as e:
                flash(str(e), "error")
                return redirect(url_for("configure", job_id=job_id))
            if form.get("provider") not in providers:
                abort(400)
            db.update_job(job_id, status="queued", sheet=cfg["sheet"], header_row=cfg["header_row"],
                          col_origin=cfg["col_origin"], col_dest=cfg["col_dest"],
                          groups=json.dumps(cfg["groups"]), lookup_sheet=cfg["lookup_sheet"],
                          provider=form["provider"], max_rows=cfg["max_rows"], skip_done=int(cfg["skip_done"]),
                          phase=None, message=None, done=0, total=0, cancel_requested=0)
            worker.notify()
            return redirect(url_for("job_page", job_id=job_id))

        sheet = request.args.get("sheet") or job["sheet"]
        header_row = request.args.get("header_row", type=int) or (
            job["header_row"] if sheet == job["sheet"] else None)
        info = geocoding.inspect_sheet(worker.input_path(job_id), sheet, header_row)
        same = job["status"] != "new" and sheet == job["sheet"]
        # Fără grupuri salvate (inclusiv joburile configurate înainte de grupuri) → propunem grupurile detectate
        same_groups = same and bool(json.loads(job["groups"] or "[]"))
        selected = {
            "col_origin": job["col_origin"] if same else info["guess_origin"],
            "col_dest": job["col_dest"] if same else info["guess_dest"],
            "groups": geocoding.format_groups(json.loads(job["groups"]) if same_groups else info["guess_groups"]),
            "lookup_sheet": job["lookup_sheet"] if same_groups else info["guess_lookup"],
        }
        return render_template("configure.html", job=job, info=info, available=providers, selected=selected)

    @app.get("/jobs/<job_id>/estimate")
    @auth.login_required
    def estimate(job_id):
        get_job(job_id)
        try:
            cfg = parse_config(request.args)
            result = geocoding.estimate(worker.input_path(job_id), cfg["sheet"], cfg["header_row"], cfg,
                                        request.args["provider"], cfg["max_rows"], cfg["skip_done"])
        except (KeyError, ValueError) as e:
            return jsonify({"error": str(e)}), 400
        return jsonify(result)

    @app.get("/jobs/<job_id>")
    @auth.login_required
    def job_page(job_id):
        job = get_job(job_id)
        if job["status"] == "new":
            return redirect(url_for("configure", job_id=job_id))
        breakdown = geocoding.error_breakdown(job["message"]) if job["status"] == "done" else []
        return render_template("job.html", job=job, job_data=job_json(job), what=describe(job),
                               breakdown=breakdown)

    @app.get("/ajutor")
    @auth.login_required
    def help_page():
        return render_template("help.html", error_help=geocoding.ERROR_HELP)

    @app.get("/api/jobs/<job_id>")
    @auth.login_required
    def job_status(job_id):
        return jsonify(job_json(get_job(job_id)))

    @app.post("/jobs/<job_id>/cancel")
    @auth.login_required
    def cancel(job_id):
        job = get_job(job_id)
        if job["status"] == "queued":
            db.update_job(job_id, status="cancelled", message="Oprit înainte de pornire.")
        elif job["status"] == "running":
            db.update_job(job_id, cancel_requested=1)
        return redirect(url_for("job_page", job_id=job_id))

    @app.post("/jobs/<job_id>/retry")
    @auth.login_required
    def retry(job_id):
        job = get_job(job_id)
        if job["status"] in ("failed", "cancelled", "done"):
            db.update_job(job_id, status="queued", message=None, phase=None, cancel_requested=0)
            worker.notify()
        return redirect(url_for("job_page", job_id=job_id))

    @app.post("/jobs/<job_id>/delete")
    @auth.login_required
    def delete(job_id):
        job = get_job(job_id)
        if job["status"] == "running":
            flash("Opriți jobul înainte de a-l șterge.", "error")
            return redirect(url_for("job_page", job_id=job_id))
        db.execute("DELETE FROM jobs WHERE id = ?", (job_id,))
        shutil.rmtree(worker.job_dir(job_id), ignore_errors=True)
        flash(f"„{job['filename']}” a fost șters.", "ok")
        return redirect(url_for("index"))

    @app.get("/jobs/<job_id>/download")
    @auth.login_required
    def download(job_id):
        job = get_job(job_id)
        path = worker.output_path(job_id)
        if job["status"] != "done" or not path.exists():
            abort(404)
        return send_file(path, as_attachment=True, download_name=job["output_name"])

    # ------------------------------------------------------------------ admin

    @app.get("/admin/users")
    @auth.admin_required
    def admin_users():
        users = db.query("""SELECT users.*, COUNT(jobs.id) AS job_count FROM users
                            LEFT JOIN jobs ON jobs.user_id = users.id
                            GROUP BY users.id ORDER BY users.created_at""")
        return render_template("admin.html", users=users, new_password=session.pop("new_password", None))

    @app.post("/admin/users")
    @auth.admin_required
    def admin_create_user():
        username = request.form.get("username", "").strip()
        if not username or len(username) > 50:
            flash("Numele de utilizator este invalid.", "error")
            return redirect(url_for("admin_users"))
        if db.query("SELECT 1 FROM users WHERE username = ?", (username,), one=True):
            flash(f"Utilizatorul „{username}” există deja.", "error")
            return redirect(url_for("admin_users"))
        password = auth.generate_password()
        auth.create_user(username, password, is_admin=bool(request.form.get("is_admin")))
        session["new_password"] = {"username": username, "password": password}
        return redirect(url_for("admin_users"))

    @app.post("/admin/users/<int:user_id>/reset")
    @auth.admin_required
    def admin_reset_password(user_id):
        user = db.query("SELECT * FROM users WHERE id = ?", (user_id,), one=True) or abort(404)
        password = auth.generate_password()
        auth.set_password(user_id, password)
        session["new_password"] = {"username": user["username"], "password": password}
        return redirect(url_for("admin_users"))

    @app.post("/admin/users/<int:user_id>/delete")
    @auth.admin_required
    def admin_delete_user(user_id):
        if user_id == g.user["id"]:
            flash("Nu vă puteți șterge propriul cont.", "error")
            return redirect(url_for("admin_users"))
        if db.query("SELECT 1 FROM jobs WHERE user_id = ? AND status IN ('queued', 'running')",
                    (user_id,), one=True):
            flash("Utilizatorul are joburi în lucru — opriți-le întâi.", "error")
            return redirect(url_for("admin_users"))
        for job in db.query("SELECT id FROM jobs WHERE user_id = ?", (user_id,)):
            shutil.rmtree(worker.job_dir(job["id"]), ignore_errors=True)
        db.execute("DELETE FROM users WHERE id = ?", (user_id,))
        flash("Utilizator șters.", "ok")
        return redirect(url_for("admin_users"))
