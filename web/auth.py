"""Conturi utilizatori: parole (scrypt), login, protecție rute."""

import hashlib
import hmac
import secrets
import time
from collections import defaultdict, deque
from functools import wraps

from flask import abort, g, redirect, request, session, url_for

from . import config, db

# ---------------------------------------------------------------------------
# Parole
# ---------------------------------------------------------------------------

_SCRYPT = dict(n=2**14, r=8, p=1)


def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.scrypt(password.encode(), salt=salt, **_SCRYPT)
    return f"scrypt${salt.hex()}${digest.hex()}"


def check_password(password: str, stored: str) -> bool:
    try:
        _, salt_hex, digest_hex = stored.split("$")
    except ValueError:
        return False
    digest = hashlib.scrypt(password.encode(), salt=bytes.fromhex(salt_hex), **_SCRYPT)
    return hmac.compare_digest(digest.hex(), digest_hex)


def generate_password() -> str:
    return secrets.token_urlsafe(9)


# ---------------------------------------------------------------------------
# Utilizatori
# ---------------------------------------------------------------------------

def create_user(username: str, password: str, is_admin: bool = False):
    db.execute(
        "INSERT INTO users (username, pw_hash, is_admin, created_at) VALUES (?, ?, ?, ?)",
        (username, hash_password(password), int(is_admin), time.time()),
    )


def set_password(user_id: int, password: str):
    db.execute("UPDATE users SET pw_hash = ? WHERE id = ?", (hash_password(password), user_id))


def ensure_admin():
    """La prima pornire creează contul de administrator din ADMIN_USERNAME / ADMIN_PASSWORD."""
    if db.query("SELECT 1 FROM users LIMIT 1", one=True):
        return
    password = config.ADMIN_PASSWORD
    if not password:
        password = generate_password()
        print("=" * 60)
        print(f"  Cont administrator creat: {config.ADMIN_USERNAME}")
        print(f"  Parolă generată:          {password}")
        print("  (setați ADMIN_PASSWORD ca să o alegeți voi)")
        print("=" * 60, flush=True)
    create_user(config.ADMIN_USERNAME, password, is_admin=True)


def authenticate(username: str, password: str):
    user = db.query("SELECT * FROM users WHERE username = ?", (username,), one=True)
    if user and check_password(password, user["pw_hash"]):
        return user
    return None


def load_current_user():
    g.user = None
    user_id = session.get("user_id")
    if user_id is not None:
        g.user = db.query("SELECT * FROM users WHERE id = ?", (user_id,), one=True)
        if g.user is None:
            session.clear()


def login_required(view):
    @wraps(view)
    def wrapper(*args, **kwargs):
        if g.user is None:
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapper


def admin_required(view):
    @wraps(view)
    @login_required
    def wrapper(*args, **kwargs):
        if not g.user["is_admin"]:
            abort(403)
        return view(*args, **kwargs)
    return wrapper


# ---------------------------------------------------------------------------
# Limitare încercări de login (per IP, în memorie)
# ---------------------------------------------------------------------------

_FAILED_WINDOW = 15 * 60
_FAILED_MAX = 10
_failed: dict[str, deque] = defaultdict(deque)


def login_blocked(ip: str) -> bool:
    attempts = _failed[ip]
    while attempts and attempts[0] < time.time() - _FAILED_WINDOW:
        attempts.popleft()
    return len(attempts) >= _FAILED_MAX


def record_failed_login(ip: str):
    _failed[ip].append(time.time())


# ---------------------------------------------------------------------------
# CSRF
# ---------------------------------------------------------------------------

def csrf_token() -> str:
    if "csrf" not in session:
        session["csrf"] = secrets.token_hex(16)
    return session["csrf"]


def check_csrf():
    if request.method == "POST":
        sent = request.form.get("csrf") or request.headers.get("X-CSRF-Token", "")
        if not hmac.compare_digest(sent, session.get("csrf", "")):
            abort(400, "Token CSRF invalid — reîncărcați pagina.")
