"""Conturi utilizatori: parole (scrypt), login, protecție rute."""

import hashlib
import hmac
import secrets
import threading
import time
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
# Blocare după parole greșite (în memorie)
#   - de pe același dispozitiv (IP): după LOGIN_MAX_FAILURES greșeli
#   - pe același cont, din orice loc: după USER_MAX_FAILURES greșeli
# Blocarea durează LOGIN_LOCKOUT_MINUTES; greșelile se uită după aceeași durată.
# ---------------------------------------------------------------------------

USER_MAX_FAILURES = 10

_lock = threading.Lock()
_failures: dict[str, list[float]] = {}
_locked_until: dict[str, float] = {}


def _keys(ip: str, username: str):
    return (f"ip:{ip}", config.LOGIN_MAX_FAILURES), (f"user:{username.lower()}", USER_MAX_FAILURES)


def login_wait_seconds(ip: str, username: str) -> int:
    """Câte secunde mai durează blocarea (0 = se poate încerca)."""
    now = time.time()
    with _lock:
        wait = 0
        for key, _ in _keys(ip, username):
            until = _locked_until.get(key, 0)
            if until > now:
                wait = max(wait, int(until - now) + 1)
            else:
                _locked_until.pop(key, None)
        return wait


def record_failed_login(ip: str, username: str) -> int:
    """Notează o greșeală. Returnează câte încercări mai are dispozitivul până la blocare."""
    now = time.time()
    lockout = config.LOGIN_LOCKOUT_MINUTES * 60
    remaining = config.LOGIN_MAX_FAILURES
    with _lock:
        for key, limit in _keys(ip, username):
            times = [t for t in _failures.get(key, []) if now - t < lockout] + [now]
            if len(times) >= limit:
                _locked_until[key] = now + lockout
                times = []
            _failures[key] = times
            if key.startswith("ip:"):
                remaining = limit - len(times) if times else 0
    return remaining


def clear_failed_logins(ip: str, username: str):
    with _lock:
        for key, _ in _keys(ip, username):
            _failures.pop(key, None)


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
