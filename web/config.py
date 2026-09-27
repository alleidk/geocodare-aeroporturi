"""Configurare din variabile de mediu."""

import os
import secrets
from pathlib import Path

DATA_DIR = Path(os.environ.get("DATA_DIR", Path(__file__).resolve().parent.parent / "data"))
DB_PATH = DATA_DIR / "app.db"
JOBS_DIR = DATA_DIR / "jobs"

GOOGLE_API_KEY = os.environ.get("GOOGLE_API_KEY") or os.environ.get("GOOGLE_MAPS_API_KEY")
# Nominatim cere un User-Agent cu date de contact reale
NOMINATIM_EMAIL = os.environ.get("NOMINATIM_EMAIL", "")

# Primul cont de administrator (creat automat dacă nu există niciun utilizator)
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD")

# 1 = cookie de sesiune doar pe HTTPS (setați 1 în producție)
COOKIE_SECURE = os.environ.get("COOKIE_SECURE", "0") == "1"
# 1 = serverul rulează în spatele unui proxy (Railway, Fly.io, Render, nginx)
TRUST_PROXY = os.environ.get("TRUST_PROXY", "0") == "1"

MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "50"))
# După câte parole greșite (de pe același dispozitiv) se blochează login-ul și pentru cât timp
LOGIN_MAX_FAILURES = int(os.environ.get("LOGIN_MAX_FAILURES", "3"))
LOGIN_LOCKOUT_MINUTES = int(os.environ.get("LOGIN_LOCKOUT_MINUTES", "15"))
# Câte cereri Google în paralel (limita Google: 50/s)
GOOGLE_WORKERS = int(os.environ.get("GOOGLE_WORKERS", "4"))
# Termenii Google Maps Platform permit păstrarea coordonatelor max. 30 de zile
CACHE_TTL_DAYS = int(os.environ.get("CACHE_TTL_DAYS", "30"))

PORT = int(os.environ.get("PORT", "8000"))


def secret_key() -> str:
    """SECRET_KEY din mediu sau generat o dată și păstrat în DATA_DIR."""
    key = os.environ.get("SECRET_KEY")
    if key:
        return key
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    path = DATA_DIR / "secret_key"
    if not path.exists():
        path.write_text(secrets.token_hex(32), encoding="utf-8")
    return path.read_text(encoding="utf-8").strip()
