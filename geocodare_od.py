#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sistem de geocodare Origine–Destinație pentru fișiere Excel (.xlsx)
Versiune: 1.0 — August 2026

Funcționalități:
  - File picker grafic (Tkinter)
  - Normalizare adrese românești (abrevieri, diacritice)
  - Detectare adrese ambigue
  - Geocodare Google Geocoding API cu retry exponențial
  - Caching persistent (JSON) + memorie
  - Checkpointing la fiecare 500 rânduri
  - Logging complet (DEBUG) cu răspunsuri Google
  - Dry run implicit (100 rânduri) și procesare parțială (--start/--end)

Utilizare:
  python geocodare_od.py                        # file picker + dry run 100 rânduri
  python geocodare_od.py --dry-run 0            # procesare completă
  python geocodare_od.py --start 100 --end 500  # procesare parțială
"""

import argparse
import json
import logging
import os
import re
import sys
import time
from pathlib import Path

import openpyxl
import requests

try:
    from tqdm import tqdm
except ImportError:
    tqdm = None

# ---------------------------------------------------------------------------
# 1. CONSTANTE
# ---------------------------------------------------------------------------
GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json"
RATE_LIMIT_PAUSE = 0.15        # secunde între cereri
RETRY_COUNT = 3
RETRY_BACKOFF_BASE = 1         # 1s → 2s → 4s
TIMEOUT_SECONDS = 10
CHECKPOINT_EVERY = 500
DEFAULT_DRY_RUN = 100

# Coloane output – Origine
COLS_ORIGINE = [
    "DeCodatOriginea",
    "LatOriginea",
    "LonOriginea",
    "PlaceIdOriginea",
    "ErrorOriginea",
]
# Coloane output – Destinație
COLS_DESTINATIE = [
    "DeCodatDestinatia",
    "LatDestinatia",
    "LonDestinatia",
    "PlaceIdDestinatia",
    "ErrorDestinatia",
]

# ---------------------------------------------------------------------------
# 2. NORMALIZARE ADRESE
# ---------------------------------------------------------------------------

# Abrevieri → extensii (word boundary, case-insensitive)
ABBREVIATIONS = [
    (r'\bstr\.?\s',        'Strada '),
    (r'\bbd\b',            'Bulevardul'),
    (r'\bb-d\b',           'Bulevardul'),
    (r'\bbulev\.?\b',      'Bulevardul'),
    (r'\bnr(?!\.)\b',      'nr.'),
    (r'\bapt\b',           'ap.'),
    (r'\bbl\b',            'bl.'),
    (r'\bpl\b',            'Piața'),
    (r'\bpiata\b',         'Piața'),
    (r'\bjud\.?\b',        'Județul'),
    (r'\bjudet\b',         'Județul'),
    (r'\bcom\b',           'Comuna'),
    (r'\bal\b',            'Aleea'),
    (r'\bcal\b',           'Calea'),
    (r'\bsos\b',           'Șoseaua'),
    (r'\bfn\b',            'f.n.'),
    (r'\bsc\b',            'sc.'),
    (r'\bet\b',            'et.'),
    (r'\bmun\b',           'Municipiul'),
    (r'\bsat\b',           'Satul'),
    (r'\boras\b',          'Orașul'),
    (r'\bsf\b',            'Sfântu'),
]

# Diacritice – localități frecvente din România
DIACRITICS_MAP = {
    'timisoara':        'Timișoara',
    'cluj napoca':      'Cluj-Napoca',
    'cluj-napoca':      'Cluj-Napoca',
    'targu mures':      'Târgu Mureș',
    'targu-mures':      'Târgu Mureș',
    'tg mures':         'Târgu Mureș',
    'tg. mures':        'Târgu Mureș',
    'tg.mures':         'Târgu Mureș',
    'bucuresti':        'București',
    'iasi':             'Iași',
    'brasov':           'Brașov',
    'constanta':        'Constanța',
    'craiova':          'Craiova',
    'galati':           'Galați',
    'ploiesti':         'Ploiești',
    'oradea':           'Oradea',
    'pitesti':          'Pitești',
    'sibiu':            'Sibiu',
    'bacau':            'Bacău',
    'arad':             'Arad',
    'buzau':            'Buzău',
    'baia mare':        'Baia Mare',
    'suceava':          'Suceava',
    'botosani':         'Botoșani',
    'satu mare':        'Satu Mare',
    'ramnicu valcea':   'Râmnicu Vâlcea',
    'rm valcea':        'Râmnicu Vâlcea',
    'rm. valcea':       'Râmnicu Vâlcea',
    'deva':             'Deva',
    'alba iulia':       'Alba Iulia',
    'bistrita':         'Bistrița',
    'piatra neamt':     'Piatra Neamț',
    'targoviste':       'Târgoviște',
    'resita':           'Reșița',
    'focsani':          'Focșani',
    'miercurea ciuc':   'Miercurea Ciuc',
    'sfantu gheorghe':  'Sfântu Gheorghe',
    'sf gheorghe':      'Sfântu Gheorghe',
    'sf. gheorghe':     'Sfântu Gheorghe',
    'drobeta turnu severin': 'Drobeta-Turnu Severin',
    'zalau':            'Zalău',
    'giurgiu':          'Giurgiu',
    'targu jiu':        'Târgu Jiu',
    'calarasi':         'Călărași',
    'alexandria':       'Alexandria',
    'slobozia':         'Slobozia',
    'vaslui':           'Vaslui',
    'slatina':          'Slatina',
    'tulcea':           'Tulcea',
    'hunedoara':        'Hunedoara',
    'medias':           'Mediaș',
    'lugoj':            'Lugoj',
    'dej':              'Dej',
    'turda':            'Turda',
    'sighisoara':       'Sighișoara',
    'campina':          'Câmpina',
    'campulung':        'Câmpulung',
    'reghin':           'Reghin',
    'fagaras':          'Făgăraș',
    'petrosani':        'Petroșani',
    'odorheiu secuiesc': 'Odorheiu Secuiesc',
    'gheorgheni':       'Gheorgheni',
    'mangalia':         'Mangalia',
    'navodari':         'Năvodari',
    'pascani':          'Pașcani',
    'roman':            'Roman',
    'tecuci':           'Tecuci',
    'voluntari':        'Voluntari',
    'pantelimon':       'Pantelimon',
    'otopeni':          'Otopeni',
    'bragadiru':        'Bragadiru',
    'popesti leordeni': 'Popești-Leordeni',
    'popesti-leordeni': 'Popești-Leordeni',
    'chiajna':          'Chiajna',
    'floresti':         'Florești',
    'baicoi':           'Băicoi',
    'cugir':            'Cugir',
    'sebes':            'Sebeș',
    'aiud':             'Aiud',
    'blaj':             'Blaj',
    'cisnadie':         'Cisnădie',
}

# Regex compilate pentru abrevieri
_ABBREV_COMPILED = [(re.compile(pat, re.IGNORECASE), repl) for pat, repl in ABBREVIATIONS]

# Pattern caractere redundante
_REDUNDANT_CHARS = re.compile(r'[/]{2,}|[*]{2,}')


def normalize_address(raw: str) -> str:
    """Normalizează o adresă conform regulilor stabilite."""
    if not raw or not isinstance(raw, str):
        return ""

    text = raw.strip()

    # Elimină caractere redundante (///, ***, //)
    text = _REDUNDANT_CHARS.sub(' ', text)

    # Elimină spații multiple
    text = ' '.join(text.split())

    # Extindere abrevieri
    for pat, repl in _ABBREV_COMPILED:
        text = pat.sub(repl, text)

    # Corectare diacritice – localități (caută în text, case-insensitive)
    text_lower = text.lower()
    for key, correct in sorted(DIACRITICS_MAP.items(), key=lambda x: -len(x[0])):
        # Caută cuvântul în text (word boundary mai relaxat pentru localități)
        pattern = re.compile(re.escape(key), re.IGNORECASE)
        if pattern.search(text_lower):
            text = pattern.sub(correct, text)
            text_lower = text.lower()  # re-compute

    # Elimină spații duble reziduale
    text = ' '.join(text.split())

    return text.strip()


# ---------------------------------------------------------------------------
# 3. DETECTARE AMBIGUITATE
# ---------------------------------------------------------------------------

_AMBIGUOUS_PATTERNS = [
    re.compile(r'\bPO\s*Box\b', re.IGNORECASE),
    re.compile(r'\bcasut[aă]\s*po[sș]tal[aă]\b', re.IGNORECASE),
    re.compile(r'\bcol[tț]\b', re.IGNORECASE),
    re.compile(r'\bcorner\b', re.IGNORECASE),
    re.compile(r'\binterse[cC][tț]i[ea]\b', re.IGNORECASE),
]

# Delimitatori care sugerează adrese multiple
_MULTI_ADDR_PATTERN = re.compile(r'[;]|\band\b', re.IGNORECASE)
# 'și' between two street references → ambiguous
_MULTI_ADDR_SI_PATTERN = re.compile(
    r'\b(?:str|strada|bd|bulevardul)\b.*\bși\b.*\b(?:str|strada|bd|bulevardul)\b',
    re.IGNORECASE,
)


def is_ambiguous(address: str) -> bool:
    """Returnează True dacă adresa este considerată ambiguă."""
    if not address:
        return False

    for pat in _AMBIGUOUS_PATTERNS:
        if pat.search(address):
            return True

    # Două adrese separate prin ; / and / și
    if _MULTI_ADDR_PATTERN.search(address):
        return True

    # 'și' între două referințe de stradă
    if _MULTI_ADDR_SI_PATTERN.search(address):
        return True

    return False


# ---------------------------------------------------------------------------
# 4. GOOGLE GEOCODER
# ---------------------------------------------------------------------------

class GoogleGeocoder:
    """Wrapper peste Google Geocoding API cu retry, backoff, rate limiting."""

    def __init__(self, api_key: str, logger: logging.Logger):
        self.api_key = api_key
        self.logger = logger
        self.session = requests.Session()
        self._last_request_time = 0.0

    def _rate_limit(self):
        """Respectă pauza minimă între cereri."""
        elapsed = time.time() - self._last_request_time
        if elapsed < RATE_LIMIT_PAUSE:
            time.sleep(RATE_LIMIT_PAUSE - elapsed)

    def _safe_url(self, address: str) -> str:
        """Returnează URL-ul fără API key (pentru log)."""
        return f"{GEOCODE_URL}?address={requests.utils.quote(address)}"

    def geocode(self, address: str) -> dict:
        """
        Geocodează o adresă.
        Returnează dict cu: formatted, lat, lon, place_id, error
        """
        result = {
            "formatted": None,
            "lat": None,
            "lon": None,
            "place_id": None,
            "error": None,
        }

        if not address or not address.strip():
            result["error"] = "ERROR: EmptyAddress"
            return result

        params = {
            "address": address,
            "key": self.api_key,
        }

        self.logger.debug("Geocoding: %s", address)
        self.logger.debug("URL (fără cheie): %s", self._safe_url(address))

        last_exception = None
        for attempt in range(RETRY_COUNT):
            try:
                self._rate_limit()
                self._last_request_time = time.time()

                resp = self.session.get(
                    GEOCODE_URL,
                    params=params,
                    timeout=TIMEOUT_SECONDS,
                )
                data = resp.json()

                self.logger.debug("Răspuns Google (attempt %d): %s",
                                  attempt + 1,
                                  json.dumps(data, ensure_ascii=False, indent=2))

                status = data.get("status", "UNKNOWN")

                if status == "OK" and data.get("results"):
                    top = data["results"][0]
                    result["formatted"] = top.get("formatted_address")
                    loc = top.get("geometry", {}).get("location", {})
                    result["lat"] = loc.get("lat")
                    result["lon"] = loc.get("lng")
                    result["place_id"] = top.get("place_id")

                    # Verifică partial_match
                    if top.get("partial_match"):
                        result["error"] = "ERROR: PartialMatch"
                        self.logger.warning("PartialMatch pentru: %s", address)

                    return result

                elif status == "ZERO_RESULTS":
                    result["error"] = "ERROR: NotFound"
                    self.logger.warning("ZERO_RESULTS pentru: %s", address)
                    return result

                elif status == "OVER_QUERY_LIMIT":
                    result["error"] = "OVER_QUERY_LIMIT"
                    self.logger.error("OVER_QUERY_LIMIT! Oprire necesară.")
                    return result

                elif status == "REQUEST_DENIED":
                    result["error"] = "ERROR: REQUEST_DENIED"
                    err_msg = data.get("error_message", "")
                    self.logger.error("REQUEST_DENIED: %s", err_msg)
                    return result

                elif status == "INVALID_REQUEST":
                    result["error"] = "ERROR: INVALID_REQUEST"
                    self.logger.error("INVALID_REQUEST pentru: %s", address)
                    return result

                else:
                    result["error"] = f"ERROR: {status}"
                    self.logger.error("Status necunoscut: %s", status)
                    return result

            except requests.exceptions.Timeout:
                last_exception = "Timeout"
                self.logger.warning("Timeout (attempt %d/%d) pentru: %s",
                                    attempt + 1, RETRY_COUNT, address)
            except requests.exceptions.ConnectionError as e:
                last_exception = str(e)
                self.logger.warning("ConnectionError (attempt %d/%d): %s",
                                    attempt + 1, RETRY_COUNT, e)
            except Exception as e:
                last_exception = str(e)
                self.logger.error("Eroare neașteptată (attempt %d/%d): %s",
                                  attempt + 1, RETRY_COUNT, e)

            # Backoff exponențial
            if attempt < RETRY_COUNT - 1:
                wait = RETRY_BACKOFF_BASE * (2 ** attempt)
                self.logger.debug("Retry în %ds...", wait)
                time.sleep(wait)

        result["error"] = f"ERROR: {last_exception} (după {RETRY_COUNT} încercări)"
        return result


# ---------------------------------------------------------------------------
# 5. CACHE
# ---------------------------------------------------------------------------

class GeoCache:
    """Cache dublu: memorie + persistent pe disc (JSON)."""

    def __init__(self, cache_path: str, logger: logging.Logger):
        self.cache_path = cache_path
        self.logger = logger
        self._cache: dict[str, dict] = {}
        self._load()

    def _load(self):
        """Încarcă cache-ul de pe disc."""
        if os.path.exists(self.cache_path):
            try:
                with open(self.cache_path, 'r', encoding='utf-8') as f:
                    self._cache = json.load(f)
                self.logger.info("Cache încărcat: %d intrări din %s",
                                 len(self._cache), self.cache_path)
            except (json.JSONDecodeError, IOError) as e:
                self.logger.warning("Nu s-a putut încărca cache-ul: %s", e)
                self._cache = {}
        else:
            self.logger.info("Cache nou – nu există fișier: %s", self.cache_path)

    def save(self):
        """Salvează cache-ul pe disc."""
        try:
            with open(self.cache_path, 'w', encoding='utf-8') as f:
                json.dump(self._cache, f, ensure_ascii=False, indent=2)
            self.logger.debug("Cache salvat: %d intrări", len(self._cache))
        except IOError as e:
            self.logger.error("Eroare la salvarea cache-ului: %s", e)

    def get(self, normalized_address: str) -> dict | None:
        """Caută în cache (cheie = adresă normalizată lowercase)."""
        key = normalized_address.strip().lower()
        return self._cache.get(key)

    def put(self, normalized_address: str, result: dict):
        """Adaugă în cache."""
        key = normalized_address.strip().lower()
        self._cache[key] = result

    def __len__(self):
        return len(self._cache)


# ---------------------------------------------------------------------------
# 6. UTILITĂȚI EXCEL (openpyxl)
# ---------------------------------------------------------------------------

def find_or_create_column(ws, header_row: int, col_name: str) -> int:
    """
    Găsește coloana cu numele dat sau o creează la sfârșit.
    Returnează indexul coloanei (1-based).
    """
    for col_idx in range(1, ws.max_column + 1):
        cell_val = ws.cell(row=header_row, column=col_idx).value
        if cell_val and str(cell_val).strip() == col_name:
            return col_idx

    # Creează coloana la sfârșit
    new_col = ws.max_column + 1
    ws.cell(row=header_row, column=new_col, value=col_name)
    return new_col


def ensure_output_columns(ws, header_row: int) -> dict[str, int]:
    """
    Creează sau identifică toate cele 10 coloane de output.
    Returnează dict {nume_coloana: index}.
    """
    col_map = {}
    all_cols = COLS_ORIGINE + COLS_DESTINATIE
    for name in all_cols:
        col_map[name] = find_or_create_column(ws, header_row, name)
    return col_map


# ---------------------------------------------------------------------------
# 7. LOGICA PRINCIPALĂ
# ---------------------------------------------------------------------------

def handle_address(
    address_raw: str,
    geocoder: GoogleGeocoder,
    cache: GeoCache,
    logger: logging.Logger,
) -> dict:
    """
    Procesează o singură adresă: normalizare → cache → geocodare.
    Returnează dict cu: formatted, lat, lon, place_id, error
    """
    result = {
        "formatted": None,
        "lat": None,
        "lon": None,
        "place_id": None,
        "error": None,
    }

    if not address_raw or not str(address_raw).strip():
        result["error"] = "ERROR: EmptyAddress"
        return result

    raw = str(address_raw).strip()
    normalized = normalize_address(raw)
    logger.debug("Normalizat: '%s' → '%s'", raw, normalized)

    if not normalized:
        result["error"] = "ERROR: EmptyAfterNormalization"
        return result

    # Verifică ambiguitate
    if is_ambiguous(normalized):
        result["error"] = "ERROR: Ambiguous"
        logger.warning("Adresă ambiguă: '%s'", normalized)
        return result

    # Verifică cache
    cached = cache.get(normalized)
    if cached is not None:
        logger.debug("Cache hit: '%s'", normalized)
        return cached.copy()

    # Geocodare efectivă
    geo_result = geocoder.geocode(normalized)

    # Salvează în cache (inclusiv erorile, pentru a nu repeta)
    if geo_result.get("error") != "OVER_QUERY_LIMIT":
        cache.put(normalized, geo_result)

    return geo_result


def process_file(
    input_path: str,
    sheet_name: str,
    header_row: int,
    col_origine_name: str,
    col_destinatie_name: str,
    api_key: str,
    dry_run: int,
    start_row: int | None,
    end_row: int | None,
):
    """Logica principală de procesare a fișierului Excel."""

    # --- Derivăm numele fișierelor output ---
    input_p = Path(input_path)
    stem = input_p.stem
    output_path = input_p.parent / f"{stem}Codat.xlsx"
    cache_path = input_p.parent / f"{stem}_cache.json"
    log_path = input_p.parent / f"{stem}Codat.log"

    # --- Configurare logging ---
    logger = logging.getLogger("geocodare_od")
    logger.setLevel(logging.DEBUG)
    # Elimină handler-e vechi
    logger.handlers.clear()

    # Handler fișier – DEBUG
    fh = logging.FileHandler(str(log_path), encoding='utf-8', mode='a')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        '%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%dT%H:%M:%S%z'
    ))
    logger.addHandler(fh)

    # Handler consolă – INFO
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
    logger.addHandler(ch)

    logger.info("=" * 60)
    logger.info("Început procesare: %s", input_path)
    logger.info("Sheet: %s | Antet: rând %d", sheet_name, header_row)
    logger.info("Origine: '%s' | Destinație: '%s'", col_origine_name, col_destinatie_name)
    logger.info("Output: %s", output_path)
    logger.info("Cache: %s", cache_path)
    logger.info("Log: %s", log_path)

    # --- Încărcăm workbook-ul ---
    logger.info("Se încarcă workbook-ul...")
    wb = openpyxl.load_workbook(str(input_path))
    ws = wb[sheet_name]

    # --- Găsim coloanele Origine/Destinație ---
    col_origine_idx = None
    col_destinatie_idx = None
    for col_idx in range(1, ws.max_column + 1):
        cell_val = ws.cell(row=header_row, column=col_idx).value
        if cell_val is not None:
            name = str(cell_val).strip()
            if name == col_origine_name:
                col_origine_idx = col_idx
            elif name == col_destinatie_name:
                col_destinatie_idx = col_idx

    if col_origine_idx is None:
        logger.error("Coloana origine '%s' nu a fost găsită!", col_origine_name)
        print(f"\n❌ EROARE: Coloana '{col_origine_name}' nu există în sheet-ul '{sheet_name}'!")
        return
    if col_destinatie_idx is None:
        logger.error("Coloana destinație '%s' nu a fost găsită!", col_destinatie_name)
        print(f"\n❌ EROARE: Coloana '{col_destinatie_name}' nu există în sheet-ul '{sheet_name}'!")
        return

    logger.info("Coloana origine: index %d | Coloana destinație: index %d",
                col_origine_idx, col_destinatie_idx)

    # --- Creăm coloanele de output ---
    col_map = ensure_output_columns(ws, header_row)
    logger.info("Coloane output create/identificate: %s",
                {k: v for k, v in col_map.items()})

    # --- Inițializăm cache și geocoder ---
    cache = GeoCache(str(cache_path), logger)
    geocoder = GoogleGeocoder(api_key, logger)

    # --- Determinăm rândurile de procesat ---
    total_rows = ws.max_row
    first_data_row = header_row + 1
    last_data_row = total_rows

    if start_row is not None:
        first_data_row = max(first_data_row, start_row)
    if end_row is not None:
        last_data_row = min(last_data_row, end_row)

    # Dry run
    if dry_run > 0:
        last_data_row = min(last_data_row, first_data_row + dry_run - 1)
        logger.info("DRY RUN: se procesează doar %d rânduri (de la %d la %d)",
                     dry_run, first_data_row, last_data_row)

    rows_to_process = list(range(first_data_row, last_data_row + 1))
    num_rows = len(rows_to_process)
    logger.info("Total rânduri de procesat: %d (rândurile %d–%d)",
                num_rows, first_data_row, last_data_row)

    # --- Procesare ---
    processed = 0
    api_calls = 0
    cache_hits = 0
    errors_count = 0
    over_query_limit = False

    # Progress bar
    if tqdm is not None:
        pbar = tqdm(total=num_rows, desc="Geocodare", unit="rând")
    else:
        pbar = None

    for row_idx in rows_to_process:
        # --- Verificăm dacă rândul este deja procesat (reluare) ---
        # Un rând e considerat procesat dacă cel puțin una din coloanele
        # DeCodatOriginea / ErrorOriginea are valoare
        existing_decoded = ws.cell(row=row_idx, column=col_map["DeCodatOriginea"]).value
        existing_error = ws.cell(row=row_idx, column=col_map["ErrorOriginea"]).value
        already_done = existing_decoded is not None or existing_error is not None

        if already_done:
            # Verificăm și destinația
            existing_decoded_d = ws.cell(row=row_idx, column=col_map["DeCodatDestinatia"]).value
            existing_error_d = ws.cell(row=row_idx, column=col_map["ErrorDestinatia"]).value
            if existing_decoded_d is not None or existing_error_d is not None:
                if pbar:
                    pbar.update(1)
                processed += 1
                continue

        # --- ORIGINE ---
        addr_orig = ws.cell(row=row_idx, column=col_origine_idx).value
        result_orig = handle_address(addr_orig, geocoder, cache, logger)

        if result_orig.get("error") == "OVER_QUERY_LIMIT":
            logger.error("OVER_QUERY_LIMIT la rândul %d (Origine). Salvare și oprire.", row_idx)
            over_query_limit = True
            break

        # Scriem rezultatele Origine
        ws.cell(row=row_idx, column=col_map["DeCodatOriginea"]).value = result_orig.get("formatted")
        ws.cell(row=row_idx, column=col_map["LatOriginea"]).value = result_orig.get("lat")
        ws.cell(row=row_idx, column=col_map["LonOriginea"]).value = result_orig.get("lon")
        ws.cell(row=row_idx, column=col_map["PlaceIdOriginea"]).value = result_orig.get("place_id")
        ws.cell(row=row_idx, column=col_map["ErrorOriginea"]).value = result_orig.get("error")

        if result_orig.get("error"):
            errors_count += 1

        # --- DESTINAȚIE ---
        addr_dest = ws.cell(row=row_idx, column=col_destinatie_idx).value
        result_dest = handle_address(addr_dest, geocoder, cache, logger)

        if result_dest.get("error") == "OVER_QUERY_LIMIT":
            logger.error("OVER_QUERY_LIMIT la rândul %d (Destinație). Salvare și oprire.", row_idx)
            over_query_limit = True
            break

        # Scriem rezultatele Destinație
        ws.cell(row=row_idx, column=col_map["DeCodatDestinatia"]).value = result_dest.get("formatted")
        ws.cell(row=row_idx, column=col_map["LatDestinatia"]).value = result_dest.get("lat")
        ws.cell(row=row_idx, column=col_map["LonDestinatia"]).value = result_dest.get("lon")
        ws.cell(row=row_idx, column=col_map["PlaceIdDestinatia"]).value = result_dest.get("place_id")
        ws.cell(row=row_idx, column=col_map["ErrorDestinatia"]).value = result_dest.get("error")

        if result_dest.get("error"):
            errors_count += 1

        processed += 1

        if pbar:
            pbar.update(1)
        else:
            if processed % 50 == 0 or processed == num_rows:
                print(f"  Progres: {processed}/{num_rows} rânduri procesate")

        # --- CHECKPOINT ---
        if processed % CHECKPOINT_EVERY == 0:
            logger.info("CHECKPOINT la rândul %d (%d rânduri procesate)", row_idx, processed)
            wb.save(str(output_path))
            cache.save()
            logger.info("Checkpoint salvat: %s", output_path)

    if pbar:
        pbar.close()

    # --- Salvare finală ---
    logger.info("Salvare finală...")
    wb.save(str(output_path))
    cache.save()

    # --- Raport ---
    print()
    print("=" * 60)
    print("RAPORT FINAL")
    print("=" * 60)
    print(f"  Fișier input:     {input_path}")
    print(f"  Fișier output:    {output_path}")
    print(f"  Fișier cache:     {cache_path}")
    print(f"  Fișier log:       {log_path}")
    print(f"  Rânduri procesate: {processed}/{num_rows}")
    print(f"  Intrări cache:    {len(cache)}")
    print(f"  Erori:            {errors_count}")
    if over_query_limit:
        print()
        print("  ⚠️  OVER_QUERY_LIMIT – scriptul s-a oprit.")
        print("  Relansați scriptul pentru a continua de unde a rămas.")
    print("=" * 60)

    logger.info("Procesare terminată. Rânduri: %d/%d | Erori: %d | Cache: %d",
                processed, num_rows, errors_count, len(cache))
    if over_query_limit:
        logger.warning("Oprire din cauza OVER_QUERY_LIMIT.")


# ---------------------------------------------------------------------------
# 8. FILE PICKER + SELECTARE SHEET + COLOANE
# ---------------------------------------------------------------------------

def pick_file() -> str | None:
    """Deschide un dialog grafic pentru selectarea fișierului .xlsx."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)

        file_path = filedialog.askopenfilename(
            title="Selectează fișierul Excel (.xlsx)",
            filetypes=[("Fișiere Excel", "*.xlsx"), ("Toate fișierele", "*.*")],
        )

        root.destroy()

        if file_path:
            return file_path
        return None
    except Exception as e:
        print(f"Eroare la deschiderea dialogului: {e}")
        print("Introduceți calea fișierului manual:")
        return input("> ").strip().strip('"').strip("'")


def select_sheet(wb) -> str:
    """Selectează sheet-ul din workbook."""
    sheets = wb.sheetnames
    if len(sheets) == 1:
        print(f"  Sheet unic detectat: '{sheets[0]}'")
        return sheets[0]

    print("\nSheet-uri disponibile:")
    for i, name in enumerate(sheets, 1):
        print(f"  {i}. {name}")

    while True:
        try:
            choice = input("\nSelectați numărul sheet-ului: ").strip()
            idx = int(choice) - 1
            if 0 <= idx < len(sheets):
                return sheets[idx]
            print(f"  Număr invalid. Alegeți între 1 și {len(sheets)}.")
        except ValueError:
            print("  Introduceți un număr valid.")


def identify_header_row(ws) -> int:
    """Identifică rândul cu antetul."""
    # Verifică dacă rândul 1 are date
    row1_values = [ws.cell(row=1, column=c).value for c in range(1, min(ws.max_column + 1, 30))]
    non_empty = [v for v in row1_values if v is not None]

    if non_empty:
        print(f"\n  Antet detectat pe rândul 1 ({len(non_empty)} coloane cu date)")
        return 1
    else:
        print("\n  Rândul 1 pare gol.")
        while True:
            try:
                row_num = int(input("  Introduceți numărul rândului cu antetul: ").strip())
                if row_num >= 1:
                    return row_num
                print("  Numărul trebuie să fie >= 1.")
            except ValueError:
                print("  Introduceți un număr valid.")


def select_column(ws, header_row: int, prompt: str) -> str:
    """
    Afișează coloanele și lasă utilizatorul să introducă numele exact.
    Returnează numele coloanei.
    """
    print(f"\nColoanele din sheet (rândul {header_row}):")
    col_names = []
    for col_idx in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col_idx).value
        if val is not None:
            name = str(val).strip()
            col_names.append(name)
            # Afișăm primele 80 caractere din nume
            display = name if len(name) <= 80 else name[:77] + "..."
            print(f"  {col_idx:3d}. {display}")

    print()
    while True:
        user_input = input(f"{prompt}\n  (introduceți numele exact sau numărul coloanei): ").strip()

        # Permite și introducerea numărului coloanei
        try:
            col_num = int(user_input)
            if 1 <= col_num <= ws.max_column:
                cell_val = ws.cell(row=header_row, column=col_num).value
                if cell_val:
                    name = str(cell_val).strip()
                    print(f"  ✓ Selectat: '{name}'")
                    return name
        except ValueError:
            pass

        # Caută după nume
        if user_input in col_names:
            print(f"  ✓ Selectat: '{user_input}'")
            return user_input

        # Caută parțial (conține)
        matches = [n for n in col_names if user_input.lower() in n.lower()]
        if len(matches) == 1:
            print(f"  ✓ Găsit: '{matches[0]}'")
            return matches[0]
        elif len(matches) > 1:
            print(f"  Mai multe potriviri: {matches}")
            print("  Fiți mai specific.")
        else:
            print(f"  Nu s-a găsit coloana '{user_input}'. Încercați din nou.")


# ---------------------------------------------------------------------------
# 9. MAIN + CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Geocodare Origine-Destinație pentru fișiere Excel (.xlsx)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemple:
  python geocodare_od.py                        # file picker + dry run 100
  python geocodare_od.py --dry-run 0            # procesare completă
  python geocodare_od.py --start 100 --end 500  # procesare parțială
  python geocodare_od.py --file date.xlsx       # fără file picker
        """
    )
    parser.add_argument("--file", "-f", type=str, default=None,
                        help="Calea către fișierul Excel (dacă nu se specifică, se deschide file picker)")
    parser.add_argument("--dry-run", type=int, default=DEFAULT_DRY_RUN,
                        help=f"Procesează doar primele N rânduri (implicit {DEFAULT_DRY_RUN}, 0=dezactivat)")
    parser.add_argument("--start", type=int, default=None,
                        help="Rândul de început (1-based, inclusiv)")
    parser.add_argument("--end", type=int, default=None,
                        help="Rândul de sfârșit (1-based, inclusiv)")
    parser.add_argument("--api-key", type=str, default=None,
                        help="Google API Key (implicit: env GOOGLE_API_KEY)")

    args = parser.parse_args()

    print()
    print("=" * 60)
    print("  GEOCODARE ORIGINE–DESTINAȚIE  v1.0")
    print("=" * 60)
    print()

    # --- API Key ---
    api_key = args.api_key or os.environ.get("GOOGLE_API_KEY") or os.environ.get("GOOGLE_MAPS_API_KEY")
    if not api_key:
        print("❌ EROARE: Cheia Google API lipsește!")
        print()
        print("Setați variabila de mediu GOOGLE_API_KEY:")
        print('  Windows:  set GOOGLE_API_KEY=cheia_ta')
        print('  Linux:    export GOOGLE_API_KEY=cheia_ta')
        print()
        print("Sau folosiți:  python geocodare_od.py --api-key CHEIA_TA")
        sys.exit(1)

    print(f"  ✓ API Key: {'*' * 8}...{api_key[-4:]}")

    # --- Selectare fișier ---
    file_path = args.file
    if not file_path:
        print("\nDeschid dialogul de selectare a fișierului...")
        file_path = pick_file()

    if not file_path:
        print("❌ Nu s-a selectat niciun fișier.")
        sys.exit(1)

    if not os.path.exists(file_path):
        print(f"❌ Fișierul nu există: {file_path}")
        sys.exit(1)

    if not file_path.lower().endswith('.xlsx'):
        print(f"❌ Fișierul trebuie să fie .xlsx: {file_path}")
        sys.exit(1)

    print(f"  ✓ Fișier: {file_path}")

    # --- Încărcăm workbook-ul (temporar, pentru selecție) ---
    print("\nSe încarcă fișierul...")
    wb_temp = openpyxl.load_workbook(file_path, read_only=True, data_only=True)

    # --- Selectare sheet ---
    sheet_name = select_sheet(wb_temp)
    ws_temp = wb_temp[sheet_name]

    # --- Identificare antet ---
    header_row = identify_header_row(ws_temp)

    # --- Selectare coloane ---
    col_origine = select_column(ws_temp, header_row,
                                "Selectați coloana ORIGINE (adresa de plecare):")
    col_destinatie = select_column(ws_temp, header_row,
                                   "Selectați coloana DESTINAȚIE (adresa de sosire):")

    # Informații
    total = ws_temp.max_row - header_row
    print(f"\n  Total rânduri de date: {total}")

    wb_temp.close()

    # --- Dry run info ---
    if args.dry_run > 0:
        print(f"\n  ⚡ DRY RUN: se vor procesa doar primele {args.dry_run} rânduri.")
        print(f"     Pentru procesare completă: python geocodare_od.py --dry-run 0")
    else:
        print(f"\n  🔄 Procesare COMPLETĂ: {total} rânduri")

    # --- Confirmare ---
    print()
    confirm = input("Continuăm? (da/nu): ").strip().lower()
    if confirm not in ('da', 'yes', 'y', 'd'):
        print("Anulat de utilizator.")
        sys.exit(0)

    # --- Procesare ---
    process_file(
        input_path=file_path,
        sheet_name=sheet_name,
        header_row=header_row,
        col_origine_name=col_origine,
        col_destinatie_name=col_destinatie,
        api_key=api_key,
        dry_run=args.dry_run,
        start_row=args.start,
        end_row=args.end,
    )


if __name__ == "__main__":
    main()
