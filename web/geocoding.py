"""
Geocodarea unui fișier Excel pe server.

Refolosește normalizarea, detectarea ambiguității și GoogleGeocoder din
geocodare_od.py, ca site-ul și scriptul CLI să dea aceleași rezultate.
Diferențe față de CLI:
  - adresele identice se geocodează o singură dată (deduplicare înainte de cereri)
  - cererile Google rulează în paralel (GOOGLE_WORKERS)
  - cache-ul e comun tuturor fișierelor (SQLite), cu expirare după CACHE_TTL_DAYS
  - REQUEST_DENIED / OVER_QUERY_LIMIT opresc jobul în loc să fie salvate ca rezultat
"""

import json
import logging
import re
import threading
import time
import unicodedata
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from functools import lru_cache

import openpyxl
import requests
from openpyxl.utils import column_index_from_string, get_column_letter

from geocodare_od import (
    GoogleGeocoder,
    RETRY_BACKOFF_BASE,
    RETRY_COUNT,
    TIMEOUT_SECONDS,
    find_or_create_column,
    is_ambiguous,
    normalize_address,
)

from . import config, db

logger = logging.getLogger("geocodare_web")

# Explicații pentru observațiile din coloanele Error…: cod -> (titlu, explicație, gravitate)
ERROR_HELP = {
    "PartialMatch": ("Potrivire aproximativă",
                     "Google a găsit un loc apropiat, dar nu exact adresa scrisă. De obicei e bun, merită verificat.", "warn"),
    "FaraLocalitate": ("Oraș fără nume",
                       "În fișier lipsește numele orașului (cod 9999 sau fără cod). S-au trimis doar strada și județul, "
                       "iar Google a ales orașul. Rezultat probabil, nu sigur.", "warn"),
    "RezultatDoarJudet": ("Doar județul",
                          "Google a găsit doar județul; coordonatele sunt centrul județului. Nu sunt utilizabile.", "bad"),
    "NotFound": ("Negăsită", "Google nu a găsit adresa.", "bad"),
    "FaraLocalitateSiStrada": ("Fără localitate și stradă",
                               "Există doar județul, deci nu s-a trimis la geocodare.", "bad"),
    "Ambiguous": ("Neclară",
                  "Adresa pare să conțină două locuri (ex. „colț cu…”, „;”) și nu s-a trimis.", "bad"),
    "EmptyAddress": ("Goală", "Celula nu conține nicio adresă.", "muted"),
    "EmptyAfterNormalization": ("Goală după curățare", "Celula conținea doar semne sau spații.", "muted"),
    "LowImportance": ("Rezultat slab", "Nominatim a găsit un rezultat puțin relevant. Merită verificat.", "warn"),
    "INVALID_REQUEST": ("Cerere invalidă", "Google nu a putut interpreta adresa.", "bad"),
}


def error_breakdown(summary: str | None) -> list[dict]:
    """„FaraLocalitate: 12, NotFound: 3” → [{code, count, title, text, level}, …]"""
    out = []
    for part in (summary or "").split(", "):
        code, _, count = part.rpartition(": ")
        if not code or not count.isdigit():
            continue
        title, text, level = ERROR_HELP.get(code, (code, "", "muted"))
        out.append({"code": code, "count": int(count), "title": title, "text": text, "level": level})
    return out


PROVIDERS = {
    "google": "Google Geocoding API",
    "nominatim": "OpenStreetMap Nominatim (gratuit, 1 adresă/secundă)",
}

# Erori după care nu are sens să continuăm (cheie invalidă, cotă depășită, blocaj)
FATAL_ERRORS = {
    "OVER_QUERY_LIMIT": "Cota Google a fost depășită (OVER_QUERY_LIMIT).",
    "ERROR: REQUEST_DENIED": "Google a refuzat cererea (REQUEST_DENIED) — verificați cheia API și dacă Geocoding API e activat.",
    "RATE_LIMITED": "Nominatim a limitat cererile (HTTP 429).",
    "ERROR: FORBIDDEN (403)": "Nominatim a blocat cererile (HTTP 403).",
}


class JobCancelled(Exception):
    pass


class JobFailed(Exception):
    pass


# ---------------------------------------------------------------------------
# Nominatim (portat din branch-ul openstreetmap)
# ---------------------------------------------------------------------------

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
NOMINATIM_PAUSE = 1.1


class NominatimGeocoder:
    def __init__(self):
        self.session = requests.Session()
        contact = f" ({config.NOMINATIM_EMAIL})" if config.NOMINATIM_EMAIL else ""
        self.session.headers["User-Agent"] = f"geocodare-od-web/1.0{contact}"
        self._last_request_time = 0.0

    def geocode(self, address: str) -> dict:
        result = {"formatted": None, "lat": None, "lon": None, "place_id": None, "error": None}
        params = {"q": address, "format": "json", "limit": 1, "countrycodes": "ro"}
        last_exception = None
        for attempt in range(RETRY_COUNT):
            try:
                elapsed = time.time() - self._last_request_time
                if elapsed < NOMINATIM_PAUSE:
                    time.sleep(NOMINATIM_PAUSE - elapsed)
                self._last_request_time = time.time()

                resp = self.session.get(NOMINATIM_URL, params=params, timeout=TIMEOUT_SECONDS)
                if resp.status_code == 429:
                    result["error"] = "RATE_LIMITED"
                    return result
                if resp.status_code == 403:
                    result["error"] = "ERROR: FORBIDDEN (403)"
                    return result
                resp.raise_for_status()
                data = resp.json()

                if data:
                    top = data[0]
                    result["formatted"] = top.get("display_name")
                    result["lat"] = float(top["lat"]) if top.get("lat") else None
                    result["lon"] = float(top["lon"]) if top.get("lon") else None
                    result["place_id"] = str(top.get("osm_id", top.get("place_id", "")))
                    if top.get("importance", 0) < 0.01:
                        result["error"] = "WARNING: LowImportance"
                else:
                    result["error"] = "ERROR: NotFound"
                return result
            except requests.exceptions.RequestException as e:
                last_exception = str(e)
                logger.warning("Nominatim (attempt %d/%d): %s", attempt + 1, RETRY_COUNT, e)
            if attempt < RETRY_COUNT - 1:
                time.sleep(RETRY_BACKOFF_BASE * (2 ** attempt))
        result["error"] = f"ERROR: {last_exception} (după {RETRY_COUNT} încercări)"
        return result


def _make_geocoder(provider: str):
    if provider == "google":
        return GoogleGeocoder(config.GOOGLE_API_KEY, logger)
    return NominatimGeocoder()


def available_providers() -> dict:
    return {k: v for k, v in PROVIDERS.items() if k != "google" or config.GOOGLE_API_KEY}


# ---------------------------------------------------------------------------
# Pregătire adrese (aceeași logică ca handle_address din CLI)
# ---------------------------------------------------------------------------

@lru_cache(maxsize=200_000)
def _prep_text(raw: str):
    normalized = normalize_address(raw)
    if not normalized:
        return None, None, "ERROR: EmptyAfterNormalization"
    if is_ambiguous(normalized):
        return None, None, "ERROR: Ambiguous"
    return normalized.strip().lower(), normalized, None


def _blank(value) -> bool:
    return value is None or not str(value).strip()


def prep_address(value):
    """Returnează (cheie cache, adresă normalizată, eroare locală)."""
    if _blank(value):
        return None, None, "ERROR: EmptyAddress"
    return _prep_text(str(value).strip())


def _is_cacheable(result: dict) -> bool:
    err = result.get("error") or ""
    return err not in FATAL_ERRORS and "încercări" not in err


# ---------------------------------------------------------------------------
# Cache (SQLite, comun tuturor joburilor)
# ---------------------------------------------------------------------------

def cache_get_many(provider: str, keys) -> dict:
    keys = list(keys)
    min_created = time.time() - config.CACHE_TTL_DAYS * 86400
    found = {}
    with db.connect() as conn:
        for i in range(0, len(keys), 500):
            chunk = keys[i:i + 500]
            marks = ",".join("?" * len(chunk))
            rows = conn.execute(
                f"SELECT key, result FROM geocache WHERE provider = ? AND created_at >= ? AND key IN ({marks})",
                (provider, min_created, *chunk),
            ).fetchall()
            for row in rows:
                found[row["key"]] = json.loads(row["result"])
    return found


def cache_put_many(provider: str, items: dict):
    now = time.time()
    with db.connect() as conn:
        conn.executemany(
            "INSERT OR REPLACE INTO geocache (provider, key, result, created_at) VALUES (?, ?, ?, ?)",
            [(provider, k, json.dumps(v, ensure_ascii=False), now) for k, v in items.items()],
        )


def purge_expired_cache():
    db.execute("DELETE FROM geocache WHERE created_at < ?",
               (time.time() - config.CACHE_TTL_DAYS * 86400,))


# ---------------------------------------------------------------------------
# Citire Excel
# ---------------------------------------------------------------------------

# Grup de coloane: județ, (ignorat), cod localitate, stradă / punct de reper
GROUP_SIZE = 4

# Răspunsuri care nu sunt o stradă: „Nu știu…”, „Da”, un număr singur etc.
_NOT_A_STREET = re.compile(r"^(nu\b.*|da|nimic|-+|\.+|\d{1,3}|x+)$", re.IGNORECASE)
_BUCHAREST = {"bucuresti", "bucurești", "municipiul bucuresti", "municipiul bucurești"}


def _label(value, idx: int) -> str:
    text = " ".join(str(value).split()) if value is not None else ""
    return text or f"(coloana {get_column_letter(idx)} fără nume)"


def detect_header_row(ws) -> int:
    """Primul rând (din primele 20) cu cel puțin 2 celule completate."""
    for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), start=1):
        if sum(v is not None and str(v).strip() != "" for v in row) >= 2:
            return idx
    return 1


_COUNTIES = {
    "alba", "arad", "arges", "bacau", "bihor", "bistrita nasaud", "botosani", "braila", "brasov",
    "bucuresti", "municipiul bucuresti", "buzau", "calarasi", "caras severin", "cluj", "constanta",
    "covasna", "dambovita", "dolj", "galati", "giurgiu", "gorj", "harghita", "hunedoara", "ialomita",
    "iasi", "ilfov", "maramures", "mehedinti", "mures", "neamt", "olt", "prahova", "salaj", "satu mare",
    "sibiu", "suceava", "teleorman", "timis", "tulcea", "valcea", "vaslui", "vrancea",
}


def _plain(text) -> str:
    """„Bistrița-Năsăud” → „bistrita nasaud”."""
    text = unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode()
    return " ".join(re.sub(r"[^a-z ]", " ", text.lower()).split())


def _looks_like_counties(values) -> bool:
    values = [v for v in values if not _blank(v)]
    return bool(values) and sum(_plain(v) in _COUNTIES for v in values) >= 0.6 * len(values)


def _fill_color(cell):
    fill = getattr(cell, "fill", None)
    if fill is None or not fill.fill_type:
        return None
    color = fill.fgColor
    if color.type == "rgb":
        return color.rgb
    if color.type == "theme":
        return f"theme{color.theme}/{color.tint}"
    return f"indexed{color.indexed}"


def _detect_groups(fills: list) -> list[int]:
    """
    Grupuri = coloane vecine din antet colorate la fel. O porțiune colorată mai lungă de 4
    se împarte de la dreapta (ex. „Q161_6, Q1611_6, Q162_6 … Q164_6” → ultimele 4;
    „M3JudetO_5 … M3StradaD_5” → 2 grupuri). Returnează indexul primei coloane din fiecare grup.
    """
    groups, i = [], 0
    while i < len(fills):
        j = i
        while fills[i] is not None and j + 1 < len(fills) and fills[j + 1] == fills[i]:
            j += 1
        if fills[i] is not None:
            run = list(range(i + 1, j + 2))          # indexuri 1-based
            while len(run) >= GROUP_SIZE:
                groups.append(run[-GROUP_SIZE])
                run = run[:-GROUP_SIZE]
        i = j + 1
    return sorted(groups)


def inspect_sheet(path, sheet: str | None, header_row: int | None) -> dict:
    """Foi, antet, coloane (cu exemple de valori) pentru pagina de configurare."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sheets = wb.sheetnames
        if sheet not in sheets:
            sheet = sheets[0]
        ws = wb[sheet]
        if not header_row:
            header_row = detect_header_row(ws)

        header = next(ws.iter_rows(min_row=header_row, max_row=header_row), ())
        columns, fills = [], []
        for idx, cell in enumerate(header, start=1):
            columns.append({"index": idx, "letter": get_column_letter(idx),
                            "label": _label(cell.value, idx), "samples": []})
            fills.append(_fill_color(cell))

        # Grupuri candidate după culoare; se păstrează doar cele cu județe în prima coloană
        candidates = {first: [] for first in _detect_groups(fills)}
        rows = 0
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            if not any(v is not None and str(v).strip() != "" for v in row):
                continue
            rows += 1
            for first, seen in candidates.items():
                v = row[first - 1] if first - 1 < len(row) else None
                if len(seen) < 200 and not _blank(v):
                    seen.append(v)
            if rows <= 3:
                for col in columns:
                    v = row[col["index"] - 1] if col["index"] - 1 < len(row) else None
                    if v is not None and str(v).strip():
                        col["samples"].append(str(v).strip()[:80])
    finally:
        wb.close()

    return {
        "sheets": sheets,
        "sheet": sheet,
        "header_row": header_row,
        "columns": columns,
        "rows": rows,
        "guess_origin": _guess(columns, "origin"),
        "guess_dest": _guess(columns, "destina"),
        "guess_groups": [first for first, seen in candidates.items() if _looks_like_counties(seen)],
        "guess_lookup": next((s for s in sheets if s != sheet and "coresp" in s.lower()), None),
    }


def _guess(columns, needle: str):
    for col in columns:
        if needle in col["label"].lower():
            return col["index"]
    return None


def parse_groups(text: str) -> list[int]:
    """„GY, HO, ID” → [207, 223, 238]. ValueError dacă o literă nu e validă."""
    out = []
    for part in re.split(r"[\s,;]+", text or ""):
        if part:
            out.append(column_index_from_string(part.upper()))
    return sorted(set(out))


def format_groups(groups) -> str:
    return ", ".join(get_column_letter(g) for g in groups or [])


def _code(value):
    """Codul localității ca text: 2184, 2184.0 și „2184” → „2184”."""
    if _blank(value):
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def load_lookup(path, sheet: str | None) -> dict:
    """Foaia de corespondență: prima coloană = cod localitate, a doua = nume."""
    if not sheet:
        return {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        lookup = {}
        for row in wb[sheet].iter_rows(max_col=2, values_only=True):
            if len(row) == 2 and _code(row[0]) and not _blank(row[1]):
                lookup[_code(row[0])] = str(row[1]).strip()
        return lookup
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Ținte: ce se geocodează dintr-un rând
# ---------------------------------------------------------------------------

@dataclass
class Target:
    suffix: str             # „Originea” sau „_Q162_1” → coloanele DeCodat<sufix>, Lat<sufix>…
    cols: tuple             # coloanele citite
    group: bool = False     # True = grup de 4 coloane (adresa se compune)


def build_targets(ws, header_row: int, cfg: dict) -> list[Target]:
    header = next(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True), ())
    targets = []
    if cfg.get("col_origin"):
        targets.append(Target("Originea", (cfg["col_origin"],)))
    if cfg.get("col_dest"):
        targets.append(Target("Destinatia", (cfg["col_dest"],)))
    used = set()
    for first in cfg.get("groups") or []:
        name = header[first - 1] if first - 1 < len(header) else None
        name = re.sub(r"\W+", "_", str(name).strip()) if not _blank(name) else get_column_letter(first)
        if name in used:
            name = f"{name}_{get_column_letter(first)}"
        used.add(name)
        targets.append(Target(f"_{name}", tuple(range(first, first + GROUP_SIZE)), group=True))
    return targets


def compose(target: Target, values: list, lookup: dict):
    """
    Returnează (adresă, eroare locală, avertisment) sau None dacă nu e nimic de geocodat.
    Ambiguitatea („colț cu…”, „și”) se verifică doar pe stradă, nu pe numele localității
    (ex. comuna „COCORASTII COLT”).
    Pentru grupuri: „stradă, localitate, județul X, România”; când codul lipsește sau e
    9999 (URBAN) se trimite doar strada + județul și rezultatul primește „FaraLocalitate”.
    """
    if not target.group:
        v = values[0]
        return None if _blank(v) else (str(v).strip(), None, None, True)

    county, _, code, street = values
    if _blank(county) and _blank(code) and _blank(street):
        return None
    county = None if _blank(county) else " ".join(str(county).split())
    street = None if _blank(street) else " ".join(str(street).split())
    if street and _NOT_A_STREET.match(street):
        street = None
    if street and is_ambiguous(normalize_address(street)):
        return None, "ERROR: Ambiguous", None, False
    locality = lookup.get(_code(code))
    if locality and locality.upper() == "URBAN":
        locality = None

    parts = [p for p in (street, locality) if p]
    if county and county.lower() in _BUCHAREST:
        parts.append("București")
        locality = locality or "București"
    elif county:
        parts.append(f"județul {county}")
    if not street and not locality:
        return None, "ERROR: FaraLocalitateSiStrada", None, False
    parts.append("România")
    return ", ".join(parts), None, (None if locality else "WARNING: FaraLocalitate"), False


def _header_map(ws, header_row: int) -> dict:
    header = next(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True), ())
    return {str(v).strip(): idx for idx, v in enumerate(header, start=1) if v is not None}


def _read_addresses(ws, header_row: int, targets, lookup, max_rows: int | None, skip_done: bool):
    """
    Returnează ([(rând, {sufix: (adresă, eroare locală, avertisment)})], câte adrese aveau deja rezultat).

    Cu skip_done, o adresă e sărită dacă rândul are deja DeCodat<sufix> sau Error<sufix>
    completat (fișiere geocodate parțial). Rândurile fără nicio adresă de procesat sunt sărite.
    max_rows numără doar rândurile cu ceva de procesat, ca un test să nu cadă pe rânduri deja gata.
    """
    done_cols = {}
    if skip_done:
        headers = _header_map(ws, header_row)
        for t in targets:
            cols = [headers.get(f"DeCodat{t.suffix}"), headers.get(f"Error{t.suffix}")]
            if any(cols):
                done_cols[t.suffix] = [c for c in cols if c]

    needed = [c for t in targets for c in t.cols] + [c for cols in done_cols.values() for c in cols]
    lo, hi = min(needed), max(needed)
    first = header_row + 1
    out, already = [], 0
    for r_idx, row in enumerate(
        ws.iter_rows(min_row=first, max_row=ws.max_row, min_col=lo, max_col=hi, values_only=True),
        start=first,
    ):
        if max_rows and len(out) >= max_rows:
            break
        items = {}
        for t in targets:
            composed = compose(t, [_at(row, c - lo) for c in t.cols], lookup)
            if composed is None:
                continue
            if any(not _blank(_at(row, c - lo)) for c in done_cols.get(t.suffix, ())):
                already += 1
                continue
            items[t.suffix] = composed
        if items:
            out.append((r_idx, items))
    return out, already


def _at(row, i):
    return row[i] if i < len(row) else None


def _prep_composed(composed):
    address, local_err, _, check_ambiguous = composed
    if local_err:
        return None, None, local_err
    if check_ambiguous:
        return prep_address(address)
    normalized = normalize_address(address)
    return normalized.strip().lower(), normalized, None


def job_config(job) -> dict:
    return {
        "col_origin": job["col_origin"],
        "col_dest": job["col_dest"],
        "groups": json.loads(job["groups"]) if job["groups"] else [],
        "lookup_sheet": job["lookup_sheet"],
    }


def estimate(path, sheet, header_row, cfg, provider, max_rows, skip_done) -> dict:
    """Câte adrese unice trebuie trimise efectiv la geocodare."""
    lookup = load_lookup(path, cfg.get("lookup_sheet"))
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet]
        targets = build_targets(ws, header_row, cfg)
        rows, already = _read_addresses(ws, header_row, targets, lookup, max_rows, skip_done)
    finally:
        wb.close()
    keys, local_errors, addresses, no_locality = set(), 0, 0, 0
    for _, items in rows:
        for composed in items.values():
            addresses += 1
            no_locality += composed[2] is not None
            key, _, err = _prep_composed(composed)
            if key:
                keys.add(key)
            else:
                local_errors += 1
    cached = len(cache_get_many(provider, keys))
    return {
        "rows": len(rows),
        "addresses": addresses,
        "already": already,
        "unique": len(keys),
        "cached": cached,
        "todo": len(keys) - cached,
        "skipped": local_errors,
        "no_locality": no_locality,
    }


# ---------------------------------------------------------------------------
# Rulare job
# ---------------------------------------------------------------------------

# Google a găsit doar județul (coordonatele sunt centrul județului), ex. „Ialomița County, Romania”
_COUNTY_ONLY = re.compile(r"^(Județul [^,]+|[^,]+ County), Rom[aâ]nia$", re.IGNORECASE)

OUTPUT_FIELDS = [
    ("DeCodat{}", "formatted"),
    ("Lat{}", "lat"),
    ("Lon{}", "lon"),
    ("PlaceId{}", "place_id"),
    ("Error{}", "error"),
]


def run_job(job, input_path, output_path, should_cancel):
    job_id = job["id"]
    provider = job["provider"]
    cfg = job_config(job)
    lookup = load_lookup(input_path, cfg["lookup_sheet"])

    db.update_job(job_id, phase="Citire fișier")
    wb = openpyxl.load_workbook(input_path)
    ws = wb[job["sheet"]]
    targets = build_targets(ws, job["header_row"], cfg)
    rows, already = _read_addresses(ws, job["header_row"], targets, lookup, job["max_rows"],
                                    job["skip_done"] != 0)

    prepared = []           # (rând, {sufix: (compus, (cheie, normalizat, eroare locală))})
    to_resolve = {}         # cheie -> adresă normalizată
    for r_idx, items in rows:
        prepped = {suffix: (composed, _prep_composed(composed)) for suffix, composed in items.items()}
        prepared.append((r_idx, prepped))
        for _, (key, normalized, _) in prepped.values():
            if key:
                to_resolve.setdefault(key, normalized)

    results = cache_get_many(provider, to_resolve)
    todo = {k: v for k, v in to_resolve.items() if k not in results}
    db.update_job(job_id, rows_count=len(rows), already_done=already, unique_count=len(to_resolve),
                  cache_hits=len(results), total=len(todo), done=0, phase="Geocodare")

    if todo:
        _geocode_all(job_id, provider, todo, results, should_cancel)

    db.update_job(job_id, phase="Scriere rezultate")
    # Coloanele noi se adaugă la finalul foii, ca formulele existente să nu se strice
    col_map = {}
    for t in targets:
        names = ([f"Adresa{t.suffix}"] if t.group else []) + [tpl.format(t.suffix) for tpl, _ in OUTPUT_FIELDS]
        for name in names:
            col_map[name] = find_or_create_column(ws, job["header_row"], name)

    error_types = Counter()
    addresses = found = 0
    for r_idx, prepped in prepared:
        for suffix, ((address, _, warning, _), (key, _, local_err)) in prepped.items():
            res = dict(results[key]) if key else {"error": local_err}
            county_only = res.get("formatted") and _COUNTY_ONLY.match(res["formatted"])
            res["error"] = "; ".join(e for e in (warning, county_only and "WARNING: RezultatDoarJudet",
                                                  res.get("error")) if e) or None
            if f"Adresa{suffix}" in col_map:
                ws.cell(row=r_idx, column=col_map[f"Adresa{suffix}"]).value = address
            for col_tpl, field in OUTPUT_FIELDS:
                ws.cell(row=r_idx, column=col_map[col_tpl.format(suffix)]).value = res.get(field)
            for err in (res["error"] or "").split("; "):
                if err:
                    error_types[err.replace("ERROR: ", "").replace("WARNING: ", "")] += 1
            addresses += 1
            found += res.get("lat") is not None

    db.update_job(job_id, phase="Salvare fișier", addresses=addresses, found=found)
    wb.save(output_path)
    summary = ", ".join(f"{name}: {n}" for name, n in error_types.most_common())
    return sum(error_types.values()), summary


def _geocode_all(job_id, provider, todo, results, should_cancel):
    workers = config.GOOGLE_WORKERS if provider == "google" else 1  # Nominatim: strict 1/s
    local = threading.local()

    def work(address):
        if not hasattr(local, "geocoder"):
            local.geocoder = _make_geocoder(provider)
        return local.geocoder.geocode(address)

    pending_cache = {}
    done = 0
    last_update = 0.0
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(work, addr): key for key, addr in todo.items()}
        try:
            for fut in as_completed(futures):
                key = futures[fut]
                res = fut.result()
                err = res.get("error")
                if err in FATAL_ERRORS:
                    raise JobFailed(FATAL_ERRORS[err])
                results[key] = res
                if _is_cacheable(res):
                    pending_cache[key] = res
                done += 1

                now = time.time()
                if now - last_update > 1 or done == len(todo):
                    last_update = now
                    cache_put_many(provider, pending_cache)
                    pending_cache.clear()
                    db.update_job(job_id, done=done)
                    if should_cancel():
                        raise JobCancelled()
        except BaseException:
            pool.shutdown(wait=True, cancel_futures=True)
            raise
        finally:
            # Ce s-a geocodat rămâne în cache, deci reluarea nu mai plătește aceleași cereri
            if pending_cache:
                cache_put_many(provider, pending_cache)
