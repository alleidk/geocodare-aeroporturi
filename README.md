# 🌍 Geocodare Origine–Destinație

Sistem de geocodare automată pentru fișiere Excel (.xlsx) cu adrese de Origine și Destinație.  
Citește un fișier Excel, normalizează adresele, le geocodează și scrie rezultatele direct în fișierul output — păstrând formatarea originală.

## ✨ Funcționalități

- 📂 **File picker grafic** (Tkinter) pentru selectarea fișierului Excel
- 🧹 **Normalizare adrese românești** — extindere abrevieri (str→Strada, bd→Bulevardul), corectare diacritice (50+ localități)
- ⚠️ **Detectare adrese ambigue** — PO Box, intersecții, adrese multiple
- 🔄 **Retry cu backoff exponențial** — toleranță la erori de rețea
- 💾 **Cache persistent** (JSON) — evită apeluri duplicate, permite reluare
- 📌 **Checkpoint automat** la fiecare 500 rânduri — protecție la crash
- 📊 **Progress bar** (tqdm) în terminal
- 📝 **Logging complet** (DEBUG) cu răspunsuri API
- 🧪 **Dry run** implicit (primele 100 rânduri) + procesare parțială (`--start`/`--end`)

## 🔀 Branch-uri disponibile

| Branch | Backend | API Key | Cost | Rate Limit |
|--------|---------|---------|------|------------|
| **`main`** (implicit) | Google Geocoding API | ✅ Necesară | Plătit | 0.15s între cereri |
| **`openstreetmap`** | Nominatim (OpenStreetMap) | ❌ Nu e necesară | Gratuit | 1.1s între cereri |

### Când să folosești fiecare:
- **Google** (`main`) — acuratețe superioară, găsește hoteluri/puncte de interes, suport excelent pentru adrese internaționale
- **Nominatim** (`openstreetmap`) — gratuit, fără limită de apeluri, potrivit pentru adrese bine structurate

## 📦 Instalare

```bash
git clone https://github.com/alleidk/geocodare-aeroporturi.git
cd geocodare-aeroporturi

# Instalare dependențe
pip install openpyxl requests tqdm
```

### Configurare API Key (doar branch-ul `main`)

```bash
# Copiază fișierul exemplu
cp .env.example .env

# Editează .env și adaugă cheia ta Google
# SAU setează variabila de mediu:
# Windows:
set GOOGLE_API_KEY=cheia_ta
# Linux/Mac:
export GOOGLE_API_KEY=cheia_ta
```

### Pentru branch-ul OpenStreetMap (fără API key)

```bash
git checkout openstreetmap
# Gata — nu necesită configurare
```

## 🚀 Utilizare

```bash
# File picker + dry run (primele 100 rânduri)
python geocodare_od.py

# Procesare completă
python geocodare_od.py --dry-run 0

# Procesare parțială (rândurile 100-500)
python geocodare_od.py --start 100 --end 500

# Fișier specific + procesare completă
python geocodare_od.py --file date.xlsx --dry-run 0
```

### Parametri CLI

| Parametru | Descriere | Implicit |
|-----------|-----------|----------|
| `--file`, `-f` | Calea fișierului Excel | File picker grafic |
| `--dry-run` | Procesează doar N rânduri (0=dezactivat) | 100 |
| `--start` | Rândul de început (1-based) | Primul rând |
| `--end` | Rândul de sfârșit (1-based) | Ultimul rând |
| `--api-key` | Google API Key (doar `main`) | env `GOOGLE_API_KEY` |

## 📄 Coloane de output

Scriptul adaugă 10 coloane noi în fișierul Excel:

**Origine:**
| Coloană | Conținut |
|---------|----------|
| `DeCodatOriginea` | Adresa standardizată (formatted) |
| `LatOriginea` | Latitudine |
| `LonOriginea` | Longitudine |
| `PlaceIdOriginea` | Place ID (Google) / OSM ID (Nominatim) |
| `ErrorOriginea` | Erori (NotFound, Ambiguous, PartialMatch etc.) |

**Destinație:** Aceleași coloane cu sufixul `Destinatia`.

## 📁 Fișiere generate

| Fișier | Descriere |
|--------|-----------|
| `NumeCodat.xlsx` | Excel cu coordonate adăugate |
| `Nume_cache.json` | Cache persistent (permite reluare) |
| `NumeCodat.log` | Log complet DEBUG |

## 🔧 Normalizare adrese

Scriptul normalizează automat adresele românești:

| Input | Output |
|-------|--------|
| `str pacii` | `Strada pacii` |
| `bd unirii nr 10` | `Bulevardul unirii nr. 10` |
| `timisoara` | `Timișoara` |
| `targu mures` | `Târgu Mureș` |
| `cluj napoca` | `Cluj-Napoca` |
| `sos pantelimon` | `Șoseaua pantelimon` |

## 📋 Cerințe

- Python 3.10+
- `openpyxl` — citire/scriere Excel cu păstrare formatare
- `requests` — apeluri HTTP
- `tqdm` — progress bar (opțional)
- `tkinter` — file picker grafic (inclus în Python standard)

## 📜 Licență

MIT
