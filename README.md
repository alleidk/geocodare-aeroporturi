# 🌍 Geocodare adrese

Adaugă coordonate (latitudine, longitudine) pentru adresele dintr-un fișier Excel (.xlsx), folosind Google Geocoding API
(sau, opțional, OpenStreetMap Nominatim). Rezultatul este **același fișier**, cu coloane noi la final — formatarea
și coloanele existente rămân neschimbate.

Proiectul are două variante:

| | Ce este | Pentru cine |
|---|---|---|
| **Site** (`web/`) | Site privat, cu login: încarci Excelul, alegi coloanele, descarci rezultatul. Geocodarea rulează pe server. | utilizare zilnică, de pe orice calculator/telefon |
| **Script** (`geocodare_od.py`) | Program în linia de comandă | rulare locală, fără server |

---

## 🌐 Site-ul

- **Acces doar cu cont** — administratorul creează conturi din pagina *Utilizatori*; fiecare vede doar fișierele lui.
- **Ghicește singur coloanele** cu adrese și arată exemple, ca să le poți verifica.
- **Estimare înainte de pornire** — câte adrese, câte cereri noi, cost și durată aproximative.
- **Mod test** — rulezi întâi pe primele N rânduri, verifici, apoi rulezi tot.
- **Rulează în fundal** — poți închide pagina; dacă serverul repornește, geocodarea continuă de unde a rămas.
- **Memorie (cache)** — o adresă geocodată e ținută minte 30 de zile (limita din termenii Google), deci nu se plătește de două ori.
- **Fișiere geocodate parțial** — adresele care au deja rezultat (`DeCodat…`/`Error…`) sunt sărite.
- **Pagină de ajutor** cu explicația fiecărei observații din coloanele `Error…`.

### Unde pot fi adresele

**A. Într-o singură coloană** (ex. „judetul Dolj, CRAIOVA, Strada Lapus”) — opțional și o coloană de destinație.
Rezultate: `DeCodatOriginea`, `LatOriginea`, `LonOriginea`, `PlaceIdOriginea`, `ErrorOriginea` (și `…Destinatia`).

**B. Împărțite pe grupuri de câte 4 coloane**: județ · (ignorată) · cod localitate · stradă / punct de reper.
Numele localității se ia după cod dintr-o foaie de corespondență (col. 1 = cod, col. 2 = nume, ex. „Coresp Com”).
Grupurile sunt găsite după culoarea antetului (și verificate după numele de județe din prima coloană); lista se poate corecta manual.
La orașe (cod `9999` = URBAN sau fără cod) se trimit strada și județul, iar rezultatul primește `FaraLocalitate`.
Rezultate, pentru fiecare grup: `Adresa_<grup>`, `DeCodat_<grup>`, `Lat_…`, `Lon_…`, `PlaceId_…`, `Error_…`.

### Pornire pe calculatorul propriu

```bash
pip install -r requirements.txt
```

Windows (PowerShell):

```powershell
$env:GOOGLE_API_KEY="cheia_ta"; $env:ADMIN_PASSWORD="o_parola_buna"
python -m web
```

Apoi deschide http://localhost:8000 și intră cu utilizatorul `admin`. Datele site-ului (conturi, fișiere, cache)
se păstrează în folderul `data/`.

### Punere online (server)

Site-ul are nevoie de un server care rulează încontinuu și de un **disc persistent** montat la `/data`
(altfel se pierd conturile, fișierele și cache-ul la fiecare repornire). Merge cu `Dockerfile`-ul din repo pe
orice VPS cu Docker sau pe platforme ca Railway / Fly.io / Render (cu volum persistent).

```bash
docker build -t geocodare-web .
docker run -d --restart unless-stopped -p 8000:8000 -v geocodare-data:/data \
  -e GOOGLE_API_KEY=... -e ADMIN_PASSWORD=... geocodare-web
```

- Pune în fața lui un reverse proxy cu **HTTPS** (Caddy, nginx, Traefik sau HTTPS-ul platformei) și îndreaptă domeniul
  spre el. Imaginea Docker are `COOKIE_SECURE=1`, deci login-ul funcționează doar prin HTTPS.
- Un singur proces (`python -m web`) — geocodarea rulează într-un fir din interiorul lui; nu porni mai multe copii pe același `/data`.
- Pentru fișiere mari (~25 MB, ~1.000 de coloane) serverul are nevoie de ~1,5–2 GB RAM în timpul procesării.
- Cheia Google stă doar pe server; nu ajunge niciodată în browser.

| Variabilă | Descriere | Implicit |
|-----------|-----------|----------|
| `GOOGLE_API_KEY` | Cheia Google Geocoding (fără ea e disponibil doar Nominatim) | — |
| `ADMIN_USERNAME` / `ADMIN_PASSWORD` | Primul cont de administrator (creat doar dacă nu există niciun cont) | `admin` / generată și afișată în log |
| `DATA_DIR` | Baza de date, cache-ul și fișierele încărcate | `./data` (`/data` în Docker) |
| `SECRET_KEY` | Cheia pentru sesiuni | generată în `DATA_DIR` |
| `COOKIE_SECURE` | `1` = cookie de login doar pe HTTPS | `0` (`1` în Docker) |
| `TRUST_PROXY` | `1` când rulează în spatele unui reverse proxy | `0` (`1` în Docker) |
| `MAX_UPLOAD_MB` | Mărimea maximă a unui fișier | `50` |
| `GOOGLE_WORKERS` | Cereri Google în paralel | `4` |
| `CACHE_TTL_DAYS` | Câte zile se păstrează rezultatele în cache | `30` |
| `NOMINATIM_EMAIL` | E-mail de contact trimis către Nominatim (cerut de politica lor) | — |

---

## 💻 Scriptul (linia de comandă)

```bash
pip install -r requirements.txt
set GOOGLE_API_KEY=cheia_ta            # Linux/Mac: export GOOGLE_API_KEY=cheia_ta

python geocodare_od.py                        # alegi fișierul + test pe primele 100 rânduri
python geocodare_od.py --dry-run 0            # tot fișierul
python geocodare_od.py --start 100 --end 500  # doar rândurile 100–500
python geocodare_od.py --single-column        # o singură coloană de adrese (fără destinație)
```

| Parametru | Descriere | Implicit |
|-----------|-----------|----------|
| `--file`, `-f` | Calea fișierului Excel | se deschide o fereastră de alegere |
| `--dry-run` | Procesează doar N rânduri (0 = tot) | 100 |
| `--start` / `--end` | Primul / ultimul rând de procesat | tot fișierul |
| `--single-column` | O singură coloană de adrese; se scriu doar coloanele `…Originea` | — |
| `--api-key` | Cheia Google | variabila `GOOGLE_API_KEY` |

Modul cu o singură coloană pornește și dacă apeși Enter (fără text) când ți se cere coloana DESTINAȚIE.
Scriptul creează lângă fișier `NumeCodat.xlsx` (rezultatul), `Nume_cache.json` (pentru reluare) și `NumeCodat.log`.
Branch-ul `openstreetmap` are o variantă a scriptului care folosește Nominatim în loc de Google.

---

## 🔧 Normalizarea adreselor

Înainte de geocodare, adresele românești sunt curățate (comun pentru site și script):

| Înainte | După |
|---------|------|
| `str pacii` | `Strada pacii` |
| `bd. unirii nr 10` | `Bulevardul unirii nr. 10` |
| `jud. Constanța` | `Județul Constanța` |
| `targu mures` | `Târgu Mureș` |
| `bl 5 sc A et 2` | `bl. 5 sc. A et. 2` |

Adresele care par să conțină două locuri (ex. „colț cu…”, „;”) sunt marcate `Ambiguous` și nu se trimit.

## 📁 Structura

```
geocodare_od.py      scriptul + normalizarea și geocoderul Google (folosite și de site)
web/                 site-ul (Flask): app.py rute, geocoding.py procesare Excel, worker.py coada de joburi
Dockerfile           imaginea pentru server
requirements.txt     dependențe Python
data/                datele site-ului (ignorat de git)
```

Fișierele Excel, documentele și datele locale nu se urcă pe GitHub (vezi `.gitignore`).

## 📜 Licență

MIT
