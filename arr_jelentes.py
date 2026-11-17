# -*- coding: utf-8 -*-
"""
ARR Jelentés – reptéri érkezési transzfer-jelentés készítő
==========================================================
1. A „printout-with-passengers” PDF-ből táblázatként kiolvassa a hajó transzfereit
   (indulási idő, járatok, utaslétszám, jármű kategória) és a Hosp/Escort neveket.
2. AeroDataBox: a BUD érkezési lista egy hívással (12 óránként) – ebből jönnek az érkezési idők
   és a codeshare -> üzemeltető cserék is; ha egy járat nincs benne, csak arra fut külön lekérés.
3. Aviationstack: tartalék a DL / AA / UA codeshare cseréhez, ha a listából nem derül ki.
4. Kitölti az Excel sablont úgy, hogy a makrók, legördülő listák és formázás megmaradnak,
   és az Asztalra menti ÉÉÉÉ.HH.NN.HAJÓ.xlsm néven.
API kulcsok: egyszer kéri be, Windows-felhasználóhoz titkosítva menti, törölhető.
"""
import base64
import ctypes
import ctypes.wintypes as wt
import datetime as dt
import glob
import json
import os
import queue
import random
import re
import shutil
import sys
import threading
import time
import zipfile
from xml.sax.saxutils import escape

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

APP_NAME = "ARR Jelentés"
APP_VERSION = "2.6"
CONFIG_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "ARR-Jelentes")
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")

AERODATABOX_HOST = "aerodatabox.p.rapidapi.com"
AERODATABOX_MIN_INTERVAL = 1.1          # mp két hívás között (ingyenes csomag limit)
CODESHARE_AIRLINES = ("DL", "AA", "UA")
CODESHARE_CACHE_DAYS = 60
TRANSFER_DELAY_MIN = 35                 # perc – az érkezés után ennyivel indul a transzfer
TRANSFER_DELAY_MAX = 60
ARRIVAL_AIRPORT = "BUD"

# Az Excel sablon sorai: 20 fő sor + 3 extra transzfer sor
REPORT_ROWS = list(range(10, 30)) + [31, 32, 33]
# Kezdés (From) és végzés (To) cellák – ezeket üresen hagyjuk; a Sum Hours képlet marad
CLEARED_CELLS = ("K5", "M5", "K6", "M6")

# Fekete téma
BG, BG2, FG, MUTED, ACCENT, BORDER = "#0b0b0c", "#18181b", "#ececf1", "#9b9ba5", "#3b82f6", "#2e2e33"

KEY_NAMES = {
    "aviationstack": "Aviationstack (codeshare csere)",
    "aerodatabox": "AeroDataBox / RapidAPI (érkezési idők)",
}


# ----------------------------------------------------------------------------
#  Segédek: erőforrás, Asztal, titkosítás, konfiguráció
# ----------------------------------------------------------------------------
def resource_path(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def desktop_dir():
    """A valódi Asztal mappa (OneDrive-ra átirányított esetben is)."""
    try:
        class GUID(ctypes.Structure):
            _fields_ = [("d1", wt.DWORD), ("d2", wt.WORD), ("d3", wt.WORD), ("d4", ctypes.c_ubyte * 8)]
        fid = GUID(0xB4BFCC3A, 0xDB2C, 0x424C, (ctypes.c_ubyte * 8)(0xB0, 0x29, 0x7F, 0xE9, 0x9A, 0x87, 0xC6, 0x41))
        p = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(fid), 0, None, ctypes.byref(p)) == 0:
            path = p.value
            ctypes.windll.ole32.CoTaskMemFree(p)
            if path and os.path.isdir(path):
                return path
    except Exception:
        pass
    return os.path.join(os.path.expanduser("~"), "Desktop")


class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wt.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _dpapi(data: bytes, protect: bool) -> bytes:
    buf = ctypes.create_string_buffer(data, len(data))
    blob_in = _Blob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = _Blob()
    fn = ctypes.windll.crypt32.CryptProtectData if protect else ctypes.windll.crypt32.CryptUnprotectData
    if not fn(ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)):
        raise OSError("DPAPI hiba")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


def encrypt(text: str) -> str:
    try:
        return "dpapi:" + base64.b64encode(_dpapi(text.encode("utf-8"), True)).decode()
    except Exception:
        return "plain:" + text


def decrypt(stored: str) -> str:
    if not stored:
        return ""
    if stored.startswith("dpapi:"):
        try:
            return _dpapi(base64.b64decode(stored[6:]), False).decode("utf-8")
        except Exception:
            return ""
    if stored.startswith("plain:"):
        return stored[6:]
    return stored


class Config:
    def __init__(self):
        self.data = {}
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                self.data = json.load(f)
        except Exception:
            self.data = {}
        self.data.setdefault("keys", {})
        self.data.setdefault("codeshare_cache", {})
        self._lock = threading.Lock()

    def save(self):
        with self._lock:
            os.makedirs(CONFIG_DIR, exist_ok=True)
            tmp = CONFIG_PATH + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=1)
            os.replace(tmp, CONFIG_PATH)

    def get_key(self, name):
        env = os.environ.get(f"{name.upper()}_API_KEY", "").strip()
        return env or decrypt(self.data["keys"].get(name, "")).strip()

    def set_key(self, name, value):
        self.data["keys"][name] = encrypt(value.strip())
        self.save()

    def delete_keys(self):
        self.data["keys"] = {}
        self.save()

    # API hívás számláló: {"2026-11": {"aerodatabox": 12, "aviationstack": 3}}
    def count_call(self, api):
        month = dt.date.today().strftime("%Y-%m")
        with self._lock:
            calls = self.data.setdefault("api_calls", {})
            calls.setdefault(month, {}).setdefault(api, 0)
            calls[month][api] += 1

    def month_calls(self):
        return self.data.get("api_calls", {}).get(dt.date.today().strftime("%Y-%m"), {})

    def set_quota(self, api, info):
        with self._lock:
            self.data.setdefault("quota", {})[api] = dict(info, at=dt.datetime.now().isoformat(timespec="minutes"))

    # codeshare gyorsítótár: {"DL9343": ["KL1363", "2026-11-17"]}
    def cached_codeshare(self, flight):
        hit = self.data["codeshare_cache"].get(flight)
        if not hit:
            return None
        try:
            age = (dt.date.today() - dt.date.fromisoformat(hit[1])).days
        except Exception:
            return None
        return hit[0] if age <= CODESHARE_CACHE_DAYS else None

    def store_codeshare(self, flight, operating):
        self.data["codeshare_cache"][flight] = [operating, dt.date.today().isoformat()]


# ----------------------------------------------------------------------------
#  PDF feldolgozás
# ----------------------------------------------------------------------------
DATE_ROW_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})")
GROUP_RE = re.compile(r"^\d{6}([A-Z]{3})-[A-Z]+")
TIME_WINDOW_RE = re.compile(r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})")
SEATER_RE = re.compile(r"(\d+)\s*-\s*seater", re.I)
PHONE_TAIL_RE = re.compile(r"\s*\+?\d[\d\s/()-]{5,}$")


def norm_flight(raw):
    s = re.sub(r"\s+", "", (raw or "").upper())
    m = re.match(r"^([A-Z0-9]{2})\1(\d{1,4}[A-Z]?)$", s)   # pl. "UAUA90\n64" -> UA9064
    if m:
        s = m.group(1) + m.group(2)
    m = re.search(r"([A-Z0-9]{2}\d{1,4}[A-Z]?)", s)
    return m.group(1) if m else ""


def cell(row, i):
    return (row[i] if i < len(row) and row[i] else "").strip()


def parse_pdf(pdf_path):
    """Visszaadja: {'date': date|None, 'groups': [..]} – minden csoport egy táblázatsor."""
    import pdfplumber

    groups, current, service_date = [], None, None
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                for row in table:
                    first = cell(row, 0)
                    if first.startswith("Group/Position"):
                        continue
                    m = DATE_ROW_RE.match(first)
                    if m and not cell(row, 1):
                        service_date = dt.date(int(m[1]), int(m[2]), int(m[3]))
                        current = None
                        continue
                    gm = GROUP_RE.match(first)
                    if gm:
                        current = {
                            "code": first.split()[0],
                            "ship": gm.group(1),
                            "program": cell(row, 1).upper(),
                            "window": cell(row, 2),
                            "route": cell(row, 3),
                            "pax": cell(row, 4),
                            "bookings": cell(row, 5),
                            "passengers": [],
                            "date": service_date,
                        }
                        groups.append(current)
                    elif current is not None and cell(row, 5):
                        current["bookings"] += "\n" + cell(row, 5)   # oldaltörés utáni folytatás
                    if current is not None:
                        flight = norm_flight(cell(row, 11))
                        if flight:
                            current["passengers"].append({
                                "name": " ".join(cell(row, 6).split()),
                                "city": cell(row, 9),
                                "time": cell(row, 10),
                                "flight": flight,
                            })
    return {"date": service_date, "groups": groups}


def ships_in(parsed):
    found = {}
    for g in parsed["groups"]:
        found.setdefault(g["ship"], g["code"])
    return found


def escorts_for(parsed, ship):
    names = []
    for g in parsed["groups"]:
        if g["ship"] != ship or "ASSISTANCE" not in g["program"]:
            continue
        for line in g["bookings"].splitlines():
            line = line.strip()
            if line.startswith("-"):
                name = PHONE_TAIL_RE.sub("", line.lstrip("- ").strip()).strip()
                if name and name not in names:
                    names.append(name)
    return names


def transfers_for(parsed, ship):
    out = []
    for g in parsed["groups"]:
        if g["ship"] != ship or "TRSF" not in g["program"] or "ASSISTANCE" in g["program"]:
            continue
        tw = TIME_WINDOW_RE.search(g["window"])
        if not tw:
            continue
        seaters = [f"{n}-seater" for n in SEATER_RE.findall(g["bookings"])]
        flights = []
        for p in g["passengers"]:
            if p["flight"] not in flights:
                flights.append(p["flight"])
        out.append({
            "start": (int(tw[1]), int(tw[2])),
            "flights": flights,
            "pax": int(g["pax"]) if g["pax"].isdigit() else None,
            "category": " + ".join(seaters),
            "date": g["date"] or parsed["date"],
        })
    out.sort(key=lambda t: t["start"])
    return out


# ----------------------------------------------------------------------------
#  API-k
# ----------------------------------------------------------------------------
class Apis:
    def __init__(self, cfg, log):
        import requests
        self.requests = requests
        self.session = requests.Session()
        self.cfg, self.log = cfg, log
        self._last_adb = 0.0
        self._arrival_cache = {}
        self.run_calls = {"aerodatabox": 0, "aviationstack": 0}
        self._quota_logged = False
        self._quota_blocked = False

    def _count(self, api):
        self.run_calls[api] += 1
        self.cfg.count_call(api)

    # --- Aviationstack: codeshare -> üzemeltető járat ------------------------
    def operating_flight(self, flight, key):
        if not key or not flight.startswith(CODESHARE_AIRLINES):
            return flight
        cached = self.cfg.cached_codeshare(flight)
        if cached:
            return cached
        schemes = ["http"] if self.cfg.data.get("aviationstack_http_only") else ["https", "http"]
        for scheme in schemes:
            try:
                self._count("aviationstack")
                r = self.session.get(f"{scheme}://api.aviationstack.com/v1/flights",
                                     params={"access_key": key, "flight_iata": flight, "limit": 5},
                                     timeout=20)
                js = r.json()
            except Exception as e:
                self.log(f"  ! Aviationstack hiba ({flight}): {e}")
                return flight
            if not isinstance(js, dict):
                self.log(f"  ! Aviationstack ({flight}): váratlan válasz")
                return flight
            err = js.get("error") or {}
            if err.get("code") == "https_access_restricted":
                self.cfg.data["aviationstack_http_only"] = True
                continue
            if err:
                self.log(f"  ! Aviationstack ({flight}): {err.get('message') or err.get('code')}")
                return flight
            operating = flight
            for item in js.get("data") or []:
                cs = (item.get("flight") or {}).get("codeshared") or {}
                if cs.get("flight_iata"):
                    operating = cs["flight_iata"].upper()
                    break
            self.cfg.store_codeshare(flight, operating)
            return operating
        return flight

    # --- AeroDataBox ------------------------------------------------------------
    def _adb_get(self, path, params, key, what):
        """Egy AeroDataBox kérés (ütemezés, számlálás, keret, 429 újrapróbálás). JSON vagy None."""
        st = quota_status(self.cfg, "aerodatabox")
        if self._quota_blocked or (st["exact"] and st["left"] is not None and st["left"] <= 0):
            if not self._quota_blocked:
                self.log("  ! Elfogyott az AeroDataBox keret – nem küldök több kérést (nehogy túllépési díj legyen).")
            self._quota_blocked = True
            return None
        for attempt in range(3):
            wait = AERODATABOX_MIN_INTERVAL - (time.time() - self._last_adb)
            if wait > 0:
                time.sleep(wait)
            try:
                self._count("aerodatabox")
                r = self.session.get(f"https://{AERODATABOX_HOST}{path}", params=params,
                                     headers={"X-RapidAPI-Key": key, "X-RapidAPI-Host": AERODATABOX_HOST},
                                     timeout=30)
            except Exception as e:
                self.log(f"  ! AeroDataBox hiba ({what}): {e}")
                return None
            self._last_adb = time.time()
            q = rapidapi_quota(r.headers)
            if q:
                first = not self._quota_logged
                self._quota_logged = True
                self.cfg.set_quota("aerodatabox", q)
                if first:
                    self.log(f"  AeroDataBox keret (RapidAPI): maradt {q['remaining']} / {q.get('limit') or '?'}"
                             f" {q['unit']}" + (f", megújul: {q['renew'][:10].replace('-', '.')}" if q.get("renew") else ""))
            if q and q["remaining"] <= 0:
                self._quota_blocked = True
            if r.status_code == 429:
                if q and q["remaining"] <= 0:
                    self.log("  ! Elfogyott az AeroDataBox keret – nem küldök több kérést.")
                    return None
                time.sleep(2 + attempt * 2)
                continue
            if r.status_code == 204 or not r.content:
                return None
            if r.status_code != 200:
                self.log(f"  ! AeroDataBox HTTP {r.status_code} ({what})")
                return None
            try:
                return r.json()
            except ValueError:
                return None
        return None

    def airport_arrivals(self, windows, key):
        """A BUD-ra érkező összes járat a megadott (legfeljebb 12 órás) időablakokban – ablakonként 1 hívás.
        Visszaad: {"KL1363": {"time": (11, 24), "origin": "AMS", "sched": "2026-11-17 11:30", "status": "IsOperator"}}"""
        out = {}
        for frm, to in windows:
            js = self._adb_get(
                f"/flights/airports/iata/{ARRIVAL_AIRPORT}/{frm:%Y-%m-%dT%H:%M}/{to:%Y-%m-%dT%H:%M}",
                {"direction": "Arrival", "withLeg": "false", "withCancelled": "true", "withCodeshared": "true",
                 "withCargo": "false", "withPrivate": "false", "withLocation": "false"},
                key, f"BUD érkezések {frm:%H:%M}-{to:%H:%M}")
            for f in (js or {}).get("arrivals") or []:
                num = re.sub(r"\s+", "", str(f.get("number") or "")).upper()
                mv = f.get("movement") or f.get("arrival") or {}
                if not num or not mv:
                    continue
                out[num] = {"time": best_time(mv), "origin": (mv.get("airport") or {}).get("iata"),
                            "sched": ((mv.get("scheduledTime") or {}).get("local") or "")[:16],
                            "status": f.get("codeshareStatus") or "Unknown"}
        return out

    def arrival_time(self, flight, date, key):
        """Egyetlen járat lekérdezése – csak ha az érkezési listában nem szerepelt."""
        if not key or not date:
            return None
        ck = (flight, date)
        if ck not in self._arrival_cache:
            js = self._adb_get(f"/flights/number/{flight}/{date.isoformat()}",
                               {"dateLocalRole": "Arrival", "withAircraftImage": "false", "withLocation": "false"},
                               key, flight)
            legs = js if isinstance(js, list) else ([js] if js else [])
            leg = next((l for l in legs if ((l.get("arrival") or {}).get("airport") or {}).get("iata") == ARRIVAL_AIRPORT),
                       legs[0] if legs else None)
            self._arrival_cache[ck] = best_time((leg or {}).get("arrival") or {})
        return self._arrival_cache[ck]


def best_time(movement):
    """(óra, perc) a legjobb elérhető időből: módosított > leszállás > várható > menetrendi."""
    for k in ("revisedTime", "runwayTime", "predictedTime", "scheduledTime"):
        m = re.search(r"[ T](\d{2}):(\d{2})", (movement.get(k) or {}).get("local") or "")
        if m:
            return int(m[1]), int(m[2])
    return None


def arrival_windows(date, transfers):
    """Időablakok a BUD érkezési listához: az első transzfer előtt 3 órától az utolsó után 1 óráig,
    legfeljebb 12 órás darabokban (egy átlagos nap = 1 hívás)."""
    starts = [t["start"][0] * 60 + t["start"][1] for t in transfers]
    lo = max(0, min(starts) - 180)
    hi = min(24 * 60 - 1, max(starts) + 60)
    day = dt.datetime.combine(date, dt.time())
    wins, cur = [], lo
    while cur < hi:
        nxt = min(cur + 12 * 60 - 1, hi)
        wins.append((day + dt.timedelta(minutes=cur), day + dt.timedelta(minutes=nxt)))
        cur = nxt + 1
    return wins


def codeshare_from_fids(fids, flight):
    """Codeshare -> üzemeltető az érkezési listából (azonos indulási reptér és menetrendi idő)."""
    me = fids.get(flight)
    if not me:
        return None
    if me["status"] == "IsOperator":
        return flight                                  # saját üzemeltetésű, nincs mit cserélni
    ok = ("IsOperator",) if me["status"] != "IsCodeshared" else ("IsOperator", "Unknown")
    ops = [n for n, v in fids.items() if v["status"] in ok and n != flight
           and v["origin"] == me["origin"] and v["sched"] == me["sched"]]
    ops.sort(key=lambda n: (fids[n]["status"] != "IsOperator", int(re.sub(r"\D", "", n[2:]) or 0)))
    return ops[0] if ops else None


# ----------------------------------------------------------------------------
#  Excel írás (közvetlen XML szerkesztés – makró, legördülő lista, formázás marad)
# ----------------------------------------------------------------------------
def col_index(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


def set_cell(xml, ref, value):
    """Beír egy értéket (str/int/float/None) a sheet XML-be, megtartva a cella stílusát."""
    col, rownum = split_ref(ref)
    cell_re = re.compile(r'<c r="%s"(?P<attrs>[^>]*?)(?:/>|>.*?</c>)' % ref, re.S)
    m = cell_re.search(xml)
    style = ""
    if m:
        sm = re.search(r'\ss="(\d+)"', m.group("attrs"))
        style = f' s="{sm.group(1)}"' if sm else ""
    if value is None or value == "":
        new = f'<c r="{ref}"{style}/>'
    elif isinstance(value, (int, float)):
        new = f'<c r="{ref}"{style}><v>{repr(float(value)) if isinstance(value, float) else value}</v></c>'
    else:
        new = f'<c r="{ref}"{style} t="inlineStr"><is><t xml:space="preserve">{escape(str(value))}</t></is></c>'
    if m:
        return xml[:m.start()] + new + xml[m.end():]
    # nincs ilyen cella: beszúrás a megfelelő sorba / oszlopsorrendbe
    row_re = re.compile(r'<row r="%d"[^>]*?(?:/>|>(?P<body>.*?)</row>)' % rownum, re.S)
    rm = row_re.search(xml)
    if not rm:
        rows = [(int(x.group(1)), x.start()) for x in re.finditer(r'<row r="(\d+)"', xml)]
        pos = next((p for r, p in rows if r > rownum), xml.index("</sheetData>"))
        return xml[:pos] + f'<row r="{rownum}">{new}</row>' + xml[pos:]
    if rm.group("body") is None:
        return xml[:rm.start()] + rm.group(0)[:-2] + f">{new}</row>" + xml[rm.end():]
    body_start = rm.start("body")
    for cm in re.finditer(r'<c r="([A-Z]+)\d+"', rm.group("body")):
        if col_index(cm.group(1)) > col_index(col):
            pos = body_start + cm.start()
            return xml[:pos] + new + xml[pos:]
    pos = rm.end("body")
    return xml[:pos] + new + xml[pos:]


REF_RE = re.compile(r"(?<![\w.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(])")


def col_letters(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def shift_formula(formula, origin, target):
    """A relatív cellahivatkozásokat eltolja origin -> target között (szövegkonstansokat kihagyva)."""
    oc, orow = split_ref(origin)
    tc, trow = split_ref(target)
    dc, dr = col_index(tc) - col_index(oc), trow - orow

    def shift(m):
        c = m.group(2) if m.group(1) else col_letters(col_index(m.group(2)) + dc)
        r = m.group(4) if m.group(3) else str(int(m.group(4)) + dr)
        return f"{m.group(1)}{c}{m.group(3)}{r}"
    parts = re.split(r'("(?:[^"]|"")*")', formula)
    return "".join(p if p.startswith('"') else REF_RE.sub(shift, p) for p in parts)


def unshare_formulas(xml):
    """A megosztott (shared) képleteket önálló képletekké alakítja, hogy a felülírt
    „mester” cella ne rontsa el a tartomány többi celláját."""
    from html import unescape

    masters = {}
    master_re = re.compile(r'<c r="([A-Z]+\d+)"[^>]*?>\s*<f t="shared" ref="[^"]*" si="(\d+)">(.*?)</f>', re.S)
    for m in master_re.finditer(xml):
        masters[m.group(2)] = (m.group(1), unescape(m.group(3)))
    if not masters:
        return xml

    def fix_master(m):
        return m.group(0).replace(re.search(r'<f t="shared"[^>]*>', m.group(0)).group(0), "<f>")
    xml = master_re.sub(fix_master, xml)

    dep_re = re.compile(r'(<c r="([A-Z]+\d+)"[^>]*?>\s*)<f t="shared" si="(\d+)"\s*/>', re.S)

    def fix_dep(m):
        anchor, text = masters.get(m.group(3), (None, None))
        if anchor is None:
            return m.group(0)
        formula = shift_formula(text, anchor, m.group(2))
        return f"{m.group(1)}<f>{escape(formula)}</f>"
    return dep_re.sub(fix_dep, xml)


def time_fraction(hm):
    return round((hm[0] * 60 + hm[1]) / 1440.0, 10) if hm else None


def first_sheet_path(z):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid = re.search(r'<sheet\b[^>]*\br:id="([^"]+)"', wb).group(1)
    target = re.search(r'<Relationship\b[^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or \
        re.search(r'<Relationship\b[^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    t = target.group(1).lstrip("/")
    return t if t.startswith("xl/") else "xl/" + t


def write_report(template, out_path, cells):
    with zipfile.ZipFile(template) as zin:
        sheet = first_sheet_path(zin)
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    out_items = []
    for info, data in items:
        name = info.filename
        if name == "xl/calcChain.xml":
            continue                       # Excel újraépíti; a felülírt képletek miatt kell
        if name == sheet:
            xml = unshare_formulas(data.decode("utf-8"))
            for ref, val in cells.items():
                xml = set_cell(xml, ref, val)
            data = xml.encode("utf-8")
        elif name == "xl/workbook.xml":
            xml = data.decode("utf-8")
            if "<calcPr" in xml:
                xml = re.sub(r"<calcPr\b([^>]*?)(/?)>",
                             lambda m: "<calcPr" + re.sub(r'\sfullCalcOnLoad="[^"]*"', "", m.group(1)) +
                             ' fullCalcOnLoad="1"' + m.group(2) + ">", xml, count=1)
            else:
                xml = xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
            data = xml.encode("utf-8")
        elif name == "xl/_rels/workbook.xml.rels":
            data = re.sub(rb'<Relationship\b[^>]*calcChain[^>]*/>', b"", data)
        elif name == "[Content_Types].xml":
            data = re.sub(rb'<Override\b[^>]*calcChain[^>]*/>', b"", data)
        out_items.append((info, data))
    tmp = out_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in out_items:
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            zout.writestr(zi, data)
    os.replace(tmp, out_path)


def unique_path(path):
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 2
    while os.path.exists(f"{base} ({n}){ext}"):
        n += 1
    return f"{base} ({n}){ext}"


# ----------------------------------------------------------------------------
#  Feldolgozás
# ----------------------------------------------------------------------------
def build_report(cfg, pdf_path, ship, date, template, keys, log, parsed=None):
    parsed = parsed or parse_pdf(pdf_path)
    transfers = transfers_for(parsed, ship)
    if not transfers:
        raise RuntimeError(f"A PDF-ben nincs {ship} transzfer (TRSF) sor.")
    escorts = escorts_for(parsed, ship)
    log(f"{ship}: {len(transfers)} transzfer, dolgozók: {', '.join(escorts) or '–'}")
    if len(transfers) > len(REPORT_ROWS):
        log(f"  ! {len(transfers)} transzfer van, a sablonban csak {len(REPORT_ROWS)} sor – a többi kimarad!")

    apis = Apis(cfg, log)
    adb_key = keys.get("aerodatabox")
    # 1) egy hívással az összes BUD érkezés a nap érintett részére (12 óránként 1 hívás)
    fids = {}
    if adb_key:
        wins = arrival_windows(transfers[0]["date"] or date, transfers)
        fids = apis.airport_arrivals(wins, adb_key)
        log(f"  BUD érkezési lista: {len(fids)} járat, {len(wins)} hívással")
    # 2) codeshare csere: előbb az érkezési listából, csak ha ott nincs meg, Aviationstack
    all_flights = sorted({f for t in transfers for f in t["flights"]})
    mapping = {}
    for f in all_flights:
        op = codeshare_from_fids(fids, f)          # bármely codeshare-t cserél, ha a listából kiderül
        if op and op != f:
            cfg.store_codeshare(f, op)
        if not op:
            op = apis.operating_flight(f, keys.get("aviationstack")) if f.startswith(CODESHARE_AIRLINES) else f
        mapping[f] = op
        if mapping[f] != f:
            log(f"  codeshare: {f} → {mapping[f]}")
    cfg.save()

    cells = {"C7": date.strftime("%Y.%m.%d")}
    # kezdés / végzés (From, To) törlése mindkét képviselőnél (Sum Hours képlet marad)
    for ref in CLEARED_CELLS:
        cells[ref] = None
    if escorts:
        cells["K4"] = escorts[0]
    if len(escorts) > 1:
        cells["M4"] = ", ".join(escorts[1:])

    for t, row in zip(transfers, REPORT_ROWS):
        flights = []
        for f in t["flights"]:
            op = mapping.get(f, f)
            if op not in flights:
                flights.append(op)
        # 2) érkezési idők – a transzfer az utolsó érkező járat után indul
        arrivals = []
        for f in flights:
            a = (fids.get(f) or {}).get("time") or apis.arrival_time(f, t["date"] or date, adb_key)
            if a:
                arrivals.append(a)
        actual = max(arrivals) if arrivals else None
        left = None
        if actual:
            mins = actual[0] * 60 + actual[1] + random.randint(TRANSFER_DELAY_MIN, TRANSFER_DELAY_MAX)
            left = ((mins // 60) % 24, mins % 60)
        hhmm = lambda x: f"{x[0]:02d}:{x[1]:02d}" if x else "–"
        log(f"  {hhmm(t['start'])}  {'/'.join(flights):<22} {t['pax'] or '?':>3} fő  "
            f"{t['category'] or '?':<10} érk.: {hhmm(actual)}  indult: {hhmm(left)}")
        cells.update({
            f"B{row}": time_fraction(t["start"]),
            f"C{row}": "/".join(flights),
            f"D{row}": t["pax"],
            f"E{row}": t["category"] or None,
            f"F{row}": time_fraction(actual),
            f"H{row}": t["category"] or None,
            f"J{row}": time_fraction(left),
        })

    cfg.save()
    log(f"API hívások most: AeroDataBox {apis.run_calls['aerodatabox']}, "
        f"Aviationstack {apis.run_calls['aviationstack']}")
    out = unique_path(os.path.join(desktop_dir(), f"{date.strftime('%Y.%m.%d')}.{ship}.xlsm"))
    write_report(template, out, cells)
    return out


# AeroDataBox Basic (ingyenes): 400 egység/hó, egy TIER 2 hívás 2 egység → 200 hívás.
# Aviationstack ingyenes csomag: 100 hívás/hó.
DEFAULT_LIMITS = {"aerodatabox": 200, "aviationstack": 100}
API_LABELS = {"aerodatabox": "AeroDataBox", "aviationstack": "Aviationstack"}


def rapidapi_quota(headers):
    """A RapidAPI keret-fejlécei: x-ratelimit-<objektum>-limit / -remaining / -reset.
    Ha a csomag egységekben (pl. TIER) számol, azt részesíti előnyben a sima kérésszámmal szemben."""
    found = {}
    for k, v in headers.items():
        m = re.match(r"x-ratelimit-(.+)-(limit|remaining|reset)$", k.lower())
        if m and str(v).strip().isdigit():
            found.setdefault(m.group(1), {})[m.group(2)] = int(v)
    names = [n for n, d in found.items() if "remaining" in d]
    if not names:
        return None
    name = next((n for n in names if n != "requests"), names[0])
    d = found[name]
    info = {"remaining": d["remaining"], "limit": d.get("limit"),
            "unit": "kérés" if name == "requests" else name.replace("-", " ")}
    if d.get("reset"):
        info["renew"] = (dt.datetime.now() + dt.timedelta(seconds=d["reset"])).isoformat(timespec="minutes")
    return info


def monthly_limit(cfg, api):
    return cfg.data.get("limits", {}).get(api) or DEFAULT_LIMITS.get(api)


def quota_status(cfg, api):
    """{used, limit, left, unit, renew, calls, exact} – a RapidAPI pontos adata, ha friss; különben a helyi számláló."""
    calls = cfg.month_calls().get(api, 0)
    q = cfg.data.get("quota", {}).get(api) or {}
    now = dt.datetime.now()
    fresh = q.get("remaining") is not None and (
        (q.get("renew") and now.isoformat() < q["renew"]) or
        (not q.get("renew") and q.get("at", "")[:7] == now.strftime("%Y-%m")))
    if fresh and (q.get("limit") or monthly_limit(cfg, api)):
        lim = q.get("limit") or monthly_limit(cfg, api)
        return {"used": max(lim - q["remaining"], 0), "limit": lim, "left": q["remaining"],
                "unit": q.get("unit", "kérés"), "renew": q.get("renew"), "calls": calls, "exact": True}
    lim = monthly_limit(cfg, api)
    return {"used": calls, "limit": lim, "left": max(lim - calls, 0) if lim else None,
            "unit": "hívás", "renew": None, "calls": calls, "exact": False}


def quota_text(cfg, api):
    s = quota_status(cfg, api)
    if not s["limit"]:
        return f"{s['calls']} hívás ebben a hónapban · keret: ismeretlen (első valódi futás után látszik)"
    txt = f"{s['used']} / {s['limit']} {s['unit']} · maradt {s['left']}"
    if s["renew"]:
        txt += f" · megújul: {s['renew'][:10].replace('-', '.')}."
    if s["exact"]:
        txt += f"  ({s['calls']} hívás e hónapban)"
    return txt


def api_stats_text(cfg):
    return "   ·   ".join(f"{label}: {quota_text(cfg, api)}" for api, label in API_LABELS.items())


# ----------------------------------------------------------------------------
#  Grafikus felület
# ----------------------------------------------------------------------------
def apply_dark_theme(root):
    """Fekete téma a ttk elemekre, a sima tk elemekre (párbeszédablakok) és a címsorra."""
    root.configure(bg=BG)
    for pat, val in (("*Background", BG), ("*Foreground", FG), ("*Entry.Background", BG2),
                     ("*Entry.Foreground", FG), ("*Entry.insertBackground", FG),
                     ("*Button.Background", BG2), ("*Button.Foreground", FG),
                     ("*Button.activeBackground", BORDER), ("*Button.activeForeground", FG),
                     ("*selectBackground", ACCENT), ("*selectForeground", "#ffffff"),
                     ("*TCombobox*Listbox.background", BG2), ("*TCombobox*Listbox.foreground", FG),
                     ("*TCombobox*Listbox.selectBackground", ACCENT),
                     ("*TCombobox*Listbox.selectForeground", "#ffffff")):
        root.option_add(pat, val)
    st = ttk.Style(root)
    st.theme_use("clam")
    st.configure(".", background=BG, foreground=FG, fieldbackground=BG2, bordercolor=BORDER,
                 lightcolor=BG2, darkcolor=BG2, troughcolor=BG2, insertcolor=FG, font=("Segoe UI", 10))
    st.configure("TLabel", background=BG, foreground=FG)
    st.configure("Muted.TLabel", foreground=MUTED)
    st.configure("TFrame", background=BG)
    st.configure("TButton", background=BG2, foreground=FG, padding=(12, 6), borderwidth=1)
    st.map("TButton", background=[("active", BORDER), ("disabled", BG)], foreground=[("disabled", MUTED)])
    st.configure("Accent.TButton", background=ACCENT, foreground="#ffffff", bordercolor=ACCENT)
    st.map("Accent.TButton", background=[("active", "#2563eb"), ("disabled", BORDER)])
    st.configure("TEntry", fieldbackground=BG2, foreground=FG, insertcolor=FG)
    st.configure("TCombobox", fieldbackground=BG2, background=BG2, foreground=FG, arrowcolor=FG)
    st.map("TCombobox", fieldbackground=[("readonly", BG2)], foreground=[("readonly", FG)],
           selectbackground=[("!focus", BG2)], selectforeground=[("!focus", FG)])
    for name, color in (("Ok", ACCENT), ("Low", "#fbbf24"), ("Crit", "#f87171")):
        st.configure(f"{name}.Horizontal.TProgressbar", background=color, troughcolor=BG2, bordercolor=BORDER,
                     lightcolor=color, darkcolor=color, thickness=8)
    try:   # sötét címsor (Windows 10 20H1+ / 11)
        root.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        val = ctypes.c_int(1)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(val), ctypes.sizeof(val))
    except Exception:
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.minsize(640, 460)
        apply_dark_theme(self)
        self.cfg = Config()
        self.parsed = None
        self.msgs = queue.Queue()

        self.pdf_var = tk.StringVar()
        self.ship_var = tk.StringVar(value=self.cfg.data.get("last_ship", ""))
        self.date_var = tk.StringVar()
        self.tpl_var = tk.StringVar()
        self.escort_var = tk.StringVar(value="–")

        pad = dict(padx=6, pady=4)
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="PDF:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.pdf_var).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Tallózás…", command=self.pick_pdf).grid(row=0, column=2, **pad)

        ttk.Label(frm, text="Hajó:").grid(row=1, column=0, sticky="w", **pad)
        self.ship_box = ttk.Combobox(frm, textvariable=self.ship_var, width=12)
        self.ship_box.grid(row=1, column=1, sticky="w", **pad)
        self.ship_box.bind("<<ComboboxSelected>>", lambda e: self.update_escorts())
        self.ship_box.bind("<KeyRelease>", lambda e: self.update_escorts())

        ttk.Label(frm, text="Dátum:").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.date_var, width=14).grid(row=2, column=1, sticky="w", **pad)

        ttk.Label(frm, text="Dolgozók:").grid(row=3, column=0, sticky="w", **pad)
        ttk.Label(frm, textvariable=self.escort_var).grid(row=3, column=1, columnspan=2, sticky="w", **pad)

        ttk.Label(frm, text="Sablon:").grid(row=4, column=0, sticky="w", **pad)
        ttk.Label(frm, textvariable=self.tpl_var).grid(row=4, column=1, sticky="w", **pad)
        ttk.Button(frm, text="Másik sablon…", command=self.pick_template).grid(row=4, column=2, **pad)

        # API keret: név + kitöltődő csík + szöveg, API-nként egy sor
        qf = ttk.Frame(frm)
        qf.grid(row=7, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))
        qf.columnconfigure(2, weight=1)
        self.quota_widgets = {}
        for i, (api, label) in enumerate(API_LABELS.items()):
            ttk.Label(qf, text=label).grid(row=i, column=0, sticky="w", padx=(0, 10), pady=3)
            bar = ttk.Progressbar(qf, length=170, maximum=100, style="Ok.Horizontal.TProgressbar")
            bar.grid(row=i, column=1, sticky="w", pady=3)
            var = tk.StringVar()
            ttk.Label(qf, textvariable=var, style="Muted.TLabel").grid(row=i, column=2, sticky="w", padx=(10, 0), pady=3)
            self.quota_widgets[api] = (bar, var)
        ttk.Button(frm, text="Havi keretek…", command=self.edit_limits).grid(row=7, column=2, sticky="ne", **pad)
        self.refresh_quota()

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(8, 4))
        self.go_btn = ttk.Button(btns, text="Jelentés készítése", command=self.run, style="Accent.TButton")
        self.go_btn.pack(side="left", padx=6)
        ttk.Button(btns, text="API kulcsok megadása", command=self.edit_keys).pack(side="left", padx=6)
        ttk.Button(btns, text="API kulcsok törlése", command=self.delete_keys).pack(side="left", padx=6)

        self.logbox = tk.Text(frm, height=14, wrap="word", font=("Consolas", 9), state="disabled",
                              bg=BG2, fg=FG, insertbackground=FG, relief="flat", highlightthickness=1,
                              highlightbackground=BORDER, highlightcolor=BORDER, padx=8, pady=6)
        self.logbox.grid(row=6, column=0, columnspan=3, sticky="nsew", **pad)
        frm.rowconfigure(6, weight=1)

        self.refresh_template_label()
        self.after(100, self.pump)
        self.autoload_pdf()

    # --- napló ---------------------------------------------------------------
    def log(self, msg):
        self.msgs.put(("log", msg))

    def pump(self):
        try:
            while True:
                kind, payload = self.msgs.get_nowait()
                if kind == "log":
                    self.logbox.configure(state="normal")
                    self.logbox.insert("end", payload + "\n")
                    self.logbox.see("end")
                    self.logbox.configure(state="disabled")
                elif kind == "done":
                    self.go_btn.configure(state="normal")
                    self.refresh_quota()
                    if payload.get("error"):
                        messagebox.showerror("Hiba", payload["error"])
                    else:
                        out = payload["out"]
                        if messagebox.askyesno("Kész", f"Elkészült:\n{out}\n\nMegnyissam?"):
                            os.startfile(out)
        except queue.Empty:
            pass
        self.after(100, self.pump)

    # --- sablon --------------------------------------------------------------
    def template_path(self):
        p = self.cfg.data.get("template")
        if p and os.path.exists(p):
            return p
        return resource_path("sablon.xlsm")

    def refresh_template_label(self):
        p = self.cfg.data.get("template")
        self.tpl_var.set(os.path.basename(p) if p and os.path.exists(p) else "beépített (BASZÓ ARR 2024 v3.2)")

    def pick_template(self):
        p = filedialog.askopenfilename(title="Excel sablon", filetypes=[("Excel makrós", "*.xlsm")])
        if p:
            self.cfg.data["template"] = p
            self.cfg.save()
            self.refresh_template_label()

    # --- PDF -----------------------------------------------------------------
    def autoload_pdf(self):
        dl = os.path.join(os.path.expanduser("~"), "Downloads")
        pdfs = sorted(glob.glob(os.path.join(dl, "printout-with-passengers*.pdf")), key=os.path.getmtime)
        if pdfs and time.time() - os.path.getmtime(pdfs[-1]) < 3 * 86400:
            self.load_pdf(pdfs[-1])

    def pick_pdf(self):
        p = filedialog.askopenfilename(title="Válaszd ki a PDF fájlt", filetypes=[("PDF", "*.pdf")],
                                       initialdir=os.path.join(os.path.expanduser("~"), "Downloads"))
        if p:
            self.load_pdf(p)

    def load_pdf(self, path):
        self.pdf_var.set(path)
        try:
            self.parsed = parse_pdf(path)
        except Exception as e:
            self.parsed = None
            messagebox.showerror("Hiba", f"Nem sikerült beolvasni a PDF-et:\n{e}")
            return
        ships = sorted({g["ship"] for g in self.parsed["groups"]
                        if "TRSF" in g["program"] and "ASSISTANCE" not in g["program"]})
        self.ship_box["values"] = ships
        if self.ship_var.get().upper() not in ships and ships:
            self.ship_var.set(ships[0])
        d = self.parsed["date"] or dt.date.today()
        self.date_var.set(d.strftime("%Y.%m.%d"))
        self.log(f"PDF beolvasva: {os.path.basename(path)} – {d:%Y.%m.%d}, hajók: {', '.join(ships)}")
        self.update_escorts()

    def update_escorts(self):
        ship = self.ship_var.get().strip().upper()
        names = escorts_for(self.parsed, ship) if self.parsed and ship else []
        self.escort_var.set(", ".join(names) if names else "–")

    # --- API kulcsok -----------------------------------------------------------
    def ask_key(self, name, force=False):
        key = self.cfg.get_key(name)
        if key and not force:
            return key
        val = simpledialog.askstring("API kulcs", f"{KEY_NAMES[name]} API kulcs:\n(üresen hagyva kihagyja)",
                                     parent=self, show="*")
        if val and val.strip():
            self.cfg.set_key(name, val)
            return val.strip()
        return key

    def edit_keys(self):
        for name in KEY_NAMES:
            self.ask_key(name, force=True)
        self.log("API kulcsok mentve.")

    def edit_limits(self):
        limits = self.cfg.data.setdefault("limits", {})
        for api, label in API_LABELS.items():
            cur = monthly_limit(self.cfg, api)
            val = simpledialog.askinteger("Havi keret", f"{label} havi hívás-kerete:\n(üresen hagyva marad: {cur or 'nincs'})",
                                          parent=self, minvalue=1, initialvalue=cur)
            if val:
                limits[api] = val
        self.cfg.save()
        self.refresh_quota()

    def refresh_quota(self):
        for api, (bar, var) in self.quota_widgets.items():
            s = quota_status(self.cfg, api)
            pct = min(100, s["used"] / s["limit"] * 100) if s["limit"] else 0
            bar.configure(value=pct, style=("Crit" if pct >= 90 else "Low" if pct >= 75 else "Ok") + ".Horizontal.TProgressbar")
            var.set(quota_text(self.cfg, api))

    def delete_keys(self):
        if messagebox.askyesno("API kulcsok törlése", "Biztosan törlöd az elmentett API kulcsokat?"):
            self.cfg.delete_keys()
            self.log("API kulcsok törölve.")

    # --- futtatás --------------------------------------------------------------
    def run(self):
        pdf = self.pdf_var.get().strip()
        ship = self.ship_var.get().strip().upper()
        if not pdf or not os.path.exists(pdf):
            messagebox.showwarning("Hiányzó PDF", "Válaszd ki a PDF fájlt.")
            return
        if not re.fullmatch(r"[A-Z]{3}", ship):
            messagebox.showwarning("Hajó", "Add meg a hajó 3 betűs kódját (pl. VLH).")
            return
        try:
            date = dt.datetime.strptime(re.sub(r"[-/ ]", ".", self.date_var.get().strip()).rstrip("."), "%Y.%m.%d").date()
        except ValueError:
            messagebox.showwarning("Dátum", "A dátum formátuma: ÉÉÉÉ.HH.NN")
            return
        if self.parsed is None or self.pdf_var.get() != pdf:
            self.load_pdf(pdf)
        keys = {n: self.ask_key(n) for n in KEY_NAMES}
        self.cfg.data["last_ship"] = ship
        self.cfg.save()
        self.go_btn.configure(state="disabled")
        self.log(f"--- {ship} {date:%Y.%m.%d} ---")

        def work():
            try:
                out = build_report(self.cfg, pdf, ship, date, self.template_path(), keys, self.log, self.parsed)
                self.log(f"Mentve: {out}")
                self.msgs.put(("done", {"out": out}))
            except Exception as e:
                self.log(f"HIBA: {e}")
                self.msgs.put(("done", {"error": str(e)}))

        threading.Thread(target=work, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()
