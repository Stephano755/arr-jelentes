import { parsePdf, shipsIn, escortsFor, transfersFor, codesharesFromPdf } from './pdfparse.js';
import { fillTemplate } from './xlsm.js';

const AERODATABOX_HOST = 'aerodatabox.p.rapidapi.com';
const AERODATABOX_MIN_INTERVAL = 1100;
const CODESHARE_AIRLINES = ['DL', 'AA', 'UA'];
const CODESHARE_CACHE_DAYS = 60;
const TRANSFER_DELAY_MIN = 35, TRANSFER_DELAY_MAX = 60;
const REPORT_ROWS = [...Array(20).keys()].map(i => i + 10).concat([31, 32, 33]);
const CLEARED_CELLS = ['K5', 'M5', 'K6', 'M6'];   // kezdés, végzés (a Sum Hours képlet marad)

const $ = id => document.getElementById(id);
const store = {
  get(k, d = null) { try { const v = localStorage.getItem('arr.' + k); return v === null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem('arr.' + k, JSON.stringify(v)); return true; } catch { return false; } },
  del(k) { try { localStorage.removeItem('arr.' + k); } catch { /* */ } },
};

let parsed = null, lastResult = null;

// ---------------------------------------------------------------- napló
function log(msg, cls = '') {
  const line = document.createElement('div');
  line.textContent = msg;
  if (cls) line.className = cls;
  $('log').append(line);
  $('log').scrollTop = $('log').scrollHeight;
}

// ---------------------------------------------------------------- sablon
function b64ToBuf(b64) { const s = atob(b64); const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u.buffer; }
function bufToB64(buf) { let s = ''; const u = new Uint8Array(buf); for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }

const BUILTIN_TEMPLATE = { url: 'sablon.xlsm', name: 'beépített (BASZÓ ARR 2024 v3.2)' };

/** A saját (lecserélt) sablon, ha van; különben a beépített. */
async function getTemplate() {
  const own = store.get('template');
  if (own) return b64ToBuf(own);
  const r = await fetch(BUILTIN_TEMPLATE.url);
  if (!r.ok) throw new Error('A beépített sablon nem tölthető be (nincs internet?).');
  return r.arrayBuffer();
}

function refreshTemplateLabel() {
  const name = store.get('templateName');
  $('tplName').textContent = name || BUILTIN_TEMPLATE.name;
  $('tplReset').hidden = !store.get('template');
  updateGo();
}
$('tplReset').addEventListener('click', () => {
  store.del('template'); store.del('templateName');
  log('Visszaállítva a beépített sablonra.');
  refreshTemplateLabel();
});
$('tplInput').addEventListener('change', async e => {
  const f = e.target.files[0];
  if (!f) return;
  const buf = await f.arrayBuffer();
  if (!store.set('template', bufToB64(buf))) { alert('Nem sikerült elmenteni a sablont (kevés tárhely).'); return; }
  store.set('templateName', f.name);
  log(`Sablon elmentve: ${f.name}`);
  refreshTemplateLabel();
  e.target.value = '';
});

// ---------------------------------------------------------------- API kulcsok
function refreshKeys() {
  const av = store.get('key.aviationstack'), adb = store.get('key.aerodatabox');
  $('keyState').textContent = `AeroDataBox: ${adb ? 'mentve ✓' : 'nincs'} · Aviationstack: ${av ? 'mentve ✓' : 'nincs'}`;
  $('keyAdb').value = ''; $('keyAv').value = '';
  $('keyAdb').placeholder = adb ? '•••••••• (mentve)' : 'AeroDataBox / RapidAPI kulcs';
  $('keyAv').placeholder = av ? '•••••••• (mentve)' : 'Aviationstack kulcs (nem kötelező)';
}
$('saveKeys').addEventListener('click', () => {
  const adb = $('keyAdb').value.trim();
  if (adb) store.set('key.aerodatabox', findRapidKey(adb) || adb);
  if ($('keyAv').value.trim()) store.set('key.aviationstack', $('keyAv').value.trim());
  refreshKeys();
  log('API kulcsok mentve.');
});
$('deleteKeys').addEventListener('click', () => {
  if (!confirm('Biztosan törlöd az elmentett API kulcsokat?')) return;
  store.del('key.aerodatabox'); store.del('key.aviationstack');
  refreshKeys();
  log('API kulcsok törölve.');
  renderWizard();
});

// ---------------------------------------------------------------- kulcsvarázsló
// RapidAPI kulcs formája: …msh…p1…jsn… (kb. 50 karakter). Bármilyen kimásolt szövegből kiszedi.
const RAPID_KEY_RE = /[A-Za-z0-9]{8,12}msh[A-Za-z0-9]{12,20}p1[A-Za-z0-9]{10,14}jsn[A-Za-z0-9]{10,14}/;
function findRapidKey(text) {
  const t = String(text || '');
  const m = t.match(RAPID_KEY_RE);
  if (m) return m[0];
  const s = t.trim();
  return /^[A-Za-z0-9]{40,60}$/.test(s) ? s : null;   // ismeretlen, de kulcsnak tűnő formátum
}

// eszköz felismerése az eszközfüggő varázsló-szövegekhez (iPadOS 13+ Macnek mutatja magát)
document.body.dataset.dev = /iPhone|iPad|iPod/.test(navigator.userAgent) ||
  (/Macintosh/.test(navigator.userAgent) && navigator.maxTouchPoints > 1) ? 'ios'
  : /Android/.test(navigator.userAgent) ? 'android' : 'pc';

let wizForced = false, wizJustDone = false;
const wizStep = () => store.get('wizStep', 1);

function renderWizard() {
  const hasKey = !!store.get('key.aerodatabox');
  const show = wizJustDone || wizForced || (!hasKey && !store.get('wizDismissed'));
  $('wizard').hidden = !show;
  if (!show) return;
  const step = wizStep();
  document.querySelectorAll('#wizard .steps li').forEach(li => {
    const n = +li.dataset.step;
    li.className = wizJustDone || n < step ? 'done' : n === step ? 'now' : 'todo';
  });
  $('wizDone').hidden = !wizJustDone;
  $('wizClose').textContent = wizJustDone ? 'Bezárás' : 'Később';
}
function setWizStep(n) { store.set('wizStep', Math.max(wizStep(), n)); renderWizard(); }

function saveWizardKey(key) {
  store.set('key.aerodatabox', key);
  store.del('wizDismissed');
  wizJustDone = true; wizForced = false;
  refreshKeys();
  renderWizard();
  log(`AeroDataBox kulcs elmentve (…${key.slice(-4)}).`, 'ok');
  $('wizard').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/** Vágólap olvasása. fromTap: a felhasználó koppintott (iPhone-on csak így engedi). */
async function tryClipboard(fromTap) {
  if (!navigator.clipboard?.readText) {
    if (fromTap) $('wizPasteText').textContent = 'Ez a böngésző nem engedi a vágólap olvasását – illeszd be a kulcsot alul kézzel.';
    return false;
  }
  try {
    const key = findRapidKey(await navigator.clipboard.readText());
    if (key) { saveWizardKey(key); return true; }
    if (fromTap) $('wizPasteText').textContent = 'A vágólapon nem találtam RapidAPI kulcsot. Másold ki újra a kulcsot (3. lépés), majd koppints ismét – vagy illeszd be alul kézzel.';
  } catch {
    if (fromTap) $('wizPasteText').textContent = 'A vágólap olvasása nincs engedélyezve – illeszd be a kulcsot alul kézzel (nyomd meg hosszan a mezőt → Beillesztés).';
  }
  return false;
}

document.querySelectorAll('#wizard [data-go]').forEach(el =>
  el.addEventListener('click', () => setTimeout(() => setWizStep(+el.dataset.go), el.tagName === 'A' ? 400 : 0)));
$('wizOpenKeys').addEventListener('click', () => store.set('wizAwaitKey', true));
$('wizPaste').addEventListener('click', () => tryClipboard(true));
$('wizManualSave').addEventListener('click', () => {
  const v = $('wizManual').value;
  const key = findRapidKey(v) || (v.trim().length >= 30 ? v.trim() : null);
  if (!key) { $('wizPasteText').textContent = 'Ez nem tűnik RapidAPI kulcsnak (kb. 50 betű és szám, benne „msh” és „jsn”).'; return; }
  $('wizManual').value = '';
  saveWizardKey(key);
});
$('wizClose').addEventListener('click', () => {
  if (!wizJustDone) store.set('wizDismissed', true);
  wizJustDone = false; wizForced = false;
  renderWizard();
});
$('wizOpen').addEventListener('click', () => {
  wizForced = true; wizJustDone = false;
  store.del('wizDismissed');
  if (!store.get('key.aerodatabox')) store.set('wizStep', Math.min(wizStep(), 3));
  $('settings').open = false;
  renderWizard();
  window.scrollTo({ top: 0, behavior: 'smooth' });
});
// Visszatérés a RapidAPI-ról: ha a kulcs már a vágólapon van, magától elmenti; ha nem, a beillesztés gombra irányít.
document.addEventListener('visibilitychange', async () => {
  if (document.visibilityState !== 'visible' || $('wizard').hidden || wizJustDone || !store.get('wizAwaitKey')) return;
  setWizStep(4);
  if (!(await tryClipboard(false))) {
    $('wizPaste').classList.remove('pulse'); void $('wizPaste').offsetWidth; $('wizPaste').classList.add('pulse');
    $('wizPaste').scrollIntoView({ behavior: 'smooth', block: 'center' });
  }
});

// ---------------------------------------------------------------- PDF
$('pdfInput').addEventListener('change', async e => {
  const f = e.target.files[0];
  if (!f) return;
  $('pdfName').textContent = f.name;
  await loadPdf(await f.arrayBuffer(), f.name);
});

export async function loadPdf(buf, name = 'PDF') {
  $('pdfName').textContent = name;
  try {
    parsed = await parsePdf(buf);
  } catch (err) {
    parsed = null;
    log('Nem sikerült beolvasni a PDF-et: ' + err.message, 'err');
    updateGo();
    return;
  }
  const ships = shipsIn(parsed);
  const sel = $('ship');
  sel.innerHTML = '';
  for (const s of ships) sel.append(new Option(s, s));
  const last = store.get('lastShip');
  if (last && ships.includes(last)) sel.value = last;
  $('date').value = parsed.date || new Date().toISOString().slice(0, 10);
  log(`PDF beolvasva: ${name} – ${parsed.date || '?'}, hajók: ${ships.join(', ') || 'nincs'}`);
  updateEscorts();
  updateGo();
}
$('ship').addEventListener('change', updateEscorts);
function updateEscorts() {
  const names = parsed ? escortsFor(parsed, $('ship').value) : [];
  $('escorts').textContent = names.length ? names.join(', ') : '–';
}
function updateGo() {
  $('go').disabled = !(parsed && $('ship').value);
}

// ---------------------------------------------------------------- API hívás számláló + havi keret
const API_LABELS = { aerodatabox: 'AeroDataBox', aviationstack: 'Aviationstack' };
const DEFAULT_LIMITS = { aerodatabox: null, aviationstack: 100 };   // Aviationstack ingyenes: 100/hó
const monthKey = () => new Date().toISOString().slice(0, 7);
let runCalls = { aerodatabox: 0, aviationstack: 0 };

function countCall(api) {
  runCalls[api]++;
  const all = store.get('apiCalls', {});
  const m = all[monthKey()] || (all[monthKey()] = {});
  m[api] = (m[api] || 0) + 1;
  store.set('apiCalls', all);
  renderStats();
}
const limitOf = api => store.get('limits', {})[api] || store.get('quota', {})[api]?.limit || DEFAULT_LIMITS[api];

function renderStats() {
  const m = store.get('apiCalls', {})[monthKey()] || {};
  const box = $('stats');
  box.innerHTML = '';
  for (const [api, label] of Object.entries(API_LABELS)) {
    const used = m[api] || 0, lim = limitOf(api);
    const q = store.get('quota', {})[api];
    const left = q && q.month === monthKey() && q.remaining != null ? q.remaining : (lim ? Math.max(lim - used, 0) : null);
    const row = document.createElement('div');
    row.className = 'stat';
    const pct = lim ? Math.min(100, Math.round(used / lim * 100)) : 0;
    row.innerHTML = `<div class="stat-top"><b></b><span></span></div><div class="bar"><i></i></div>`;
    row.querySelector('b').textContent = label;
    row.querySelector('span').textContent = lim ? `${used} / ${lim}  ·  maradt ${left}` : `${used} / ? (keret nincs megadva)`;
    row.querySelector('i').style.width = pct + '%';
    if (lim && left / lim < 0.15) row.classList.add('low');
    box.append(row);
  }
  $('statsMonth').textContent = monthKey().replace('-', '.') + '. hónap';
  $('limAdb').placeholder = limitOf('aerodatabox') ? String(limitOf('aerodatabox')) : 'pl. 300';
  $('limAv').placeholder = String(limitOf('aviationstack') || '');
}
$('saveLimits').addEventListener('click', () => {
  const lim = store.get('limits', {});
  for (const [api, el] of [['aerodatabox', $('limAdb')], ['aviationstack', $('limAv')]]) {
    const v = parseInt(el.value, 10);
    if (v > 0) lim[api] = v;
    el.value = '';
  }
  store.set('limits', lim);
  renderStats();
  log('Havi keretek mentve.');
});

// ---------------------------------------------------------------- API hívások
const sleep = ms => new Promise(r => setTimeout(r, ms));
let lastAdb = 0;
let adbAuthFailed = false;   // 401/403 után ebben a futásban nem próbálkozik tovább

async function adbGet(path, key, what) {
  if (adbAuthFailed) return null;
  for (let attempt = 0; attempt < 3; attempt++) {
    const wait = AERODATABOX_MIN_INTERVAL - (Date.now() - lastAdb);
    if (wait > 0) await sleep(wait);
    let r;
    try {
      countCall('aerodatabox');
      r = await fetch(`https://${AERODATABOX_HOST}${path}`, { headers: { 'X-RapidAPI-Key': key, 'X-RapidAPI-Host': AERODATABOX_HOST } });
    } catch (err) {
      log(`  ! AeroDataBox hiba (${what}): ${err.message}`, 'err');
      return null;
    }
    lastAdb = Date.now();
    const rem = r.headers.get('x-ratelimit-requests-remaining');
    if (rem !== null && /^\d+$/.test(rem)) {
      const lim = r.headers.get('x-ratelimit-requests-limit');
      const q = store.get('quota', {});
      q.aerodatabox = { remaining: +rem, limit: /^\d+$/.test(lim || '') ? +lim : null, month: monthKey() };
      store.set('quota', q);
      renderStats();
    }
    if (r.status === 429) { await sleep(2000 + attempt * 2000); continue; }
    if (r.status === 204) return null;
    if (r.status === 401 || r.status === 403) {
      adbAuthFailed = true;
      log(r.status === 401
        ? '  ! AeroDataBox: a kulcs hibás. Ellenőrizd a Beállításokban, vagy futtasd újra a kulcsvarázslót.'
        : '  ! AeroDataBox: nincs előfizetésed az AeroDataBox-ra (403). A RapidAPI-n válaszd az ingyenes Basic csomagot (Útmutató / kulcsvarázsló 2. lépés).', 'err');
      return null;
    }
    if (!r.ok) { log(`  ! AeroDataBox HTTP ${r.status} (${what})`, 'err'); return null; }
    try { return await r.json(); } catch { return null; }
  }
  return null;
}

function bestTime(mv) {
  for (const k of ['revisedTime', 'runwayTime', 'predictedTime', 'scheduledTime']) {
    const m = (mv?.[k]?.local || '').match(/[ T](\d{2}):(\d{2})/);
    if (m) return [+m[1], +m[2]];
  }
  return null;
}

/** Időablakok: első transzfer előtt 3 órától az utolsó után 1 óráig, max. 12 órás darabokban. */
function arrivalWindows(date, transfers) {
  const starts = transfers.map(t => t.start[0] * 60 + t.start[1]);
  const lo = Math.max(0, Math.min(...starts) - 180), hi = Math.min(24 * 60 - 1, Math.max(...starts) + 60);
  const hm = m => `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;
  const wins = [];
  for (let cur = lo; cur < hi;) {
    const nxt = Math.min(cur + 12 * 60 - 1, hi);
    wins.push([`${date}T${hm(cur)}`, `${date}T${hm(nxt)}`]);
    cur = nxt + 1;
  }
  return wins;
}

/** A BUD-ra érkező összes járat – ablakonként 1 hívás. {KL1363: {time, origin, sched, status}} */
async function airportArrivals(windows, key) {
  const out = {};
  for (const [from, to] of windows) {
    const js = await adbGet(`/flights/airports/iata/BUD/${from}/${to}?direction=Arrival&withLeg=false&withCancelled=true` +
      '&withCodeshared=true&withCargo=false&withPrivate=false&withLocation=false', key, `BUD érkezések ${from.slice(11)}-${to.slice(11)}`);
    for (const f of js?.arrivals || []) {
      const num = String(f.number || '').replace(/\s+/g, '').toUpperCase();
      const mv = f.movement || f.arrival;
      if (!num || !mv) continue;
      out[num] = { time: bestTime(mv), origin: mv.airport?.iata, sched: (mv.scheduledTime?.local || '').slice(0, 16),
                   status: f.codeshareStatus || 'Unknown' };
    }
  }
  return out;
}

/** Codeshare -> üzemeltető az érkezési listából (azonos indulási reptér és menetrendi idő). */
function codeshareFromFids(fids, flight) {
  const me = fids[flight];
  if (!me) return null;
  if (me.status === 'IsOperator') return flight;
  const ok = me.status === 'IsCodeshared' ? ['IsOperator', 'Unknown'] : ['IsOperator'];
  const ops = Object.keys(fids).filter(n => n !== flight && ok.includes(fids[n].status)
    && fids[n].origin === me.origin && fids[n].sched === me.sched);
  ops.sort((a, b) => (fids[a].status !== 'IsOperator') - (fids[b].status !== 'IsOperator')
    || parseInt(a.slice(2).replace(/\D/g, '') || 0) - parseInt(b.slice(2).replace(/\D/g, '') || 0));
  return ops[0] || null;
}

const arrivalCache = {};
async function arrivalTime(flight, date, key) {
  const k = flight + date;
  if (!(k in arrivalCache)) {
    let legs = await adbGet(`/flights/number/${flight}/${date}?dateLocalRole=Arrival&withAircraftImage=false&withLocation=false`, key, flight);
    legs = Array.isArray(legs) ? legs : (legs ? [legs] : []);
    const leg = legs.find(l => l?.arrival?.airport?.iata === 'BUD') || legs[0];
    arrivalCache[k] = bestTime(leg?.arrival);
  }
  return arrivalCache[k];
}

async function aviationstackOperating(flight, key) {
  const cache = store.get('codeshareCache', {});
  const hit = cache[flight];
  if (hit && (Date.now() - hit[1]) / 86400000 <= CODESHARE_CACHE_DAYS) return hit[0];
  if (store.get('aviationstackHttpsBlocked')) return null;
  try {
    countCall('aviationstack');
    const r = await fetch(`https://api.aviationstack.com/v1/flights?access_key=${encodeURIComponent(key)}&flight_iata=${flight}&limit=5`);
    const js = await r.json();
    if (js?.error) {
      if (js.error.code === 'https_access_restricted') {
        store.set('aviationstackHttpsBlocked', true);
        log('  Az Aviationstack ingyenes csomagja mobilról nem érhető el – a PDF alapján párosítok.', 'muted');
      } else log(`  ! Aviationstack (${flight}): ${js.error.message || js.error.code}`, 'err');
      return null;
    }
    let op = flight;
    for (const it of js?.data || []) { const cs = it?.flight?.codeshared; if (cs?.flight_iata) { op = cs.flight_iata.toUpperCase(); break; } }
    cache[flight] = [op, Date.now()];
    store.set('codeshareCache', cache);
    return op;
  } catch (err) {
    log(`  ! Aviationstack hiba (${flight}): ${err.message}`, 'err');
    return null;
  }
}

// ---------------------------------------------------------------- jelentés
const fmt = hm => hm ? `${String(hm[0]).padStart(2, '0')}:${String(hm[1]).padStart(2, '0')}` : '–';
const frac = hm => hm ? Math.round((hm[0] * 60 + hm[1]) / 1440 * 1e10) / 1e10 : null;

$('go').addEventListener('click', async () => {
  const ship = $('ship').value;
  const date = $('date').value;          // ÉÉÉÉ-HH-NN
  if (!parsed || !ship || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
  const adbKey = store.get('key.aerodatabox'), avKey = store.get('key.aviationstack');
  if (!adbKey && !confirm('Nincs AeroDataBox kulcs megadva – érkezési idők nélkül készüljön el a jelentés?')) return;
  store.set('lastShip', ship);
  runCalls = { aerodatabox: 0, aviationstack: 0 };
  adbAuthFailed = false;
  $('go').disabled = true;
  $('result').hidden = true;
  $('go').textContent = 'Készül…';
  try {
    const transfers = transfersFor(parsed, ship);
    if (!transfers.length) throw new Error(`A PDF-ben nincs ${ship} transzfer.`);
    const escorts = escortsFor(parsed, ship);
    log(`— ${ship} ${date.replace(/-/g, '.')}: ${transfers.length} transzfer, kísérők: ${escorts.join(', ') || '–'}`);
    if (transfers.length > REPORT_ROWS.length) log(`  ! ${transfers.length} transzfer van, a sablonban csak ${REPORT_ROWS.length} sor.`, 'err');

    // 1) egy hívással az összes BUD érkezés (12 óránként 1 hívás)
    let fids = {};
    if (adbKey) {
      const wins = arrivalWindows(transfers[0].date || date, transfers);
      fids = await airportArrivals(wins, adbKey);
      log(`  BUD érkezési lista: ${Object.keys(fids).length} járat, ${wins.length} hívással`, 'muted');
    }
    // 2) codeshare csere: érkezési lista → PDF párosítás → Aviationstack (csak DL/AA/UA)
    const pdfMap = codesharesFromPdf(parsed, CODESHARE_AIRLINES);
    const mapping = {};
    for (const f of [...new Set(transfers.flatMap(t => t.flights))]) {
      let op = codeshareFromFids(fids, f);
      if (!op && CODESHARE_AIRLINES.some(a => f.startsWith(a))) op = pdfMap[f] || (avKey && await aviationstackOperating(f, avKey));
      mapping[f] = op || f;
      if (mapping[f] !== f) log(`  codeshare: ${f} → ${mapping[f]}`, 'muted');
    }

    const cells = { C7: date.replace(/-/g, '.') };
    for (const ref of CLEARED_CELLS) cells[ref] = null;
    if (escorts[0]) cells.K4 = escorts[0];
    if (escorts.length > 1) cells.M4 = escorts.slice(1).join(', ');

    for (let i = 0; i < Math.min(transfers.length, REPORT_ROWS.length); i++) {
      const t = transfers[i], row = REPORT_ROWS[i];
      const flights = [...new Set(t.flights.map(f => mapping[f] || f))];
      const arrivals = [];
      if (adbKey) for (const f of flights) {
        const a = fids[f]?.time || await arrivalTime(f, t.date || date, adbKey);
        if (a) arrivals.push(a);
      }
      const actual = arrivals.length ? arrivals.reduce((a, b) => (a[0] * 60 + a[1] >= b[0] * 60 + b[1] ? a : b)) : null;
      let left = null;
      if (actual) {
        const mins = actual[0] * 60 + actual[1] + TRANSFER_DELAY_MIN + Math.floor(Math.random() * (TRANSFER_DELAY_MAX - TRANSFER_DELAY_MIN + 1));
        left = [Math.floor(mins / 60) % 24, mins % 60];
      }
      log(`  ${fmt(t.start)}  ${flights.join('/')}  ${t.pax ?? '?'} fő  ${t.category || '?'}  érk.: ${fmt(actual)}  indult: ${fmt(left)}`);
      Object.assign(cells, {
        [`B${row}`]: frac(t.start), [`C${row}`]: flights.join('/'), [`D${row}`]: t.pax,
        [`E${row}`]: t.category || null, [`F${row}`]: frac(actual),
        [`H${row}`]: t.category || null, [`J${row}`]: frac(left),
      });
    }

    const blob = await fillTemplate(await getTemplate(), cells);
    const filename = `${date.replace(/-/g, '.')}.${ship}.xlsm`;
    lastResult = { blob, filename };
    $('resultName').textContent = filename;
    $('result').hidden = false;
    $('share').hidden = !(navigator.canShare && navigator.canShare({ files: [new File([blob], filename, { type: blob.type })] }));
    log(`API hívások most: AeroDataBox ${runCalls.aerodatabox}, Aviationstack ${runCalls.aviationstack}`, 'muted');
    log(`Elkészült: ${filename}`, 'ok');
    $('result').scrollIntoView({ behavior: 'smooth', block: 'center' });
  } catch (err) {
    log('HIBA: ' + err.message, 'err');
    alert(err.message);
  } finally {
    $('go').textContent = 'Jelentés készítése';
    updateGo();
  }
});

$('download').addEventListener('click', () => {
  if (!lastResult) return;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(lastResult.blob);
  a.download = lastResult.filename;
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
});
$('share').addEventListener('click', async () => {
  if (!lastResult) return;
  const file = new File([lastResult.blob], lastResult.filename, { type: lastResult.blob.type });
  try { await navigator.share({ files: [file], title: lastResult.filename }); }
  catch (err) { if (err.name !== 'AbortError') log('Megosztás sikertelen: ' + err.message, 'err'); }
});

// ---------------------------------------------------------------- indulás
refreshTemplateLabel();
refreshKeys();
renderStats();
renderWizard();
if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('./sw.js').catch(() => {});
window.__arr = { loadPdf };   // teszteléshez
