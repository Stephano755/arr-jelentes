// PDF táblázat beolvasása pdf.js-sel: a cellák téglalapjai + a szövegek pozíciója alapján
// ugyanazt a sor/oszlop szerkezetet állítja elő, mint az asztali változat (pdfplumber).
import * as pdfjsLib from './vendor/pdf.min.mjs';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('./vendor/pdf.worker.min.mjs', import.meta.url).href;
const OPS = pdfjsLib.OPS;

const DATE_ROW_RE = /^(\d{4})-(\d{2})-(\d{2})/;
const GROUP_RE = /^\d{6}([A-Z]{3})-[A-Z]+/;

function mul(m, n) {
  return [m[0] * n[0] + m[1] * n[2], m[0] * n[1] + m[1] * n[3],
          m[2] * n[0] + m[3] * n[2], m[2] * n[1] + m[3] * n[3],
          m[4] * n[0] + m[5] * n[2] + n[4], m[4] * n[1] + m[5] * n[3] + n[5]];
}
const apply = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

async function pageRects(page, H) {
  const ol = await page.getOperatorList();
  let ctm = [1, 0, 0, 1, 0, 0];
  const stack = [], rects = [], seen = new Set();
  for (let i = 0; i < ol.fnArray.length; i++) {
    const fn = ol.fnArray[i], a = ol.argsArray[i];
    if (fn === OPS.save) stack.push(ctm);
    else if (fn === OPS.restore) ctm = stack.pop() || [1, 0, 0, 1, 0, 0];
    else if (fn === OPS.transform) ctm = mul(a, ctm);
    else if (fn === OPS.constructPath) {
      const ops = a[0], co = a[1];
      let k = 0;
      for (const op of ops) {
        if (op === OPS.rectangle) {
          const [x, y, w, h] = co.slice(k, k + 4); k += 4;
          const p1 = apply(ctm, x, y), p2 = apply(ctm, x + w, y + h);
          const r = { x0: Math.min(p1[0], p2[0]), x1: Math.max(p1[0], p2[0]),
                      top: H - Math.max(p1[1], p2[1]), bottom: H - Math.min(p1[1], p2[1]) };
          if (r.x1 - r.x0 < 2 || r.bottom - r.top < 2) continue;
          const key = [r.x0, r.x1, r.top, r.bottom].map(v => Math.round(v * 2)).join(',');
          if (!seen.has(key)) { seen.add(key); rects.push(r); }
        } else if (op === OPS.moveTo || op === OPS.lineTo) k += 2;
        else if (op === OPS.curveTo) k += 6;
        else if (op === OPS.curveTo2 || op === OPS.curveTo3) k += 4;
      }
    }
  }
  return rects;
}

async function pageWords(page, H) {
  const tc = await page.getTextContent();
  const out = [];
  for (const it of tc.items) {
    if (!it.str || !it.str.trim()) continue;
    const h = it.height || Math.abs(it.transform[3]) || 6;
    out.push({ str: it.str, x: it.transform[4], cx: it.transform[4] + (it.width || 0) / 2,
               cy: H - (it.transform[5] + h / 2), y: H - it.transform[5] });
  }
  return out;
}

function textIn(words, r) {
  const ws = words.filter(w => w.cx >= r.x0 - 1 && w.cx <= r.x1 + 1 && w.cy >= r.top - 1 && w.cy <= r.bottom + 1);
  ws.sort((a, b) => (Math.abs(a.y - b.y) > 2 ? a.y - b.y : a.x - b.x));
  const lines = [];
  let last = null;
  for (const w of ws) {
    if (last !== null && Math.abs(w.y - last) <= 2) lines[lines.length - 1].push(w.str);
    else lines.push([w.str]);
    last = w.y;
  }
  return lines.map(l => l.join(' ').replace(/\s+/g, ' ').trim()).join('\n').trim();
}

export function normFlight(raw) {
  let s = (raw || '').toUpperCase().replace(/\s+/g, '');
  const m = s.match(/^([A-Z0-9]{2})\1(\d{1,4}[A-Z]?)$/);
  if (m) s = m[1] + m[2];
  const f = s.match(/([A-Z0-9]{2}\d{1,4}[A-Z]?)/);
  return f ? f[1] : '';
}

/** A PDF feldolgozása: {date, groups:[{code, ship, program, window, pax, bookings, passengers, date}]} */
export async function parsePdf(arrayBuffer) {
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer), isEvalSupported: false }).promise;
  const groups = [];
  let current = null, serviceDate = null;
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const H = page.view[3] - page.view[1];
    const [rects, words] = await Promise.all([pageRects(page, H), pageWords(page, H)]);
    // oszlophatárok a fejléc cellákból
    const head = words.find(w => w.str.startsWith('Group/Position'));
    if (!head) continue;
    const headerCells = rects.filter(r => head.cy >= r.top && head.cy <= r.bottom && r.x1 - r.x0 < 300)
                             .sort((a, b) => a.x0 - b.x0);
    const cols = [];
    for (const r of headerCells) if (!cols.length || r.x0 > cols[cols.length - 1].x0 + 2) cols.push(r);
    if (cols.length < 12) continue;
    const colOf = r => { for (let c = cols.length - 1; c >= 0; c--) if (r.x0 >= cols[c].x0 - 2) return c; return 0; };
    const body = rects.filter(r => r.top >= cols[0].bottom - 1);
    const byCol = Array.from({ length: 12 }, () => []);
    for (const r of body) byCol[colOf(r)].push(r);
    byCol.forEach(list => list.sort((a, b) => a.top - b.top));
    const inside = (r, row) => { const c = (r.top + r.bottom) / 2; return c > row.top && c < row.bottom; };

    for (const row of byCol[0]) {
      const first = textIn(words, row);
      const dm = first.match(DATE_ROW_RE);
      if (dm && row.x1 - row.x0 > (cols[0].x1 - cols[0].x0) * 1.5) {
        serviceDate = `${dm[1]}-${dm[2]}-${dm[3]}`;
        current = null;
        continue;
      }
      const cellText = c => { const r = byCol[c].find(x => inside(x, row)); return r ? textIn(words, r) : ''; };
      const gm = first.match(GROUP_RE);
      if (gm) {
        current = { code: first.split(/\s/)[0], ship: gm[1], program: cellText(1).toUpperCase(),
                    window: cellText(2), route: cellText(3), pax: cellText(4), bookings: cellText(5),
                    passengers: [], date: serviceDate };
        groups.push(current);
      } else if (current && cellText(5)) {
        current.bookings += '\n' + cellText(5);
      }
      if (!current) continue;
      for (const sub of byCol[11].filter(x => inside(x, row))) {
        const t = c => { const r = byCol[c].find(x => inside(x, sub)); return r ? textIn(words, r) : ''; };
        const flight = normFlight(textIn(words, sub));
        if (flight) current.passengers.push({ name: t(6).replace(/\s+/g, ' '), city: t(9), time: t(10), flight });
      }
    }
  }
  return { date: serviceDate, groups };
}

// ---------------------------------------------------------------------------
const TIME_WINDOW_RE = /(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/;
const SEATER_RE = /(\d+)\s*-\s*seater/gi;
const PHONE_TAIL_RE = /\s*\+?\d[\d\s/()-]{5,}$/;

export function shipsIn(parsed) {
  return [...new Set(parsed.groups.filter(g => g.program.includes('TRSF') && !g.program.includes('ASSISTANCE'))
    .map(g => g.ship))].sort();
}

export function escortsFor(parsed, ship) {
  const names = [];
  for (const g of parsed.groups) {
    if (g.ship !== ship || !g.program.includes('ASSISTANCE')) continue;
    for (let line of g.bookings.split('\n')) {
      line = line.trim();
      if (!line.startsWith('-')) continue;
      const name = line.replace(/^[-\s]+/, '').replace(PHONE_TAIL_RE, '').trim();
      if (name && !names.includes(name)) names.push(name);
    }
  }
  return names;
}

export function transfersFor(parsed, ship) {
  const out = [];
  for (const g of parsed.groups) {
    if (g.ship !== ship || !g.program.includes('TRSF') || g.program.includes('ASSISTANCE')) continue;
    const tw = g.window.match(TIME_WINDOW_RE);
    if (!tw) continue;
    const seaters = [...g.bookings.matchAll(SEATER_RE)].map(m => `${m[1]}-seater`);
    const flights = [];
    for (const p of g.passengers) if (!flights.includes(p.flight)) flights.push(p.flight);
    out.push({ start: [+tw[1], +tw[2]], flights, pax: /^\d+$/.test(g.pax) ? +g.pax : null,
               category: seaters.join(' + '), date: g.date || parsed.date });
  }
  out.sort((a, b) => a.start[0] * 60 + a.start[1] - (b.start[0] * 60 + b.start[1]));
  return out;
}

/** Codeshare párosítás a PDF-ből: ugyanabból a városból, ugyanakkor érkező másik járat. */
export function codesharesFromPdf(parsed, airlines) {
  const slots = new Map();
  for (const g of parsed.groups) for (const p of g.passengers) {
    const k = p.city + '|' + p.time;
    if (!slots.has(k)) slots.set(k, new Set());
    slots.get(k).add(p.flight);
  }
  const map = {};
  for (const set of slots.values()) {
    const all = [...set];
    const operators = all.filter(f => !airlines.some(a => f.startsWith(a)))
      .sort((a, b) => parseInt(a.slice(2)) - parseInt(b.slice(2)));
    if (!operators.length) continue;
    for (const f of all) if (airlines.some(a => f.startsWith(a))) map[f] = operators[0];
  }
  return map;
}
