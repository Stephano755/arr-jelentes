// Excel (.xlsm) sablon kitöltése közvetlen XML szerkesztéssel – a makró, a legördülő listák
// és a formázás érintetlen marad (ugyanaz a módszer, mint az asztali változatban).
/* global JSZip */

const escXml = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unescXml = s => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');

const colIndex = col => [...col].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
function colLetters(n) { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
function splitRef(ref) { const m = ref.match(/^([A-Z]+)(\d+)$/); return [m[1], +m[2]]; }

function shiftFormula(formula, origin, target) {
  const [oc, orow] = splitRef(origin), [tc, trow] = splitRef(target);
  const dc = colIndex(tc) - colIndex(oc), dr = trow - orow;
  const re = /(^|[^\w.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(])/g;
  return formula.split(/("(?:[^"]|"")*")/).map(part => part.startsWith('"') ? part :
    part.replace(re, (all, pre, d1, c, d2, r) =>
      pre + d1 + (d1 ? c : colLetters(colIndex(c) + dc)) + d2 + (d2 ? r : String(+r + dr)))).join('');
}

function unshareFormulas(xml) {
  const masters = {};
  const masterRe = /<c r="([A-Z]+\d+)"[^>]*?>\s*<f t="shared" ref="[^"]*" si="(\d+)">([\s\S]*?)<\/f>/g;
  for (const m of xml.matchAll(masterRe)) masters[m[2]] = [m[1], unescXml(m[3])];
  if (!Object.keys(masters).length) return xml;
  xml = xml.replace(masterRe, all => all.replace(/<f t="shared"[^>]*>/, '<f>'));
  return xml.replace(/(<c r="([A-Z]+\d+)"[^>]*?>\s*)<f t="shared" si="(\d+)"\s*\/>/g, (all, pre, ref, si) => {
    const ms = masters[si];
    return ms ? `${pre}<f>${escXml(shiftFormula(ms[1], ms[0], ref))}</f>` : all;
  });
}

function setCell(xml, ref, value) {
  const [col, rownum] = splitRef(ref);
  const cellRe = new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>[\\s\\S]*?</c>)`);
  const m = xml.match(cellRe);
  let style = '';
  if (m) { const sm = m[1].match(/\ss="(\d+)"/); if (sm) style = ` s="${sm[1]}"`; }
  let cell;
  if (value === null || value === undefined || value === '') cell = `<c r="${ref}"${style}/>`;
  else if (typeof value === 'number') cell = `<c r="${ref}"${style}><v>${value}</v></c>`;
  else cell = `<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${escXml(value)}</t></is></c>`;
  if (m) return xml.slice(0, m.index) + cell + xml.slice(m.index + m[0].length);

  const rowRe = new RegExp(`<row r="${rownum}"[^>]*?(?:/>|>([\\s\\S]*?)</row>)`);
  const rm = xml.match(rowRe);
  if (!rm) {
    let pos = xml.indexOf('</sheetData>');
    for (const r of xml.matchAll(/<row r="(\d+)"/g)) if (+r[1] > rownum) { pos = r.index; break; }
    return xml.slice(0, pos) + `<row r="${rownum}">${cell}</row>` + xml.slice(pos);
  }
  if (rm[1] === undefined) {
    return xml.slice(0, rm.index) + rm[0].slice(0, -2) + `>${cell}</row>` + xml.slice(rm.index + rm[0].length);
  }
  const bodyStart = rm.index + rm[0].indexOf('>') + 1;
  for (const c of rm[1].matchAll(/<c r="([A-Z]+)\d+"/g)) {
    if (colIndex(c[1]) > colIndex(col)) {
      const pos = bodyStart + c.index;
      return xml.slice(0, pos) + cell + xml.slice(pos);
    }
  }
  const pos = bodyStart + rm[1].length;
  return xml.slice(0, pos) + cell + xml.slice(pos);
}

async function firstSheetPath(zip) {
  const wb = await zip.file('xl/workbook.xml').async('string');
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const rid = wb.match(/<sheet\b[^>]*\br:id="([^"]+)"/)[1];
  const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map(x => x[0]).find(x => x.includes(`Id="${rid}"`));
  const t = rel.match(/Target="([^"]+)"/)[1].replace(/^\//, '');
  return t.startsWith('xl/') ? t : 'xl/' + t;
}

/** template: ArrayBuffer, cells: {"B10": 0.4, "C10": "LH1334", "K5": null, ...} → Blob */
export async function fillTemplate(template, cells) {
  const zip = await JSZip.loadAsync(template);
  const sheet = await firstSheetPath(zip);
  let xml = unshareFormulas(await zip.file(sheet).async('string'));
  for (const [ref, val] of Object.entries(cells)) xml = setCell(xml, ref, val);
  zip.file(sheet, xml);

  if (zip.file('xl/calcChain.xml')) zip.remove('xl/calcChain.xml');
  let wb = await zip.file('xl/workbook.xml').async('string');
  wb = wb.includes('<calcPr')
    ? wb.replace(/<calcPr\b([^>]*?)(\/?)>/, (a, attrs, sl) => `<calcPr${attrs.replace(/\sfullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"${sl}>`)
    : wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
  zip.file('xl/workbook.xml', wb);
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  zip.file('xl/_rels/workbook.xml.rels', rels.replace(/<Relationship\b[^>]*calcChain[^>]*\/>/g, ''));
  const ct = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', ct.replace(/<Override\b[^>]*calcChain[^>]*\/>/g, ''));

  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE',
    mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12' });
}
