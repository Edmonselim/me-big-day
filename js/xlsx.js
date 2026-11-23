// Read cell values from an .xlsx and write individual cells back without
// disturbing anything else (formatting, dropdowns, conditional colors, formulas).
import { Zip } from './zip.js';

const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
export const unescapeXml = (s) =>
  s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, e) =>
    e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ENT[e]);
export const escapeXml = (s) =>
  String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    // strip characters that are illegal in XML 1.0
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');

export function colToNum(c) { let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
export function numToCol(n) { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - 1 - r) / 26; } return s; }
export function splitRef(ref) { const m = /^([A-Z]+)(\d+)$/.exec(ref); return { col: m[1], row: +m[2] }; }

// Excel serial date (1900 system) <-> ISO yyyy-mm-dd
export function serialToISO(serial) {
  if (serial === null || serial === undefined || serial === '' || isNaN(serial)) return '';
  const ms = Math.round((Number(serial) - 25569) * 86400000);
  return new Date(ms).toISOString().slice(0, 10);
}
export function isoToSerial(iso) {
  if (!iso) return null;
  const [y, m, d] = iso.split('-').map(Number);
  return Date.UTC(y, m - 1, d) / 86400000 + 25569;
}

function attr(tag, name) {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag);
  return m ? m[1] : null;
}

function textOfSi(si) {
  // concatenate every <t> (handles rich text runs); ignore phonetic <rPh>
  const clean = si.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
  let out = '';
  const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g;
  let m;
  while ((m = re.exec(clean))) out += m[1] ? unescapeXml(m[1]) : '';
  return out;
}

export function parseSheetCells(xml, shared) {
  const cells = new Map();
  const re = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  let m;
  while ((m = re.exec(xml))) {
    const head = m[1];
    const body = m[2] || '';
    const ref = attr(head, 'r');
    if (!ref) continue;
    const t = attr(head, 't');
    const fm = /<f\b([^>]*)(?:\/>|>([\s\S]*?)<\/f>)/.exec(body);
    const vm = /<v>([\s\S]*?)<\/v>/.exec(body);
    let v = null;
    if (t === 's') v = vm ? shared[+vm[1]] ?? '' : '';
    else if (t === 'inlineStr') { const im = /<is>([\s\S]*?)<\/is>/.exec(body); v = im ? textOfSi(im[1]) : ''; }
    else if (t === 'str' || t === 'e') v = vm ? unescapeXml(vm[1]) : '';
    else if (t === 'b') v = vm ? vm[1] === '1' : null;
    else v = vm ? Number(vm[1]) : null;
    const cell = { v, t };
    if (fm) {
      cell.f = fm[2] ? unescapeXml(fm[2]) : '';
      const ft = attr(fm[1], 't');
      if (ft === 'shared') { cell.shared = true; cell.si = attr(fm[1], 'si'); cell.master = !!attr(fm[1], 'ref'); }
      if (ft === 'array') cell.array = true;
    }
    if (t === 'e') cell.error = true;
    cells.set(ref, cell);
  }
  return cells;
}

export class Workbook {
  static async open(bytes) {
    const wb = new Workbook();
    wb.zip = await Zip.read(bytes);
    const wbXml = await wb.zip.text('xl/workbook.xml');
    const rels = await wb.zip.text('xl/_rels/workbook.xml.rels');
    const relMap = {};
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = attr(m[0], 'Id');
      let target = attr(m[0], 'Target');
      target = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
      relMap[id] = target;
    }
    wb.sheetPaths = {};
    wb.sheetOrder = [];
    wb.hidden = new Set();
    for (const m of wbXml.matchAll(/<sheet\b[^>]*>/g)) {
      const name = unescapeXml(attr(m[0], 'name'));
      wb.sheetPaths[name] = relMap[attr(m[0], 'r:id')];
      wb.sheetOrder.push(name);
      if (attr(m[0], 'state')) wb.hidden.add(name);
    }
    wb.shared = [];
    const ss = await wb.zip.text('xl/sharedStrings.xml');
    if (ss) for (const m of ss.matchAll(/<si>([\s\S]*?)<\/si>/g)) wb.shared.push(textOfSi(m[1]));
    wb.sheets = {};
    for (const name of wb.sheetOrder) {
      const xml = await wb.zip.text(wb.sheetPaths[name]);
      wb.sheets[name] = parseSheetCells(xml, wb.shared);
    }
    return wb;
  }

  cell(sheet, ref) { return this.sheets[sheet]?.get(ref) || null; }
  value(sheet, ref) { const c = this.cell(sheet, ref); return c ? c.v : null; }

  // Set a cell in the in-memory model (used for instant UI updates before sync).
  setLocal(sheet, ref, value) {
    const s = this.sheets[sheet];
    const old = s.get(ref) || {};
    const isNum = typeof value === 'number' && isFinite(value);
    s.set(ref, { v: value === '' || value === undefined ? null : value, t: isNum ? null : (value === null || value === '' ? null : 'inlineStr') , s: old.s });
  }

  // Apply edits to the underlying XML. edits: [{sheet, ref, value}]
  async applyEdits(edits) {
    const bySheet = new Map();
    for (const e of edits) {
      if (!this.sheetPaths[e.sheet]) throw new Error(`Sheet "${e.sheet}" not found in workbook`);
      if (!bySheet.has(e.sheet)) bySheet.set(e.sheet, []);
      bySheet.get(e.sheet).push(e);
    }
    for (const [sheet, list] of bySheet) {
      const path = this.sheetPaths[sheet];
      let xml = await this.zip.text(path);
      for (const e of list) xml = patchCell(xml, e.ref, e.value);
      this.zip.setText(path, xml);
      this.sheets[sheet] = parseSheetCells(xml, this.shared);
    }
    await this.dropCalcChain();
  }

  // calcChain lists formula cells; replacing a formula with a value can make Excel
  // "repair" the file. Excel rebuilds the chain itself when it is missing.
  async dropCalcChain() {
    if (!this.zip.has('xl/calcChain.xml')) return;
    this.zip.remove('xl/calcChain.xml');
    let ct = await this.zip.text('[Content_Types].xml');
    ct = ct.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, '');
    this.zip.setText('[Content_Types].xml', ct);
    let rels = await this.zip.text('xl/_rels/workbook.xml.rels');
    rels = rels.replace(/<Relationship[^>]*Target="calcChain\.xml"[^>]*\/>/, '');
    this.zip.setText('xl/_rels/workbook.xml.rels', rels);
    let wbx = await this.zip.text('xl/workbook.xml');
    if (!/fullCalcOnLoad/.test(wbx)) {
      wbx = /<calcPr/.test(wbx) ? wbx.replace('<calcPr', '<calcPr fullCalcOnLoad="1"') : wbx.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
      this.zip.setText('xl/workbook.xml', wbx);
    }
  }

  save() { return this.zip.write(); }
}

function cellXml(ref, style, value) {
  const s = style ? ` s="${style}"` : '';
  if (value === null || value === undefined || value === '') return `<c r="${ref}"${s}/>`;
  if (typeof value === 'number' && isFinite(value)) return `<c r="${ref}"${s}><v>${value}</v></c>`;
  if (typeof value === 'boolean') return `<c r="${ref}"${s} t="b"><v>${value ? 1 : 0}</v></c>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}

export function patchCell(xml, ref, value) {
  const { col, row } = splitRef(ref);
  const cnum = colToNum(col);
  const sdStart = xml.indexOf('<sheetData');
  const sdOpenEnd = xml.indexOf('>', sdStart);
  if (xml[sdOpenEnd - 1] === '/') {
    // empty self-closing <sheetData/>
    const fresh = `<sheetData><row r="${row}">${cellXml(ref, null, value)}</row></sheetData>`;
    return xml.slice(0, sdStart) + fresh + xml.slice(sdOpenEnd + 1);
  }
  const sdEnd = xml.indexOf('</sheetData>');
  const before = xml.slice(0, sdOpenEnd + 1);
  const data = xml.slice(sdOpenEnd + 1, sdEnd);
  const after = xml.slice(sdEnd);

  const rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
  let m;
  let insertAt = data.length;
  while ((m = rowRe.exec(data))) {
    const r = +attr(m[1], 'r');
    if (r === row) {
      const head = m[1];
      const body = m[2] || '';
      const newBody = patchRow(body, ref, cnum, value, attr(head, 's'));
      const rowXml = `<row${head}>${newBody}</row>`;
      return before + data.slice(0, m.index) + rowXml + data.slice(m.index + m[0].length) + after;
    }
    if (r > row) { insertAt = m.index; break; }
  }
  const newRow = `<row r="${row}">${cellXml(ref, null, value)}</row>`;
  return before + data.slice(0, insertAt) + newRow + data.slice(insertAt) + after;
}

function patchRow(body, ref, cnum, value, rowStyle) {
  const re = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  let m;
  while ((m = re.exec(body))) {
    const r = attr(m[1], 'r');
    const c = colToNum(splitRef(r).col);
    if (r === ref) {
      const fm = /<f\b([^>]*)/.exec(m[2] || '');
      if (fm && attr(fm[1], 't') === 'shared' && attr(fm[1], 'ref')) {
        throw new Error(`Cell ${ref} holds a shared formula other cells depend on – not editable.`);
      }
      return body.slice(0, m.index) + cellXml(ref, attr(m[1], 's'), value) + body.slice(m.index + m[0].length);
    }
    if (c > cnum) {
      return body.slice(0, m.index) + cellXml(ref, rowStyle, value) + body.slice(m.index);
    }
  }
  return body + cellXml(ref, rowStyle, value);
}
