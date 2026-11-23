// Small UI helpers: escaping, money formatting, icons, bottom sheets, toasts, generic forms.
import { serialToISO, isoToSerial } from './xlsx.js';

export const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
export const $ = (sel, root = document) => root.querySelector(sel);
export const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

const nf0 = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 });
export const usd = (n) => (n < 0 ? '-$' : '$') + nf0.format(Math.abs(Math.round(n || 0)));
export const egp = (n) => 'EGP ' + nf0.format(Math.round(n || 0));
export const fmtNum = (n) => nf0.format(Math.round(n || 0));

// USD first, EGP underneath (all workbook amounts are EGP)
export function money(egpAmount, rate, cls = '') {
  return `<span class="money ${cls}"><b>${usd(egpAmount / rate)}</b><small>${egp(egpAmount)}</small></span>`;
}

export function fmtDate(iso, opts = { month: 'short', day: 'numeric', year: 'numeric' }) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, m - 1, d).toLocaleDateString('en-US', opts);
}

const P = (d) => `<svg viewBox="0 0 24 24" fill="none" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${d}</svg>`;
export const ICON = {
  home: P('<path d="M3 10.5 12 3l9 7.5"/><path d="M5 9.5V21h14V9.5"/><path d="M10 21v-6h4v6"/>'),
  money: P('<rect x="2.5" y="6" width="19" height="13" rx="2.5"/><circle cx="12" cy="12.5" r="2.8"/><path d="M6 9.5h.01M18 15.5h.01"/>'),
  vendors: P('<path d="M4 7h16l-1 13H5L4 7Z"/><path d="M8.5 7a3.5 3.5 0 0 1 7 0"/>'),
  guests: P('<circle cx="9" cy="8" r="3.2"/><path d="M3 20c0-3.3 2.7-6 6-6s6 2.7 6 6"/><circle cx="17" cy="9" r="2.5"/><path d="M15.5 14.2c3 .2 5.5 2.5 5.5 5.8"/>'),
  more: P('<circle cx="5" cy="12" r="1.6"/><circle cx="12" cy="12" r="1.6"/><circle cx="19" cy="12" r="1.6"/>'),
  plus: P('<path d="M12 5v14M5 12h14"/>'),
  sync: P('<path d="M20 11a8 8 0 0 0-14.5-4.5L4 8"/><path d="M4 4v4h4"/><path d="M4 13a8 8 0 0 0 14.5 4.5L20 16"/><path d="M20 20v-4h-4"/>'),
  chev: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m9 6 6 6-6 6"/></svg>',
  check: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round"><path d="m5 12 5 5 9-10"/></svg>',
  search: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#a3929a" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/></svg>',
  ms: '<svg viewBox="0 0 21 21"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>',
};

// ---------- toast ----------
let toastTimer;
export function toast(msg) {
  let t = $('.toast');
  if (!t) { t = document.createElement('div'); t.className = 'toast'; document.body.appendChild(t); }
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 2200);
}

// ---------- bottom sheet ----------
export function openSheet({ title, body, action, onAction, onMount }) {
  closeSheet(true);
  const bg = document.createElement('div');
  bg.className = 'sheet-bg';
  const sh = document.createElement('div');
  sh.className = 'sheet';
  sh.innerHTML = `<div class="sheet-head"><button data-close>Cancel</button><h3>${esc(title)}</h3>${action ? `<button data-act>${esc(action)}</button>` : '<span style="width:52px"></span>'}</div><div class="sheet-body">${body}</div>`;
  document.body.append(bg, sh);
  requestAnimationFrame(() => { bg.classList.add('show'); sh.classList.add('show'); });
  bg.onclick = () => closeSheet();
  sh.querySelector('[data-close]').onclick = () => closeSheet();
  if (action) sh.querySelector('[data-act]').onclick = () => onAction?.(sh);
  document.body.style.overflow = 'hidden';
  onMount?.(sh);
  return sh;
}
export function closeSheet(instant) {
  const bg = $('.sheet-bg'), sh = $('.sheet');
  document.body.style.overflow = '';
  if (!sh) return;
  if (instant) { bg?.remove(); sh.remove(); return; }
  bg?.classList.remove('show'); sh.classList.remove('show');
  setTimeout(() => { bg?.remove(); sh.remove(); }, 250);
}

export function confirmSheet(title, text, okLabel, onOk, danger = true) {
  openSheet({
    title, body: `<p style="line-height:1.5;color:var(--ink-2)">${esc(text)}</p>
      <button class="btn ${danger ? 'danger' : 'primary'}" data-ok>${esc(okLabel)}</button>`,
    onMount: (sh) => { sh.querySelector('[data-ok]').onclick = () => { closeSheet(); onOk(); }; },
  });
}

// ---------- generic forms ----------
// field: {key, label, col, type: text|textarea|number|money|usd|percent|date|select|yesno|rating, options, hint, group, placeholder}
export function formHtml(fields, getValue) {
  let html = '';
  let group = null;
  for (const f of fields) {
    if (f.group !== undefined && f.group !== group) {
      if (group !== null) html += '</div>';
      group = f.group;
      html += `<div class="form-group">${group ? `<h4>${esc(group)}</h4>` : ''}`;
    }
    if (group === null) { group = ''; html += '<div class="form-group">'; }
    html += fieldHtml(f, getValue(f));
  }
  return html + '</div>';
}

function fieldHtml(f, raw) {
  const id = 'f_' + f.key;
  const ph = f.placeholder ? ` placeholder="${esc(f.placeholder)}"` : '';
  const hint = f.hint ? `<div class="hint" data-hint="${f.key}">${esc(f.hint)}</div>` : `<div class="hint" data-hint="${f.key}"></div>`;
  switch (f.type) {
    case 'textarea':
      return `<div class="field"><label for="${id}">${esc(f.label)}</label><textarea id="${id}" name="${f.key}"${ph}>${esc(raw ?? '')}</textarea></div>`;
    case 'select': {
      const opts = ['', ...f.options].map((o) => `<option value="${esc(o)}" ${String(raw ?? '') === o ? 'selected' : ''}>${esc(o || '—')}</option>`);
      if (raw && !f.options.includes(String(raw))) opts.push(`<option selected value="${esc(raw)}">${esc(raw)}</option>`);
      return `<div class="field"><label for="${id}">${esc(f.label)}</label><select id="${id}" name="${f.key}">${opts.join('')}</select>${f.hint ? hint : ''}</div>`;
    }
    case 'yesno':
      return `<div class="field inline"><label for="${id}">${esc(f.label)}</label><label class="toggle"><input type="checkbox" id="${id}" name="${f.key}" ${raw === 'Yes' ? 'checked' : ''}><i></i></label></div>`;
    case 'rating': {
      const v = Number(raw) || 0;
      return `<div class="field"><label>${esc(f.label)}</label><div class="stars" data-stars="${f.key}">${[1, 2, 3, 4, 5].map((i) => `<button type="button" data-v="${i}" class="${i <= v ? 'on' : ''}">★</button>`).join('')}</div><input type="hidden" name="${f.key}" value="${v || ''}"></div>`;
    }
    case 'date':
      return `<div class="field"><label for="${id}">${esc(f.label)}</label><input type="date" id="${id}" name="${f.key}" value="${serialToISO(raw)}"></div>`;
    case 'percent':
      return `<div class="field"><label for="${id}">${esc(f.label)}</label><div class="money-in"><input inputmode="decimal" id="${id}" name="${f.key}" value="${raw === null || raw === '' || raw === undefined ? '' : +(Number(raw) * 100).toFixed(4)}"${ph}><span>%</span></div>${hint}</div>`;
    case 'money': case 'usd': case 'number':
      return `<div class="field"><label for="${id}">${esc(f.label)}</label><div class="money-in">${f.type === 'usd' ? '<span>$</span>' : f.type === 'money' ? '<span>EGP</span>' : ''}<input inputmode="decimal" id="${id}" name="${f.key}" value="${raw === null || raw === undefined || raw === '' ? '' : +Number(raw).toFixed(4)}"${ph}></div>${hint}</div>`;
    default:
      return `<div class="field"><label for="${id}">${esc(f.label)}</label><input type="${f.input || 'text'}" id="${id}" name="${f.key}" value="${esc(raw ?? '')}"${ph} autocomplete="off"></div>`;
  }
}

export function wireForm(root) {
  for (const box of $$('[data-stars]', root)) {
    box.onclick = (e) => {
      const b = e.target.closest('button'); if (!b) return;
      const hidden = root.querySelector(`input[name="${box.dataset.stars}"]`);
      const v = +b.dataset.v === +hidden.value ? 0 : +b.dataset.v;
      hidden.value = v || '';
      $$('button', box).forEach((x) => x.classList.toggle('on', +x.dataset.v <= v));
    };
  }
}

export function readForm(root, fields) {
  const out = {};
  for (const f of fields) {
    const el = root.querySelector(`[name="${f.key}"]`);
    if (!el) continue;
    let v;
    if (f.type === 'yesno') v = el.checked ? 'Yes' : 'No';
    else if (f.type === 'date') v = el.value ? isoToSerial(el.value) : null;
    else if (['money', 'usd', 'number', 'percent', 'rating'].includes(f.type)) {
      const s = el.value.replace(/[, ]/g, '');
      v = s === '' ? null : Number(s);
      if (v !== null && isNaN(v)) throw new Error(`"${f.label}" must be a number`);
      if (v !== null && f.type === 'percent') v = +(v / 100).toFixed(8);
    } else v = el.value.trim() === '' ? null : el.value.trim();
    out[f.key] = v;
  }
  return out;
}

export function same(a, b) {
  const n = (x) => (x === undefined || x === '' ? null : x);
  a = n(a); b = n(b);
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-9;
  return String(a ?? '') === String(b ?? '');
}

// ---------- charts ----------
export function donut(segments, size = 132, stroke = 20, center = '') {
  const r = (size - stroke) / 2, C = 2 * Math.PI * r;
  const total = segments.reduce((s, x) => s + x.value, 0) || 1;
  let off = 0;
  const arcs = segments.filter((s) => s.value > 0).map((s) => {
    const len = (s.value / total) * C;
    const a = `<circle r="${r}" cx="${size / 2}" cy="${size / 2}" fill="none" stroke="${s.color}" stroke-width="${stroke}" stroke-dasharray="${Math.max(len - 2, 0.1)} ${C}" stroke-dashoffset="${-off}" transform="rotate(-90 ${size / 2} ${size / 2})" stroke-linecap="butt"/>`;
    off += len;
    return a;
  });
  return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}"><circle r="${r}" cx="${size / 2}" cy="${size / 2}" fill="none" stroke="#fde8ed" stroke-width="${stroke}"/>${arcs.join('')}${center}</svg>`;
}

export const PALETTE = ['#d9798f', '#c9a14a', '#9b7fc2', '#5fa98a', '#e8946b', '#6f93d6', '#e6b3c1', '#a88b4f', '#7cc0b4', '#cf6f86'];
