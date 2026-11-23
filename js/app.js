import { CONFIG } from '../config.js';
import * as G from './graph.js';
import { state, subscribe, loadCache, edit, nextRow, sync, connect, clearFile, loadDemo, demoSave } from './store.js';
import { LAYOUT, layoutFor, EGP_TABS, USD_TABS } from './model.js';
import { serialToISO, isoToSerial } from './xlsx.js';
import {
  esc, $, $$, usd, egp, money, fmtDate, fmtNum, ICON, toast, openSheet, closeSheet, confirmSheet,
  formHtml, wireForm, readForm, same, donut, PALETTE,
} from './ui.js';
import { renderSeating } from './seating.js';

const app = $('#app');
const ui = { payFilter: 'All', guestFilter: 'All', guestSearch: '', taskFilter: 'Open', ownerFilter: 'All' };

const CAT_EMOJI = { Venue: '🏛️', Decor: '💐', Catering: '🍽️', Makeup: '💄', Hair: '💇‍♀️', Dress: '👰', Suit: '🤵', Church: '⛪',
  Transportation: '🚘', Photography: '📸', Videography: '🎥', Flights: '✈️', Honeymoon: '🌴', Apartment: '🏡', 'Other Quotes': '🧾' };

// ============ boot ============
async function boot() {
  const demo = new URLSearchParams(location.search).has('demo') && ['localhost', '127.0.0.1'].includes(location.hostname);
  if (demo) {
    const res = await fetch('./demo.xlsx');
    await loadDemo(new Uint8Array(await res.arrayBuffer()));
    subscribe(() => render());
    window.addEventListener('hashchange', () => render());
    window.__demoSave = demoSave;
    return render();
  }
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js').catch(() => {});
  if (!CONFIG.clientId || CONFIG.clientId.startsWith('PASTE')) return renderNotConfigured();
  try {
    await G.handleRedirect();
  } catch (e) { toast(e.message); }
  if (!G.isSignedIn() && !state.file) return renderWelcome();
  await loadCache().catch(() => false);
  if (!state.file) return renderConnect();
  subscribe(() => render());
  window.addEventListener('hashchange', () => render());
  window.addEventListener('online', () => sync());
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') sync(); });
  setInterval(() => { if (document.visibilityState === 'visible') sync(); }, 60000);
  render();
  sync();
}

function renderNotConfigured() {
  app.innerHTML = `<div class="welcome"><div class="photo" style="background-image:url(img/couple.jpg)"></div><div class="body">
  <h1>M<span class="amp">&amp;</span>E BIG DAY</h1><p>Almost there! Paste your Azure <b>Application (client) ID</b> into <code>config.js</code> (see SETUP.md), then reload.</p></div></div>`;
}

function renderWelcome() {
  app.innerHTML = `<div class="welcome"><div class="photo" style="background-image:url(img/couple.jpg)"></div><div class="body">
    <h1>M<span class="amp">&amp;</span>E BIG DAY</h1>
    <p>Your wedding budget, vendors, guests and seating — live from your Excel workbook in OneDrive.</p>
    <button class="btn primary ms-btn" id="signin">${ICON.ms}<span>Sign in with Microsoft</span></button>
    <p style="font-size:13px;margin-top:16px">Using a Gmail address? Create a free Microsoft account with it at account.microsoft.com, then sign in here.</p>
  </div></div>`;
  $('#signin').onclick = () => G.signIn('select_account');
}

function renderConnect(err) {
  app.innerHTML = `<div class="welcome"><div class="photo" style="height:34vh;background-image:url(img/couple.jpg)"></div><div class="body">
    <h1 style="font-size:30px">Connect the workbook</h1>
    ${err ? `<div class="banner bad">${esc(err)}</div>` : ''}
    <p>Which of you is this?</p>
    <button class="btn primary" id="owner">It's in my OneDrive (Ed)</button>
    <div class="form-group" style="margin-top:18px;text-align:left">
      <h4>Shared with me (Touta)</h4>
      <div class="field"><label>Share link or partner code</label><input id="shareval" placeholder="Paste the link Ed sent you" autocomplete="off"></div>
    </div>
    <button class="btn gold" id="shared">Connect shared file</button>
    <button class="btn ghost" id="sharedlist">Find it in "Shared with me"</button>
    <button class="btn ghost" id="out" style="margin-top:22px">Sign out</button>
  </div></div>`;
  const go = async (fn) => {
    try { toast('Looking for the workbook…'); const file = await fn(); await connect(file); location.hash = '#home'; boot(); }
    catch (e) { if (e instanceof G.AuthRequired) return G.signIn(); renderConnect(e.message); }
  };
  $('#owner').onclick = () => go(() => G.findByPath(CONFIG.ownerFilePath));
  $('#shared').onclick = () => {
    const v = $('#shareval').value.trim();
    if (!v) return toast('Paste the link or code first');
    go(() => G.findShared(v.startsWith('http') ? { link: v } : { code: v }));
  };
  $('#sharedlist').onclick = () => go(() => G.findShared({ name: CONFIG.fileName }));
  $('#out').onclick = () => G.signOut();
}

// ============ shell ============
function route() {
  const [a, b] = (location.hash.slice(1) || 'home').split('/').map(decodeURIComponent);
  return { tab: a || 'home', sub: b };
}

let lastHtmlKey = '';
function render() {
  const m = state.model;
  if (!m) {
    app.innerHTML = `${topbar()}<div class="empty"><div class="big">💗</div><p>${esc(state.message || 'Loading your workbook…')}</p></div>`;
    wireTop();
    return;
  }
  const { tab, sub } = route();
  const scrollY = window.scrollY;
  const key = tab + '/' + (sub || '');
  let body = '';
  if (tab === 'home') body = viewHome(m);
  else if (tab === 'money') body = viewMoney(m, sub || 'budget');
  else if (tab === 'vendors') body = sub ? viewQuotes(m, sub) : viewVendors(m);
  else if (tab === 'guests') body = viewGuests(m, sub || 'list');
  else if (tab === 'more') body = viewMore(m, sub);
  app.innerHTML = `${topbar()}${banner()}<main>${body}</main>${tabbar(tab)}`;
  wireTop();
  wireView(tab, sub, m);
  if (key === lastHtmlKey) window.scrollTo(0, scrollY); else window.scrollTo(0, 0);
  lastHtmlKey = key;
}

function topbar() {
  const s = state.status;
  const label = s === 'syncing' ? state.message : s === 'synced' ? `Synced ${timeAgo(state.lastSync)}` : s === 'offline' ? 'Offline'
    : s === 'locked' ? 'Open in Excel' : s === 'auth' ? 'Sign in needed' : s === 'error' ? 'Sync problem' : state.queue.length ? `${state.queue.length} unsaved` : 'Ready';
  return `<header class="topbar"><h1>M<span class="amp">&amp;</span>E BIG DAY</h1>
    <span class="sync-pill ${s}"><span class="dot"></span>${esc(label)}</span>
    <button class="icon-btn" data-action="sync" aria-label="Sync now">${ICON.sync}</button></header>`;
}
function banner() {
  const s = state.status;
  if (s === 'auth') return `<div class="banner bad">Your Microsoft sign-in expired, so changes can't be saved yet.<button data-action="signin">Sign in</button></div>`;
  if (s === 'locked' || s === 'error' || (s === 'offline' && state.queue.length)) return `<div class="banner ${s === 'error' ? 'bad' : 'warn'}">${esc(state.message)}</div>`;
  return '';
}
function wireTop() {
  $$('[data-action="sync"]').forEach((b) => (b.onclick = () => sync({ force: true })));
  $$('[data-action="signin"]').forEach((b) => (b.onclick = () => G.signIn()));
}
function tabbar(tab) {
  const t = (id, label, icon) => `<a href="#${id}" class="${tab === id ? 'on' : ''}">${icon}<span>${label}</span></a>`;
  return `<nav class="tabbar">${t('home', 'Home', ICON.home)}${t('money', 'Money', ICON.money)}${t('vendors', 'Vendors', ICON.vendors)}${t('guests', 'Guests', ICON.guests)}${t('more', 'More', ICON.more)}</nav>`;
}
function timeAgo(iso) {
  if (!iso) return '';
  const s = (Date.now() - new Date(iso)) / 1000;
  if (s < 60) return 'just now';
  if (s < 3600) return `${Math.floor(s / 60)}m ago`;
  if (s < 86400) return `${Math.floor(s / 3600)}h ago`;
  return fmtDate(iso.slice(0, 10), { month: 'short', day: 'numeric' });
}
const fab = (action) => `<button class="fab" data-fab="${action}" aria-label="Add">${ICON.plus}</button>`;
const statusChip = (s) => {
  const cls = { 'Over budget': 'bad', 'No allocation': 'bad', 'Near limit': 'warn', Committed: 'info', Estimated: 'gold', 'Not started': 'grey',
    'Paid in full': 'ok', Overdue: 'bad', 'Due soon': 'warn', Scheduled: 'info', 'No due date': 'grey',
    Attending: 'ok', Declined: 'bad', Pending: 'warn', Maybe: 'gold', 'Not sent': 'grey', Done: 'ok', 'In progress': 'gold', Waiting: 'info', 'To do': 'grey' }[s] || 'grey';
  return `<span class="chip ${cls}">${esc(s)}</span>`;
};

// ============ HOME ============
function viewHome(m) {
  const r = m.settings.rate, T = m.totals;
  const pct = m.settings.budget ? Math.min(100, T.forecast / m.settings.budget * 100) : 0;
  const paidPct = m.settings.budget ? Math.min(100, T.paid / m.settings.budget * 100) : 0;
  const groups = {};
  for (const b of m.budget) groups[b.group || 'Other'] = (groups[b.group || 'Other'] || 0) + b.forecast;
  const segs = Object.entries(groups).filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]).map(([k, v], i) => ({ label: k, value: v, color: PALETTE[i % PALETTE.length] }));
  const allocSegs = Object.entries(m.budget.reduce((o, b) => { o[b.group] = (o[b.group] || 0) + b.alloc; return o; }, {}))
    .filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]).map(([k, v], i) => ({ label: k, value: v, color: PALETTE[i % PALETTE.length] }));
  const showSegs = segs.length ? segs : allocSegs;
  const dueSoon = m.payments.filter((p) => p.status === 'Overdue' || p.status === 'Due soon').sort((a, b) => a.daysToDue - b.daysToDue);
  const openTasks = m.tasks.filter((t) => t.status !== 'Done').sort((a, b) => (a.dueSerial || 9e9) - (b.dueSerial || 9e9)).slice(0, 4);
  const doneCount = m.tasks.filter((t) => t.status === 'Done').length;
  const decided = Object.values(m.quotes).filter((q) => q.list.some((x) => x.isSelected)).length;
  const withQuotes = Object.values(m.quotes).filter((q) => q.list.length).length;
  const over = m.budget.filter((b) => b.status === 'Over budget' || b.status === 'No allocation');
  const rs = m.rsvp;

  return `
  <div class="hero"><img src="img/couple.jpg" alt=""><div class="txt"><div class="gold-line"></div>
    <div class="count">${m.daysToGo ?? '–'}<small>days to go</small></div>
    <div class="date">${esc(fmtDate(m.settings.weddingISO, { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' }))}</div></div></div>

  <div class="section">
    <div class="kpis">
      ${kpi('k1', 'Total budget', m.settings.budget, r)}
      ${kpi('k2', 'Forecast cost', T.forecast, r)}
      ${kpi('k3', 'Committed', T.committed, r)}
      ${kpi('k4', 'Paid to date', T.paid, r)}
      ${kpi('k5', 'Outstanding', T.outstanding, r)}
      ${kpi('k6', 'Remaining', T.remaining, r)}
    </div>
  </div>

  <div class="section" style="padding-top:0">
    <div class="card pad">
      <div class="row"><b>Budget used</b><span class="spacer"></span><span class="muted" style="font-size:13px">${pct.toFixed(0)}% forecast · ${paidPct.toFixed(0)}% paid</span></div>
      <div class="bar ${pct >= 100 ? 'over' : ''}" style="margin-top:10px;height:14px"><i style="width:${pct}%"></i><i class="paid" style="width:${paidPct}%"></i></div>
      <div class="legend"><span><i style="background:var(--rose)"></i>Forecast ${usd(T.forecast / r)}</span><span><i style="background:var(--ok)"></i>Paid ${usd(T.paid / r)}</span><span><i style="background:var(--blush-100)"></i>Budget ${usd(m.settings.budget / r)}</span></div>
      ${T.unallocated < 0 ? `<div class="chip bad" style="margin-top:10px">Allocations exceed budget by ${usd(-T.unallocated / r)}</div>` : ''}
    </div>
  </div>

  ${over.length ? `<div class="section" style="padding-top:0"><div class="card pad" style="border-color:var(--bad-bg)"><b style="color:var(--bad)">⚠️ ${over.length} categor${over.length > 1 ? 'ies' : 'y'} over budget</b>
    <div style="margin-top:6px;font-size:14px">${over.map((b) => `${esc(b.category)} <span class="muted">(${usd(-b.remaining / r)} over)</span>`).join('<br>')}</div></div></div>` : ''}

  <div class="section" style="padding-top:0">
    <div class="section-title"><h2>${segs.length ? 'Where the money goes' : 'How the budget is split'}</h2><a href="#money/budget">Budget ›</a></div>
    <div class="card pad"><div class="donut-wrap">${donut(showSegs, 132, 20, `<text x="66" y="62" text-anchor="middle" font-size="11" fill="#6e5a61">${segs.length ? 'forecast' : 'allocated'}</text><text x="66" y="80" text-anchor="middle" font-size="16" font-weight="700" fill="#3b2a30">${usd(showSegs.reduce((s, x) => s + x.value, 0) / r)}</text>`)}
      <div class="donut-list">${showSegs.slice(0, 7).map((s) => `<div><i style="background:${s.color}"></i><span>${esc(s.label)}</span><b>${usd(s.value / r)}</b></div>`).join('') || '<span class="muted">No amounts yet</span>'}</div></div></div>
  </div>

  <div class="section" style="padding-top:0">
    <div class="section-title"><h2>Payments due</h2><a href="#money/payments">All ›</a></div>
    ${dueSoon.length ? `<div class="list">${dueSoon.slice(0, 4).map((p) => payItem(p, r)).join('')}</div>`
      : `<div class="card pad muted" style="font-size:14px">Nothing due in the next ${m.settings.dueWindow} days 🎉</div>`}
  </div>

  <div class="section" style="padding-top:0">
    <div class="grid2">
      <a class="card pad" href="#guests" style="text-decoration:none;color:inherit">
        <div class="muted" style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.6px">Guests</div>
        <div style="font-family:var(--serif);font-size:28px;font-weight:700;color:var(--rose-deep);margin-top:4px">${rs.attending}<small style="font-size:14px;color:var(--ink-2)"> / ${rs.people || m.settings.guests}</small></div>
        <div class="muted" style="font-size:12.5px">attending · ${rs.waiting} awaiting reply</div></a>
      <a class="card pad" href="#vendors" style="text-decoration:none;color:inherit">
        <div class="muted" style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.6px">Vendors</div>
        <div style="font-family:var(--serif);font-size:28px;font-weight:700;color:var(--gold-deep);margin-top:4px">${decided}<small style="font-size:14px;color:var(--ink-2)"> / ${withQuotes}</small></div>
        <div class="muted" style="font-size:12.5px">categories decided</div></a>
    </div>
  </div>

  <div class="section" style="padding-top:0">
    <div class="section-title"><h2>Next up</h2><a href="#more/timeline">${doneCount}/${m.tasks.length} done ›</a></div>
    <div class="card">${openTasks.map((t) => taskItem(t)).join('') || '<div class="pad muted">All tasks done!</div>'}</div>
  </div>
  <div class="section muted" style="text-align:center;font-size:12px;padding-top:0">1 USD = ${m.settings.rate.toFixed(2)} EGP</div>`;
}
const kpi = (cls, label, amt, r) => `<div class="kpi ${cls}"><div class="lbl">${label}</div><div class="usd">${usd(amt / r)}</div><div class="egp">${egp(amt)}</div></div>`;

// ============ MONEY ============
function viewMoney(m, sub) {
  const seg = `<div class="seg"><button data-go="#money/budget" class="${sub === 'budget' ? 'on' : ''}">Budget</button><button data-go="#money/payments" class="${sub === 'payments' ? 'on' : ''}">Payments</button></div>`;
  return seg + (sub === 'payments' ? viewPayments(m) : viewBudget(m));
}

function viewBudget(m) {
  const r = m.settings.rate, T = m.totals;
  const byGroup = {};
  for (const b of m.budget) (byGroup[b.group || 'Other'] ||= []).push(b);
  return `<div class="section">
    <div class="card pad">
      <div class="kv"><span>Total budget</span><span>${usd(m.settings.budget / r)} · <small>${egp(m.settings.budget)}</small></span></div>
      <div class="kv"><span>Allocated</span><span>${usd(T.alloc / r)} · <small>${egp(T.alloc)}</small></span></div>
      <div class="kv"><span>Unallocated</span><span style="color:${T.unallocated < 0 ? 'var(--bad)' : 'inherit'}">${usd(T.unallocated / r)} · <small>${egp(T.unallocated)}</small></span></div>
      <div class="kv"><span>Forecast</span><span>${usd(T.forecast / r)} · <small>${egp(T.forecast)}</small></span></div>
    </div>
    ${Object.entries(byGroup).map(([g, list]) => `<div class="group-h">${esc(g)}</div><div class="list">${list.map((b) => {
      const pct = b.alloc > 0 ? Math.min(100, b.forecast / b.alloc * 100) : b.forecast > 0 ? 100 : 0;
      return `<button class="card bud-item" data-budget="${b._row}">
        <div class="row1"><span class="t">${esc(b.category)}</span>${statusChip(b.status)}</div>
        <div class="bar ${b.forecast > b.alloc ? 'over' : ''}"><i style="width:${pct}%"></i><i class="paid" style="width:${b.alloc > 0 ? Math.min(100, b.paid / b.alloc * 100) : 0}%"></i></div>
        <div class="nums"><span>Forecast <b>${usd(b.forecast / r)}</b></span><span>Budget <b>${usd(b.alloc / r)}</b></span><span>Left <b style="color:${b.remaining < 0 ? 'var(--bad)' : 'inherit'}">${usd(b.remaining / r)}</b></span></div>
      </button>`;
    }).join('')}</div>`).join('')}
  </div>`;
}

function openBudget(m, row) {
  const b = m.budget.find((x) => x._row === row);
  const r = m.settings.rate;
  const fields = [
    { key: 'alloc', label: 'Allocated budget (EGP)', type: 'money', col: 'E', group: 'Budget', hint: b.allocIsFormula ? `Currently a formula in Excel (${state.wb.cell('Budget', 'E' + row).f}) – saving replaces it with the number.` : '' },
    { key: 'planner', label: 'Planner estimate (EGP)', type: 'money', col: 'G', hint: 'Used when no quote is selected yet' },
    { key: 'notes', label: 'Notes', type: 'textarea', col: 'X' },
  ];
  const info = `<div class="card pad" style="margin-bottom:14px">
    <div class="kv"><span>Estimate (${esc(b.basis)})</span><span>${usd(b.estimate / r)} · <small>${egp(b.estimate)}</small></span></div>
    <div class="kv"><span>Selected quotes</span><span>${usd(b.selected / r)}</span></div>
    <div class="kv"><span>Committed</span><span>${usd(b.committed / r)}</span></div>
    <div class="kv"><span>Paid</span><span>${usd(b.paid / r)}</span></div>
    <div class="kv"><span>Outstanding</span><span>${usd(b.outstanding / r)}</span></div>
    <div class="kv"><span>Remaining</span><span style="color:${b.remaining < 0 ? 'var(--bad)' : 'inherit'}">${usd(b.remaining / r)}</span></div>
    ${b.sources.length ? `<div class="kv"><span>Quotes from</span><span>${b.sources.map((s) => `<a href="#vendors/${encodeURIComponent(s)}" data-closego>${esc(quoteTitle(m, s))}</a>`).join(', ')}</span></div>` : ''}
  </div>`;
  formSheet({ title: b.category, sheet: 'Budget', row, fields, top: info, usdHints: { alloc: true, planner: true } });
}

function payItem(p, r) {
  const sub = [p.category, p.desc].filter(Boolean).join(' · ');
  const when = p.dueISO ? (p.daysToDue < 0 ? `${-p.daysToDue}d overdue` : p.daysToDue === 0 ? 'due today' : `due ${fmtDate(p.dueISO, { month: 'short', day: 'numeric' })}`) : '';
  return `<button class="card item" data-pay="${p._row}">
    <div class="avatar" style="background:var(--gold-soft);color:var(--gold-deep)">${esc((p.vendor || p.category || '?')[0].toUpperCase())}</div>
    <div class="main"><div class="t">${esc(p.vendor || p.category || 'Payment')}</div><div class="s">${esc(sub)}</div>
      <div style="margin-top:5px;display:flex;gap:6px;align-items:center">${statusChip(p.status)}<small>${esc(when)}</small></div></div>
    <div class="amt"><b>${usd((p.balance > 0 ? p.balance : p.amount) / r)}</b><small>${p.balance > 0 ? 'balance' : 'paid'} · ${egp(p.balance > 0 ? p.balance : p.amount)}</small></div></button>`;
}

function viewPayments(m) {
  const r = m.settings.rate;
  const filters = ['All', 'Overdue', 'Due soon', 'Scheduled', 'Paid in full'];
  let list = m.payments;
  if (ui.payFilter !== 'All') list = list.filter((p) => p.status === ui.payFilter);
  list = [...list].sort((a, b) => (a.balance <= 0) - (b.balance <= 0) || (a.daysToDue ?? 9e9) - (b.daysToDue ?? 9e9));
  const tot = (k) => m.payments.reduce((s, p) => s + p[k], 0);
  return `<div class="section">
    <div class="card"><div class="stat-row"><div><b style="font-size:18px">${usd(tot('amount') / r)}</b><span>Contracted</span></div><div><b style="font-size:18px;color:var(--ok)">${usd(tot('paidNum') / r)}</b><span>Paid</span></div><div><b style="font-size:18px;color:var(--gold-deep)">${usd(tot('balance') / r)}</b><span>Balance</span></div><div><b style="font-size:18px;color:var(--bad)">${m.payments.filter((p) => p.status === 'Overdue').length}</b><span>Overdue</span></div></div></div>
    <div class="filters" style="margin-top:12px">${filters.map((f) => `<button data-payfilter="${f}" class="${ui.payFilter === f ? 'on' : ''}">${f}</button>`).join('')}</div>
    <div class="list">${list.map((p) => payItem(p, r)).join('') || `<div class="empty"><div class="big">🧾</div>No payments here yet.<br>Tap + to log a contract or payment.</div>`}</div>
  </div>${fab('payment')}`;
}

function paymentFields(m) {
  return [
    { key: 'category', label: 'Budget category', type: 'select', options: m.budget.map((b) => b.category), col: 'B', group: 'Payment' },
    { key: 'vendor', label: 'Vendor', col: 'C' },
    { key: 'desc', label: 'Description', col: 'D' },
    { key: 'amount', label: 'Contract / invoice amount (EGP)', type: 'money', col: 'E', group: 'Amounts' },
    { key: 'paid', label: 'Paid to date (EGP)', type: 'money', col: 'F' },
    { key: 'due', label: 'Next payment due', type: 'date', col: 'H' },
    { key: 'next', label: 'Next payment amount (EGP)', type: 'money', col: 'I', hint: 'Leave blank = full balance' },
    { key: 'method', label: 'Payment method', type: 'select', options: m.settings.lists.methods, col: 'J' },
    { key: 'notes', label: 'Notes', type: 'textarea', col: 'N', group: '' },
  ];
}
function openPayment(m, row) {
  const isNew = !row;
  if (isNew) { row = nextRow('Payments'); if (!row) return toast('The Payments tab is full (100 rows).'); }
  const p = m.payments.find((x) => x._row === row);
  formSheet({
    title: isNew ? 'New payment' : (p?.vendor || 'Payment'), sheet: 'Payments', row, isNew, fields: paymentFields(m), required: 'vendor',
    usdHints: { amount: true, paid: true, next: true },
    extraButtons: !isNew && p && p.balance > 0 ? [{ label: `Mark paid in full (${usd(p.balance / m.settings.rate)})`, cls: 'gold', run: () => { edit('Payments', { ['F' + row]: p.amount, ['I' + row]: 0 }); toast('Marked as paid 💸'); } }] : [],
    onDelete: isNew ? null : () => edit('Payments', 'BCDEFHIJN'.split('').map((c) => [c + row, null])),
  });
}

// ============ VENDORS ============
function quoteTitle(m, tab) { return m.quotes[tab]?.title || tab; }
function viewVendors(m) {
  const r = m.settings.rate;
  const order = ['Venue', 'Catering', 'Decor', 'Photography', 'Videography', 'Dress', 'Suit', 'Makeup', 'Hair', 'Church', 'Transportation', 'Flights', 'Honeymoon', 'Apartment',
    'Extra 1', 'Extra 2', 'Extra 3', 'Extra 4', 'Extra 5', 'Other Quotes'];
  const cards = order.filter((t) => m.quotes[t] && !m.quotes[t].hidden).map((t) => {
    const q = m.quotes[t];
    const sel = q.list.filter((x) => x.isSelected);
    const cheapest = q.list.filter((x) => x.rank === 1)[0];
    const amount = t === 'Venue' ? (x) => x.totalEGP : (x) => x.totalEGP;
    const s = sel.length ? `✓ ${esc(sel[0].vendor)}${sel.length > 1 ? ` +${sel.length - 1}` : ''}` : q.list.length ? `${q.list.length} quote${q.list.length > 1 ? 's' : ''}${cheapest && amount(cheapest) > 0 ? ` · from ${usd(amount(cheapest) / r)}` : ''}` : 'No quotes yet';
    return `<button class="card cat-card" data-go="#vendors/${encodeURIComponent(t)}" style="${sel.length ? 'border-color:var(--gold)' : ''}">
      <span class="emoji">${CAT_EMOJI[t] || '✨'}</span><span class="t">${esc(q.title)}</span><span class="s">${s}</span>
      ${sel.length ? '<span class="chip gold">Decided</span>' : q.list.length ? '<span class="chip warn">Comparing</span>' : '<span class="chip grey">Open</span>'}</button>`;
  });
  return `<div class="section"><div class="section-title"><h2>Vendors & quotes</h2></div><div class="cat-grid">${cards.join('')}</div></div>`;
}

function viewQuotes(m, tab) {
  const q = m.quotes[tab];
  if (!q) return `<div class="empty">Unknown tab</div>`;
  const r = m.settings.rate;
  const isVenue = tab === 'Venue';
  const val = (x) => (isVenue ? x.likeForLike : x.totalEGP);
  const list = [...q.list].sort((a, b) => (b.isSelected - a.isSelected) || ((a.rank || 999) - (b.rank || 999)));
  const max = Math.max(1, ...list.map(val));
  let groups = [['', list]];
  if (tab === 'Other Quotes') {
    const g = {};
    for (const x of list) (g[x.category || 'Uncategorised'] ||= []).push(x);
    groups = Object.entries(g);
  }
  return `<div class="section">
    <div class="row" style="margin-bottom:8px"><a href="#vendors" style="color:var(--rose-deep);text-decoration:none;font-weight:600">‹ Vendors</a></div>
    <div class="section-title"><h2>${CAT_EMOJI[tab] || '✨'} ${esc(q.title)}</h2><span class="muted" style="font-size:13px">${q.list.length} quote${q.list.length === 1 ? '' : 's'}</span></div>
    ${isVenue ? `<div class="calc">Ranked by <b>like-for-like</b> cost: package price plus the planner's benchmark for anything the package doesn't include (catering, décor, DJ, lighting, cake, coordinator).</div>` : ''}
    ${groups.map(([g, l]) => `${g ? `<div class="group-h">${esc(g)}</div>` : ''}<div class="list">${l.map((x) => `
      <button class="card item quote ${x.isSelected ? 'sel' : ''}" data-quote="${x._row}">
        <span class="rank ${x.rank === 1 ? 'r1' : ''}">${x.rank ? '#' + x.rank : '–'}</span>
        <div class="main"><div class="t">${esc(x.vendor)}${x.option ? ` <span class="muted" style="font-weight:400">· ${esc(x.option)}</span>` : ''}</div>
          <div class="s">${esc([isVenue ? x.type : x.d1, isVenue ? x.location : x.d2].filter(Boolean).join(' · '))}</div>
          <div class="compare-bar"><i style="width:${val(x) / max * 100}%"></i></div>
          <div style="margin-top:6px;display:flex;gap:6px;flex-wrap:wrap">${x.isSelected ? '<span class="chip gold">★ Selected</span>' : ''}${x.rating ? `<span class="chip grey">${'★'.repeat(x.rating)}</span>` : ''}${isVenue ? `<span class="chip grey">${usd(x.perGuest / r)}/guest</span>` : ''}</div></div>
        <div class="amt"><b>${usd(val(x) / r)}</b><small>${egp(val(x))}</small>${isVenue && x.addOn ? `<small style="display:block">pkg ${usd(x.totalEGP / r)}</small>` : ''}</div>
      </button>`).join('')}</div>`).join('') || `<div class="empty"><div class="big">${CAT_EMOJI[tab] || '✨'}</div>No quotes yet — tap + to add the first one.</div>`}
  </div>${fab('quote:' + tab)}`;
}

function quoteFields(m, tab) {
  const L = layoutFor(tab);
  const hdr = (col, fb) => String(state.wb.value(tab, col + '5') || fb).replace(/\s*\(.*?\)\s*$/, (s) => (s.length > 18 ? '' : s));
  const isUsd = USD_TABS.includes(tab);
  const f = [{ key: 'selected', label: '★ Selected (counts toward the budget)', type: 'yesno', col: L.selected, group: '' }];
  if (tab === 'Other Quotes') f.push({ key: 'category', label: 'Budget category', type: 'select', options: m.budget.map((b) => b.category), col: L.category, group: 'Vendor' });
  f.push({ key: 'vendor', label: 'Vendor / provider', col: L.vendor, group: 'Vendor' },
    { key: 'option', label: 'Option / package', col: L.option },
    { key: 'd1', label: hdr(L.d1, 'Detail 1'), col: L.d1 },
    { key: 'd2', label: hdr(L.d2, 'Detail 2'), col: L.d2 },
    { key: 'd3', label: hdr(L.d3, 'Detail 3'), col: L.d3 });
  if (!isUsd) f.push({ key: 'currency', label: 'Quote currency', type: 'select', options: ['EGP', 'USD'], col: L.currency, group: 'Price', hint: 'Blank = EGP' });
  const cur = isUsd ? 'usd' : 'number';
  f.push({ key: 'price', label: `Unit price${isUsd ? ' (USD)' : ' (quote currency)'}`, type: cur, col: L.price, group: isUsd ? 'Price' : undefined },
    { key: 'qty', label: 'Quantity', type: 'number', col: L.qty, hint: tab === 'Catering' ? `Blank = guest count (${m.settings.guests})` : 'Blank = 1' },
    { key: 'extras', label: 'Extras / add-ons', type: cur, col: L.extras },
    { key: 'discount', label: 'Discount', type: cur, col: L.discount },
    { key: 'tax', label: 'Tax / service', type: 'percent', col: L.tax },
    { key: 'deposit', label: `Deposit (${isUsd ? 'USD' : 'EGP'})`, type: isUsd ? 'usd' : 'money', col: L.deposit, group: 'Terms' },
    { key: 'terms', label: 'Payment / cancellation terms', type: 'textarea', col: L.terms },
    { key: 'included', label: 'Included services', type: 'textarea', col: L.included },
    { key: 'exclusions', label: 'Exclusions', type: 'textarea', col: L.exclusions },
    { key: 'rating', label: 'Rating', type: 'rating', col: L.rating },
    { key: 'contact', label: 'Contact', col: L.contact },
    { key: 'notes', label: 'Notes', type: 'textarea', col: L.notes });
  return f;
}

function venueFields(m) {
  const L = LAYOUT.venue;
  const yes = (key, label) => ({ key, label, type: 'yesno', col: L[key] });
  return [
    { key: 'selected', label: '★ Selected (counts toward the budget)', type: 'yesno', col: L.selected, group: '' },
    { key: 'vendor', label: 'Venue name', col: L.vendor, group: 'Package' },
    { key: 'option', label: 'Package name', col: L.option },
    { key: 'type', label: 'Package type', type: 'select', options: m.settings.lists.packageTypes, col: L.type },
    { key: 'location', label: 'Location / area', col: L.location },
    { key: 'maxCap', label: 'Max capacity', type: 'number', col: L.maxCap },
    { key: 'guestCount', label: 'Guest count', type: 'number', col: L.guestCount, hint: `Blank = Settings (${m.settings.guests})` },
    { key: 'rental', label: 'Venue rental', type: 'money', col: L.rental, group: 'Pricing (EGP)' },
    { key: 'catPerGuest', label: 'Catering per guest', type: 'money', col: L.catPerGuest },
    { key: 'catFlat', label: 'Catering flat fee', type: 'money', col: L.catFlat },
    { key: 'decoration', label: 'Decoration', type: 'money', col: L.decoration },
    { key: 'otherSvc', label: 'Other included services', type: 'money', col: L.otherSvc },
    { key: 'allIn', label: 'All-in price (if not itemised)', type: 'money', col: L.allIn },
    { key: 'service', label: 'Service charge', type: 'percent', col: L.service },
    { key: 'vat', label: 'VAT / tax', type: 'percent', col: L.vat },
    { key: 'discount', label: 'Discount', type: 'money', col: L.discount },
    { ...yes('catIncl', 'Catering included'), group: "What's included" },
    { key: 'decoLevel', label: 'Decoration level', type: 'select', options: m.settings.lists.decoLevels, col: L.decoLevel },
    yes('sound', 'Sound & DJ'), yes('lighting', 'Lighting'), yes('cake', 'Wedding cake'), yes('coordinator', 'Event coordinator'), yes('suite', 'Bridal suite'),
    { key: 'deposit', label: 'Deposit (EGP)', type: 'money', col: L.deposit, group: 'Terms & decision' },
    { key: 'terms', label: 'Payment terms', type: 'textarea', col: L.terms },
    { key: 'validUntil', label: 'Quote valid until', type: 'date', col: L.validUntil },
    { key: 'exclusions', label: 'Exclusions / not included', type: 'textarea', col: L.exclusions },
    { key: 'contact', label: 'Contact', col: L.contact },
    { key: 'notes', label: 'Notes', type: 'textarea', col: L.notes },
  ];
}

function openQuote(m, tab, row) {
  const isNew = !row;
  if (isNew) { row = nextRow(tab); if (!row) return toast(`The ${tab} tab is full.`); }
  const q = m.quotes[tab].list.find((x) => x._row === row);
  const isVenue = tab === 'Venue';
  const fields = isVenue ? venueFields(m) : quoteFields(m, tab);
  const r = m.settings.rate;
  const top = !isNew && q ? `<div class="calc">${isVenue
    ? `Package total <b>${usd(q.totalEGP / r)}</b> (${egp(q.totalEGP)}) · like-for-like <b>${usd(q.likeForLike / r)}</b> · ${usd(q.perGuest / r)} per guest · rank #${q.rank ?? '–'}`
    : `Total <b>${usd(q.totalEGP / r)}</b> · ${egp(q.totalEGP)} · price rank #${q.rank ?? '–'}`}</div>` : '';
  const L = layoutFor(tab);
  const clearCols = fields.map((f) => f.col);
  formSheet({
    title: isNew ? `New ${quoteTitle(m, tab)} quote` : q?.vendor || 'Quote', sheet: tab, row, isNew, fields, top, required: 'vendor',
    liveTotal: isVenue ? null : (vals) => {
      const isUsd = USD_TABS.includes(tab);
      const qty = vals.qty === null ? (tab === 'Catering' ? m.settings.guests : 1) : vals.qty;
      const t = ((vals.price || 0) * qty + (vals.extras || 0) - (vals.discount || 0)) * (1 + (vals.tax || 0));
      const egpT = isUsd || vals.currency === 'USD' ? t * r : t;
      return `Total <b>${usd(egpT / r)}</b> · ${egp(egpT)}`;
    },
    onDelete: isNew ? null : () => edit(tab, clearCols.map((c) => [c + row, null])),
    keyCol: L.key,
  });
}

// ============ GUESTS ============
function viewGuests(m, sub) {
  const seg = `<div class="seg"><button data-go="#guests" class="${sub === 'list' ? 'on' : ''}">Guest list</button><button data-go="#guests/seating" class="${sub === 'seating' ? 'on' : ''}">Seating</button></div>`;
  if (!m.hasGuests) return seg + `<div class="empty">This workbook has no Guests tab yet.</div>`;
  if (sub === 'seating') return seg + `<div id="seating-root"></div>`;
  const rs = m.rsvp;
  const filters = ['All', 'Attending', 'Pending', 'Not sent', 'Maybe', 'Declined', 'Bride', 'Groom', 'Unseated'];
  const q = ui.guestSearch.toLowerCase();
  let list = m.guests.filter((g) => !q || [g.name, g.group, g.phone, g.hotel, g.from, g.table].some((x) => String(x || '').toLowerCase().includes(q)));
  const f = ui.guestFilter;
  if (['Attending', 'Pending', 'Not sent', 'Maybe', 'Declined'].includes(f)) list = list.filter((g) => (g.rsvp || 'Not sent') === f);
  else if (f === 'Bride' || f === 'Groom') list = list.filter((g) => g.side === f || g.side === 'Both');
  else if (f === 'Unseated') list = list.filter((g) => !g.table && g.rsvp !== 'Declined');
  list.sort((a, b) => String(a.name).localeCompare(String(b.name)));
  return seg + `<div class="section">
    <div class="card"><div class="stat-row"><div><b>${rs.invited}</b><span>Invites</span></div><div><b>${rs.people}</b><span>People</span></div><div><b style="color:var(--ok)">${rs.attending}</b><span>Attending</span></div><div><b style="color:var(--gold-deep)">${rs.waiting}</b><span>Waiting</span></div></div></div>
    ${rs.people && rs.people !== m.settings.guests ? `<div class="calc" style="margin-top:12px">Budget assumes <b>${m.settings.guests}</b> guests; your list has <b>${rs.people}</b> people (${rs.attending} confirmed). <a href="#" data-syncguests style="color:var(--rose-deep);font-weight:700">Update the budget guest count</a></div>` : ''}
    <div class="search" style="margin-top:12px">${ICON.search}<input id="gsearch" type="search" placeholder="Search guests, hotels, tables…" value="${esc(ui.guestSearch)}"></div>
    <div class="filters">${filters.map((x) => `<button data-gfilter="${x}" class="${f === x ? 'on' : ''}">${x}</button>`).join('')}</div>
    <div class="list">${list.map((g) => `<button class="card item" data-guest="${g._row}">
      <div class="avatar" style="background:${g.side === 'Groom' ? '#e3ecfb' : g.side === 'Bride' ? 'var(--blush-100)' : 'var(--gold-soft)'};color:${g.side === 'Groom' ? '#3c5aa0' : g.side === 'Bride' ? 'var(--rose-deep)' : 'var(--gold-deep)'}">${esc(initials(g.name))}</div>
      <div class="main"><div class="t">${esc(g.name)}${g.plus ? ` <span class="muted" style="font-weight:400">+${g.plus}</span>` : ''}</div>
        <div class="s">${esc([g.side, g.group, g.table ? '🪑 ' + g.table : '', g.hotel ? '🏨 ' + g.hotel : ''].filter(Boolean).join(' · '))}</div></div>
      ${statusChip(g.rsvp || 'Not sent')}</button>`).join('') || `<div class="empty"><div class="big">💌</div>${m.guests.length ? 'No guests match.' : 'Your guest list is empty.<br>Tap + to add your first guest.'}</div>`}</div>
  </div>${fab('guest')}`;
}
const initials = (n) => String(n || '?').split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0].toUpperCase()).join('');

function guestFields(m) {
  const L = LAYOUT.guests;
  return [
    { key: 'name', label: 'Guest name', col: L.name, group: 'Guest', placeholder: 'e.g. Mina & Mariam Youssef' },
    { key: 'side', label: 'Side', type: 'select', options: ['Bride', 'Groom', 'Both'], col: L.side },
    { key: 'group', label: 'Group', type: 'select', options: ['Family', 'Friends', 'Work', 'Church', 'Other'], col: L.group },
    { key: 'rsvp', label: 'RSVP', type: 'select', options: ['Not sent', 'Pending', 'Attending', 'Declined', 'Maybe'], col: L.rsvp, group: 'RSVP' },
    { key: 'plus', label: 'Plus-ones', type: 'number', col: L.plus, hint: 'Extra people on this invitation' },
    { key: 'table', label: 'Table', type: 'select', options: m.tables.map((t) => t.name), col: L.table },
    { key: 'seat', label: 'Seat #', type: 'number', col: L.seat },
    { key: 'phone', label: 'Phone', col: L.phone, input: 'tel', group: 'Contact & travel' },
    { key: 'email', label: 'Email', col: L.email, input: 'email' },
    { key: 'from', label: 'Coming from', col: L.from },
    { key: 'arrival', label: 'Arrival date', type: 'date', col: L.arrival },
    { key: 'flight', label: 'Flight / transport', col: L.flight },
    { key: 'hotel', label: 'Hotel', col: L.hotel },
    { key: 'notes', label: 'Notes', type: 'textarea', col: L.notes, group: '' },
  ];
}
export function openGuest(m, row, preset = {}) {
  const isNew = !row;
  if (isNew) { row = nextRow('Guests'); if (!row) return toast('The Guests tab is full (500 rows).'); }
  const fields = guestFields(m);
  formSheet({
    title: isNew ? 'New guest' : state.wb.value('Guests', 'B' + row), sheet: 'Guests', row, isNew, fields, required: 'name', preset,
    onDelete: isNew ? null : () => edit('Guests', fields.map((f) => [f.col + row, null])),
  });
}

// ============ MORE ============
function viewMore(m, sub) {
  if (sub === 'timeline') return viewTimeline(m);
  if (sub === 'trip') return viewTrip(m);
  if (sub === 'settings') return viewSettings(m);
  const tile = (href, emoji, t, s) => `<a class="card item" href="${href}" style="text-decoration:none;color:inherit"><div class="avatar" style="background:var(--blush-100);font-size:20px">${emoji}</div><div class="main"><div class="t">${t}</div><div class="s">${s}</div></div><span class="chev">${ICON.chev}</span></a>`;
  const done = m.tasks.filter((t) => t.status === 'Done').length;
  return `<div class="section"><div class="list">
    ${tile('#more/timeline', '🗓️', 'Timeline & tasks', `${done} of ${m.tasks.length} done · ${m.tasks.filter((t) => t.overdue).length} overdue`)}
    ${tile('#more/trip', '🧳', 'Wedding trip', `${m.trip.length} days around the big day`)}
    ${tile('#more/settings', '⚙️', 'Settings', `Budget, exchange rate, wedding date, sharing`)}
  </div></div>`;
}

function taskItem(t) {
  const cls = t.status === 'Done' ? 'done' : t.status === 'In progress' ? 'prog' : '';
  return `<div class="task ${t.status === 'Done' ? 'done' : ''}" data-task="${t._row}">
    <button class="check ${cls}" data-taskcheck="${t._row}" aria-label="Toggle done">${t.status === 'Done' ? ICON.check : ''}</button>
    <div class="main"><div class="t">${esc(t.task)}</div><div class="s">${t.dueISO ? `<span style="color:${t.overdue ? 'var(--bad)' : 'inherit'}">${esc(fmtDate(t.dueISO, { month: 'short', day: 'numeric', year: '2-digit' }))}</span>` : ''}${t.owner ? `<span class="chip grey">${esc(t.owner)}</span>` : ''}${t.status !== 'Done' && t.status !== 'To do' && t.status ? statusChip(t.status) : ''}</div></div></div>`;
}

function viewTimeline(m) {
  const owners = ['All', 'Ed', 'Touta', 'Both', 'Planner', 'Family'];
  let list = m.tasks;
  if (ui.taskFilter === 'Open') list = list.filter((t) => t.status !== 'Done');
  else if (ui.taskFilter === 'Overdue') list = list.filter((t) => t.overdue);
  else if (ui.taskFilter === 'Done') list = list.filter((t) => t.status === 'Done');
  if (ui.ownerFilter !== 'All') list = list.filter((t) => t.owner === ui.ownerFilter || (t.owner === 'Both' && ['Ed', 'Touta'].includes(ui.ownerFilter)));
  list = [...list].sort((a, b) => (a.dueSerial || 9e9) - (b.dueSerial || 9e9));
  const phases = {};
  for (const t of list) (phases[t.phase || 'Other'] ||= []).push(t);
  const done = m.tasks.filter((t) => t.status === 'Done').length;
  return `<div class="section">
    <div class="row" style="margin-bottom:8px"><a href="#more" style="color:var(--rose-deep);text-decoration:none;font-weight:600">‹ More</a></div>
    <div class="card pad"><div class="row"><b>${done} of ${m.tasks.length} tasks done</b><span class="spacer"></span><span class="muted">${m.tasks.length ? Math.round(done / m.tasks.length * 100) : 0}%</span></div>
      <div class="bar" style="margin-top:10px"><i class="paid" style="width:${m.tasks.length ? done / m.tasks.length * 100 : 0}%"></i></div></div>
    <div class="filters" style="margin-top:12px">${['Open', 'Overdue', 'Done', 'All'].map((f) => `<button data-tfilter="${f}" class="${ui.taskFilter === f ? 'on' : ''}">${f}</button>`).join('')}<span style="width:8px"></span>${owners.map((o) => `<button data-ofilter="${o}" class="${ui.ownerFilter === o ? 'on' : ''}">${o === 'All' ? 'Everyone' : o}</button>`).join('')}</div>
    ${Object.entries(phases).sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true })).map(([p, l]) => `<div class="group-h">${esc(p)}</div><div class="card">${l.map(taskItem).join('')}</div>`).join('') || '<div class="empty">Nothing here 🎉</div>'}
  </div>${fab('task')}`;
}

function openTask(m, row) {
  const isNew = !row;
  if (isNew) { row = nextRow('Timeline'); if (!row) return toast('The Timeline tab is full.'); }
  const phases = [...new Set(m.tasks.map((t) => t.phase).filter(Boolean))];
  const L = LAYOUT.timeline;
  const fields = [
    { key: 'task', label: 'Task', col: L.task, group: '' },
    { key: 'owner', label: 'Owner', type: 'select', options: ['Ed', 'Touta', 'Both', 'Planner', 'Family'], col: L.owner },
    { key: 'status', label: 'Status', type: 'select', options: ['To do', 'In progress', 'Waiting', 'Done'], col: L.status },
    { key: 'phase', label: 'Phase', type: 'select', options: phases, col: L.phase },
    { key: 'dueDate', label: 'Due date', type: 'date', virtual: true },
    { key: 'notes', label: 'Notes', type: 'textarea', col: L.notes },
  ];
  const wedding = m.settings.weddingSerial;
  formSheet({
    title: isNew ? 'New task' : 'Task', sheet: 'Timeline', row, isNew, fields, required: 'task',
    preset: isNew ? { status: 'To do' } : {},
    virtualGet: { dueDate: () => { const d = state.wb.value('Timeline', 'G' + row); return d === null || d === '' ? null : wedding - Number(d); } },
    virtualSet: (vals, changes) => { if (vals.dueDate !== undefined) { const cur = state.wb.value('Timeline', 'G' + row); const nv = vals.dueDate === null ? null : wedding - vals.dueDate; if (!same(cur, nv)) changes.push(['G' + row, nv]); } },
    onDelete: isNew ? null : () => edit('Timeline', 'CDEFGH'.split('').map((c) => [c + row, null])),
  });
}

function viewTrip(m) {
  return `<div class="section"><div class="row" style="margin-bottom:8px"><a href="#more" style="color:var(--rose-deep);text-decoration:none;font-weight:600">‹ More</a></div>
    <div class="section-title"><h2>Wedding trip</h2></div>
    <div class="card">${m.trip.map((d) => `<div class="task" style="${d.offset === 0 ? 'background:var(--gold-soft)' : ''}">
      <div class="avatar" style="background:${d.offset === 0 ? 'var(--gold)' : 'var(--blush-100)'};color:${d.offset === 0 ? '#fff' : 'var(--rose-deep)'};flex-direction:column;font-size:11px;line-height:1.1;width:46px;height:46px">${esc(fmtDate(d.dateISO, { month: 'short' }))}<b style="font-size:17px">${esc(fmtDate(d.dateISO, { day: 'numeric' }))}</b></div>
      <div class="main"><div class="t">${esc(d.plan || '—')}</div><div class="s">${esc(fmtDate(d.dateISO, { weekday: 'long' }))} · ${esc(d.where)} · ${d.offset === 0 ? '💍 Wedding day' : d.offset < 0 ? `${-d.offset} days before` : `${d.offset} days after`}</div></div></div>`).join('')}</div>
    <p class="muted" style="font-size:12.5px;text-align:center">Dates follow the wedding date in Settings. Edit plans on the Trip tab in Excel.</p></div>`;
}

function viewSettings(m) {
  const s = m.settings;
  const who = JSON.parse(localStorage.getItem('meb.me') || 'null');
  const partnerCode = state.file ? `${state.file.driveId}|${state.file.itemId}` : '';
  return `<div class="section">
    <div class="row" style="margin-bottom:8px"><a href="#more" style="color:var(--rose-deep);text-decoration:none;font-weight:600">‹ More</a></div>
    <div class="section-title"><h2>Exchange rate</h2></div>
    <div class="card pad">
      <div class="kv"><span>Rate in workbook</span><span>1 USD = ${s.rate.toFixed(2)} EGP</span></div>
      ${s.rateIsFormula ? `<div class="muted" style="font-size:12.5px;margin:6px 0">The workbook pulls this from an Excel currency data type, which currently shows an error. Saving a rate here replaces it with a fixed number.</div>` : ''}
      <div id="fxbox" style="margin-top:10px"></div>
      <button class="btn gold" data-fx>Get today's live rate</button>
      <button class="btn ghost" data-editrate>Enter a rate manually</button>
    </div>
    <div class="section-title" style="margin-top:18px"><h2>Wedding</h2><button data-editsettings>Edit</button></div>
    <div class="card pad">
      <div class="kv"><span>Wedding date</span><span>${esc(fmtDate(s.weddingISO))}</span></div>
      <div class="kv"><span>Total budget</span><span>${usd(s.budget / s.rate)} · <small>${egp(s.budget)}</small></span></div>
      <div class="kv"><span>Guest count (for budget)</span><span>${s.guests}</span></div>
      <div class="kv"><span>"Due soon" window</span><span>${s.dueWindow} days</span></div>
    </div>
    <div class="section-title" style="margin-top:18px"><h2>Sync & sharing</h2></div>
    <div class="card pad">
      <div class="kv"><span>Signed in as</span><span>${esc(who?.name || '—')}</span></div>
      <div class="kv"><span>Workbook</span><span>${esc(state.file?.name || '')}</span></div>
      <div class="kv"><span>Last synced</span><span>${esc(state.lastSync ? new Date(state.lastSync).toLocaleString() : 'never')}</span></div>
      <div class="kv"><span>Waiting to save</span><span>${state.queue.length}</span></div>
      <div style="margin-top:12px;font-size:13px" class="muted">Partner code — Touta can paste this on the "Connect" screen (after you share the file with her in OneDrive):</div>
      <div class="code-box" style="margin-top:6px">${esc(partnerCode)}</div>
      <button class="btn ghost" data-copycode style="margin-top:10px">Copy partner code</button>
      ${state.meta?.webUrl ? `<a class="btn ghost" style="margin-top:10px;text-decoration:none" href="${esc(state.meta.webUrl)}" target="_blank" rel="noopener">Open in Excel</a>` : ''}
      <button class="btn ghost" data-forcesync style="margin-top:10px">Sync now</button>
      <button class="btn danger" data-disconnect style="margin-top:10px">Disconnect workbook</button>
      <button class="btn danger" data-signout style="margin-top:10px">Sign out</button>
    </div>
    <p class="muted" style="text-align:center;font-size:12px;margin-top:18px">M&amp;E BIG DAY · made with 💗 for Sept 25, 2027</p>
  </div>`;
}

function openSettingsForm(m) {
  const fields = [
    { key: 'date', label: 'Wedding date', type: 'date', col: 'C10', group: 'Wedding' },
    { key: 'budget', label: 'Total wedding budget (EGP)', type: 'money', col: 'C7' },
    { key: 'guests', label: 'Expected guest count', type: 'number', col: 'C9' },
    { key: 'window', label: '"Due soon" alert window (days)', type: 'number', col: 'C11' },
    { key: 'b1', label: 'Catering per guest', type: 'money', col: 'C14', group: 'Venue benchmarks (EGP)' },
    { key: 'b2', label: 'Full decoration', type: 'money', col: 'C15' },
    { key: 'b3', label: 'Partial → full decoration top-up', type: 'money', col: 'C16' },
    { key: 'b4', label: 'Sound & DJ', type: 'money', col: 'C17' },
    { key: 'b5', label: 'Lighting', type: 'money', col: 'C18' },
    { key: 'b6', label: 'Wedding cake', type: 'money', col: 'C19' },
    { key: 'b7', label: 'Event coordinator', type: 'money', col: 'C20' },
  ];
  formSheet({ title: 'Settings', sheet: 'Settings', row: '', fields, cellMode: true, usdHints: { budget: true } });
}

function openRateForm(m, preset) {
  const fields = [{ key: 'rate', label: 'EGP per 1 USD', type: 'number', col: 'C6', group: '' }];
  formSheet({ title: 'Exchange rate', sheet: 'Settings', row: '', fields, cellMode: true, preset: preset ? { rate: preset } : {} });
}

// ============ generic form sheet ============
function formSheet({ title, sheet, row, isNew, fields, top = '', required, onDelete, extraButtons = [], preset = {}, usdHints = {}, liveTotal, cellMode, virtualGet = {}, virtualSet }) {
  const ref = (f) => (cellMode ? f.col : f.col + row);
  const getValue = (f) => (f.key in preset ? preset[f.key] : f.virtual ? virtualGet[f.key]?.() : state.wb.value(sheet, ref(f)));
  const body = `${top}<div data-live class="calc ${liveTotal ? '' : 'hidden'}"></div>${formHtml(fields, getValue)}
    ${extraButtons.map((b, i) => `<button class="btn ${b.cls || 'ghost'}" data-extra="${i}">${esc(b.label)}</button>`).join('')}
    ${onDelete ? `<button class="btn danger" data-del style="margin-top:12px">Delete</button>` : ''}`;
  openSheet({
    title, body, action: 'Save',
    onMount: (sh) => {
      wireForm(sh);
      const rate = state.model.settings.rate;
      const refresh = () => {
        let vals;
        try { vals = readForm(sh, fields); } catch { return; }
        for (const f of fields) {
          const h = sh.querySelector(`[data-hint="${f.key}"]`);
          if (!h) continue;
          if (usdHints[f.key] && vals[f.key]) h.textContent = `≈ ${usd(vals[f.key] / rate)}`;
          else if (f.type === 'number' && f.key === 'price' && vals.currency !== 'USD' && vals.price) h.textContent = `≈ ${usd(vals.price / rate)}`;
          else h.textContent = f.hint || '';
        }
        if (liveTotal) sh.querySelector('[data-live]').innerHTML = liveTotal(vals);
      };
      sh.addEventListener('input', refresh);
      sh.addEventListener('change', refresh);
      refresh();
      sh.querySelectorAll('[data-extra]').forEach((b) => (b.onclick = () => { closeSheet(); extraButtons[+b.dataset.extra].run(); }));
      const del = sh.querySelector('[data-del]');
      if (del) del.onclick = () => confirmSheet('Delete?', 'This clears the row in the workbook. You can undo from the backup or Excel version history.', 'Delete', () => { onDelete(); toast('Deleted'); });
    },
    onAction: (sh) => {
      let vals;
      try { vals = readForm(sh, fields); } catch (e) { return toast(e.message); }
      if (required && !vals[required]) return toast('Please fill in ' + fields.find((f) => f.key === required).label.toLowerCase());
      const changes = [];
      for (const f of fields) {
        if (f.virtual) continue;
        const cur = state.wb.value(sheet, ref(f));
        let nv = vals[f.key];
        if (f.type === 'yesno' && isNew && nv === 'No' && (cur === null || cur === '')) nv = 'No';
        if (!same(cur, nv) || (f.key in preset && isNew)) changes.push([ref(f), nv]);
      }
      virtualSet?.(vals, changes);
      closeSheet();
      if (!changes.length) return;
      // the key column first so a new row is "claimed" in one batch
      edit(sheet, changes, { newRow: !!isNew });
      toast(navigator.onLine ? 'Saved ✓ syncing…' : 'Saved on this phone ✓ will sync when online');
    },
  });
}

// ============ wiring ============
function wireView(tab, sub, m) {
  $$('[data-go]').forEach((b) => (b.onclick = () => (location.hash = b.dataset.go)));
  $$('[data-budget]').forEach((b) => (b.onclick = () => openBudget(m, +b.dataset.budget)));
  $$('[data-pay]').forEach((b) => (b.onclick = () => openPayment(m, +b.dataset.pay)));
  $$('[data-payfilter]').forEach((b) => (b.onclick = () => { ui.payFilter = b.dataset.payfilter; render(); }));
  $$('[data-quote]').forEach((b) => (b.onclick = () => openQuote(m, sub, +b.dataset.quote)));
  $$('[data-guest]').forEach((b) => (b.onclick = () => openGuest(m, +b.dataset.guest)));
  $$('[data-gfilter]').forEach((b) => (b.onclick = () => { ui.guestFilter = b.dataset.gfilter; render(); }));
  $$('[data-tfilter]').forEach((b) => (b.onclick = () => { ui.taskFilter = b.dataset.tfilter; render(); }));
  $$('[data-ofilter]').forEach((b) => (b.onclick = () => { ui.ownerFilter = b.dataset.ofilter; render(); }));
  $$('[data-task]').forEach((el) => (el.onclick = (e) => { if (!e.target.closest('[data-taskcheck]')) openTask(m, +el.dataset.task); }));
  $$('[data-taskcheck]').forEach((b) => (b.onclick = () => {
    const row = +b.dataset.taskcheck;
    const cur = state.wb.value('Timeline', 'E' + row);
    edit('Timeline', { ['E' + row]: cur === 'Done' ? 'To do' : 'Done' });
    if (cur !== 'Done') toast('Nice! ✓');
  }));
  const gs = $('#gsearch');
  if (gs) {
    gs.oninput = () => { ui.guestSearch = gs.value; const pos = gs.selectionStart; render(); const n = $('#gsearch'); n.focus(); n.setSelectionRange(pos, pos); };
  }
  $$('[data-syncguests]').forEach((a) => (a.onclick = (e) => {
    e.preventDefault();
    const n = m.rsvp.attending || m.rsvp.people;
    confirmSheet('Update guest count?', `Set the budget guest count (Settings) to ${n}? This changes catering and venue estimates.`, `Use ${n} guests`, () => edit('Settings', { C9: n }), false);
  }));
  const f = $('[data-fab]');
  if (f) f.onclick = () => {
    const a = f.dataset.fab;
    if (a === 'payment') openPayment(m);
    else if (a === 'guest') openGuest(m);
    else if (a === 'task') openTask(m);
    else if (a.startsWith('quote:')) openQuote(m, a.slice(6));
  };
  // settings
  $$('[data-editsettings]').forEach((b) => (b.onclick = () => openSettingsForm(m)));
  $$('[data-editrate]').forEach((b) => (b.onclick = () => openRateForm(m)));
  $$('[data-fx]').forEach((b) => (b.onclick = async () => {
    const box = $('#fxbox');
    box.innerHTML = '<div class="muted">Fetching…</div>';
    try {
      const j = await (await fetch(CONFIG.fxUrl, { cache: 'no-store' })).json();
      const rate = j.rates?.EGP;
      if (!rate) throw new Error('No EGP rate in the response');
      box.innerHTML = `<div class="calc">Live rate: <b>1 USD = ${rate.toFixed(2)} EGP</b><br><small>${esc(j.time_last_update_utc || '')} · open.er-api.com</small></div>
        <button class="btn primary" data-savefx>Save ${rate.toFixed(2)} to the workbook</button><div style="height:10px"></div>`;
      $('[data-savefx]').onclick = () => { edit('Settings', { C6: +rate.toFixed(4) }); toast('Exchange rate updated'); };
    } catch (e) { box.innerHTML = `<div class="banner bad" style="margin:0 0 10px">Couldn't get a live rate (${esc(e.message)}).</div>`; }
  }));
  $$('[data-copycode]').forEach((b) => (b.onclick = async () => {
    try { await navigator.clipboard.writeText(`${state.file.driveId}|${state.file.itemId}`); toast('Copied'); } catch { toast('Select the code and copy it'); }
  }));
  $$('[data-forcesync]').forEach((b) => (b.onclick = () => sync({ force: true })));
  $$('[data-disconnect]').forEach((b) => (b.onclick = () => confirmSheet('Disconnect?', state.queue.length ? `${state.queue.length} change(s) haven't been saved yet and will be lost.` : 'The workbook stays in OneDrive; this only unlinks the app.', 'Disconnect', () => { clearFile(); localStorage.removeItem('meb.queue'); location.reload(); })));
  $$('[data-signout]').forEach((b) => (b.onclick = () => G.signOut()));
  if (tab === 'guests' && sub === 'seating' && m.hasGuests) renderSeating($('#seating-root'), m, { openGuest: (row, preset) => openGuest(state.model, row, preset) });
}

// remember who is signed in (for Settings)
if (G.isSignedIn() && !localStorage.getItem('meb.me')) G.me().then((u) => localStorage.setItem('meb.me', JSON.stringify({ name: u.displayName, mail: u.userPrincipalName }))).catch(() => {});

boot();
