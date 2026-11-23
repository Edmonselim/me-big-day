// Turns the workbook into plain objects and recomputes everything the Excel
// formulas compute, so the app shows correct totals right after an edit
// (Excel itself recalculates when the file is next opened).
import { serialToISO } from './xlsx.js';

export const EGP_TABS = ['Decor', 'Catering', 'Makeup', 'Hair', 'Dress', 'Suit', 'Church', 'Transportation',
  'Photography', 'Videography', 'Extra 1', 'Extra 2', 'Extra 3', 'Extra 4', 'Extra 5'];
export const USD_TABS = ['Flights', 'Honeymoon', 'Apartment'];

// Column layouts. key -> column letter
export const LAYOUT = {
  egp: { first: 6, last: 35, key: 'B', vendor: 'B', option: 'C', d1: 'D', d2: 'E', d3: 'F', currency: 'G', price: 'H', qty: 'I',
    extras: 'J', discount: 'K', tax: 'L', included: 'P', exclusions: 'Q', deposit: 'R', terms: 'S', rating: 'T', contact: 'U', notes: 'V', selected: 'W' },
  usd: { first: 6, last: 35, key: 'B', vendor: 'B', option: 'C', d1: 'D', d2: 'E', d3: 'F', price: 'G', qty: 'H',
    extras: 'I', discount: 'J', tax: 'K', included: 'O', exclusions: 'P', deposit: 'Q', terms: 'R', rating: 'S', contact: 'T', notes: 'U', selected: 'V' },
  other: { first: 6, last: 35, key: 'C', category: 'B', vendor: 'C', option: 'D', d1: 'E', d2: 'F', d3: 'G', currency: 'H', price: 'I', qty: 'J',
    extras: 'K', discount: 'L', tax: 'M', included: 'Q', exclusions: 'R', deposit: 'S', terms: 'T', rating: 'U', contact: 'V', notes: 'W', selected: 'X' },
  venue: { first: 6, last: 53, key: 'B', vendor: 'B', option: 'C', type: 'D', location: 'E', maxCap: 'F', guestCount: 'G', rental: 'I',
    catPerGuest: 'J', catFlat: 'K', decoration: 'M', otherSvc: 'N', allIn: 'O', service: 'P', vat: 'Q', discount: 'R',
    catIncl: 'V', decoLevel: 'W', sound: 'X', lighting: 'Y', cake: 'Z', coordinator: 'AA', suite: 'AB',
    exclusions: 'AK', terms: 'AL', deposit: 'AM', validUntil: 'AN', contact: 'AO', notes: 'AP', selected: 'AQ' },
  payments: { first: 6, last: 105, key: 'C', category: 'B', vendor: 'C', desc: 'D', amount: 'E', paid: 'F', due: 'H', next: 'I', method: 'J', notes: 'N' },
  timeline: { first: 7, last: 110, key: 'C', task: 'C', owner: 'D', status: 'E', phase: 'F', daysBefore: 'G', notes: 'H' },
  guests: { first: 6, last: 505, key: 'B', name: 'B', side: 'C', group: 'D', rsvp: 'E', plus: 'F', phone: 'H', email: 'I',
    from: 'J', arrival: 'K', flight: 'L', hotel: 'M', table: 'N', seat: 'O', notes: 'P' },
  tables: { first: 6, last: 55, key: 'B', name: 'B', shape: 'C', seats: 'D', x: 'G', y: 'H', rot: 'I', notes: 'J' },
  budget: { first: 6, last: 45, key: 'B', group: 'A', category: 'B', alloc: 'E', planner: 'G', notes: 'X' },
};

export function layoutFor(sheet) {
  if (EGP_TABS.includes(sheet)) return LAYOUT.egp;
  if (USD_TABS.includes(sheet)) return LAYOUT.usd;
  return { 'Other Quotes': LAYOUT.other, Venue: LAYOUT.venue, Payments: LAYOUT.payments, Timeline: LAYOUT.timeline,
    Guests: LAYOUT.guests, Tables: LAYOUT.tables, Budget: LAYOUT.budget }[sheet];
}

const num = (v) => (v === null || v === undefined || v === '' || isNaN(Number(v)) ? 0 : Number(v));
const str = (v) => (v === null || v === undefined ? '' : String(v));
const blank = (v) => v === null || v === undefined || v === '';

export function readRows(wb, sheet) {
  const L = layoutFor(sheet);
  const rows = [];
  for (let r = L.first; r <= L.last; r++) {
    const o = { _row: r, _sheet: sheet };
    for (const [k, c] of Object.entries(L)) {
      if (k === 'first' || k === 'last' || k === 'key') continue;
      o[k] = wb.value(sheet, c + r);
    }
    o._empty = blank(wb.value(sheet, L.key + r));
    rows.push(o);
  }
  return rows;
}

export function firstEmptyRow(wb, sheet, exclude = new Set()) {
  const L = layoutFor(sheet);
  for (let r = L.first; r <= L.last; r++) {
    if (exclude.has(r)) continue;
    if (blank(wb.value(sheet, L.key + r))) return r;
  }
  return null;
}

// Rank like COUNTIFS(range,"<"&x,range,">0")+1 (only positive totals ranked)
function rank(list, get) {
  for (const q of list) {
    const t = get(q);
    q.rank = t > 0 ? list.filter((o) => get(o) > 0 && get(o) < t).length + 1 : null;
  }
}

export function compute(wb, now = new Date()) {
  const S = (ref) => wb.value('Settings', ref);
  const settings = {
    rate: num(S('C6')) || 55,
    rateIsFormula: !!wb.cell('Settings', 'C6')?.f,
    budget: num(S('C7')),
    guests: num(S('C9')),
    weddingSerial: num(S('C10')),
    weddingISO: serialToISO(S('C10')),
    dueWindow: num(S('C11')) || 30,
    bench: { catering: num(S('C14')), fullDeco: num(S('C15')), topUp: num(S('C16')), sound: num(S('C17')),
      lighting: num(S('C18')), cake: num(S('C19')), coordinator: num(S('C20')) },
    lists: {
      groups: listCol(wb, 'Settings', 'B', 24, 33), packageTypes: listCol(wb, 'Settings', 'D', 24, 33),
      decoLevels: listCol(wb, 'Settings', 'H', 24, 30), methods: listCol(wb, 'Settings', 'L', 24, 33),
    },
  };
  const rate = settings.rate;
  const todaySerial = Math.floor((Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())) / 86400000) + 25569;

  // ---- quote tabs ----
  const quotes = {};
  for (const tab of EGP_TABS) {
    if (!wb.sheets[tab]) continue;
    const list = readRows(wb, tab).filter((q) => !q._empty);
    for (const q of list) {
      const qty = blank(q.qty) ? (tab === 'Catering' ? settings.guests : 1) : num(q.qty);
      q.totalEGP = ((num(q.price) * qty + num(q.extras) - num(q.discount)) * (1 + num(q.tax))) * (q.currency === 'USD' ? rate : 1);
      q.isSelected = q.selected === 'Yes';
    }
    rank(list, (q) => q.totalEGP);
    const title = tab.startsWith('Extra') ? str(wb.value(tab, 'C4')) || tab : tab;
    quotes[tab] = { tab, title, currency: 'EGP', list, hidden: wb.hidden.has(tab) && !wb.value(tab, 'C4') && !list.length };
  }
  for (const tab of USD_TABS) {
    if (!wb.sheets[tab]) continue;
    const list = readRows(wb, tab).filter((q) => !q._empty);
    for (const q of list) {
      q.totalUSD = (num(q.price) * (blank(q.qty) ? 1 : num(q.qty)) + num(q.extras) - num(q.discount)) * (1 + num(q.tax));
      q.totalEGP = q.totalUSD * rate;
      q.isSelected = q.selected === 'Yes';
    }
    rank(list, (q) => q.totalEGP);
    quotes[tab] = { tab, title: tab, currency: 'USD', list };
  }
  // Other Quotes
  {
    const list = readRows(wb, 'Other Quotes').filter((q) => !q._empty);
    for (const q of list) {
      q.totalEGP = ((num(q.price) * (blank(q.qty) ? 1 : num(q.qty)) + num(q.extras) - num(q.discount)) * (1 + num(q.tax))) * (q.currency === 'USD' ? rate : 1);
      q.isSelected = q.selected === 'Yes';
    }
    for (const q of list) {
      const t = q.totalEGP;
      q.rank = t > 0 ? list.filter((o) => o.category === q.category && o.totalEGP > 0 && o.totalEGP < t).length + 1 : null;
    }
    quotes['Other Quotes'] = { tab: 'Other Quotes', title: 'Other Quotes', currency: 'EGP', list };
  }
  // Venue
  {
    const b = settings.bench;
    const list = readRows(wb, 'Venue').filter((q) => !q._empty);
    for (const q of list) {
      const guests = blank(q.guestCount) ? settings.guests : num(q.guestCount);
      const cat = num(q.catPerGuest) * guests + num(q.catFlat);
      const parts = num(q.rental) + cat + num(q.decoration) + num(q.otherSvc);
      const total = (num(q.allIn) > 0 ? num(q.allIn) : parts) * (1 + num(q.service)) * (1 + num(q.vat)) - num(q.discount);
      const addOn = (q.catIncl === 'Yes' ? 0 : b.catering * guests)
        + (q.decoLevel === 'Full' ? 0 : q.decoLevel === 'Partial' ? b.topUp : b.fullDeco)
        + (q.sound === 'Yes' ? 0 : b.sound) + (q.lighting === 'Yes' ? 0 : b.lighting)
        + (q.cake === 'Yes' ? 0 : b.cake) + (q.coordinator === 'Yes' ? 0 : b.coordinator);
      Object.assign(q, {
        guestsUsed: guests, cateringTotal: cat, totalEGP: total, addOn, likeForLike: total + addOn,
        perGuest: guests > 0 ? (total + addOn) / guests : 0,
        splitVenue: parts > 0 ? (num(q.rental) + num(q.otherSvc)) * total / parts : total,
        splitCatering: parts > 0 ? cat * total / parts : 0,
        splitDeco: parts > 0 ? num(q.decoration) * total / parts : 0,
        isSelected: q.selected === 'Yes',
      });
    }
    rank(list, (q) => q.likeForLike);
    quotes.Venue = { tab: 'Venue', title: 'Venue', currency: 'EGP', list };
  }

  // column value used by Budget's SUMIFS references
  const colValue = (tab, col, q) => {
    if (tab === 'Venue') return { AH: q.splitVenue, AI: q.splitCatering, AJ: q.splitDeco, S: q.totalEGP, AD: q.likeForLike }[col] ?? 0;
    return q.totalEGP;
  };

  // ---- payments ----
  const payments = readRows(wb, 'Payments').filter((p) => !(blank(p.category) && blank(p.vendor) && num(p.amount) === 0));
  for (const p of payments) {
    p.amount = num(p.amount); p.paidNum = num(p.paid);
    p.balance = p.amount - p.paidNum;
    p.dueISO = serialToISO(p.due);
    p.status = p.balance <= 0 ? 'Paid in full' : blank(p.due) ? 'No due date'
      : num(p.due) < todaySerial ? 'Overdue' : num(p.due) - todaySerial <= settings.dueWindow ? 'Due soon' : 'Scheduled';
    p.daysToDue = blank(p.due) ? null : num(p.due) - todaySerial;
  }

  // ---- budget ----
  const budget = [];
  for (let r = 6; r <= 45; r++) {
    const name = wb.value('Budget', 'B' + r);
    if (blank(name)) continue;
    const f = wb.cell('Budget', 'H' + r)?.f || '';
    let selected = 0;
    const sources = [];
    for (const m of f.matchAll(/SUMIFS\('?([^'!]+)'?!\$([A-Z]+)\$\d+:\$[A-Z]+\$\d+,'?([^'!]+)'?!\$([A-Z]+)\$\d+:\$[A-Z]+\$\d+,"Yes"\)/g)) {
      const tab = m[1];
      if (!quotes[tab]) continue;
      sources.push(tab);
      for (const q of quotes[tab].list) if (q.isSelected) selected += colValue(tab, m[2], q);
    }
    if (f.includes("'Other Quotes'!$B")) {
      for (const q of quotes['Other Quotes'].list) if (q.isSelected && q.category === name) selected += q.totalEGP;
    }
    const alloc = num(wb.value('Budget', 'E' + r));
    const planner = num(wb.value('Budget', 'G' + r));
    const committed = payments.filter((p) => p.category === name).reduce((s, p) => s + p.amount, 0);
    const paid = payments.filter((p) => p.category === name).reduce((s, p) => s + p.paidNum, 0);
    const estimate = selected > 0 ? selected : planner > 0 ? planner : committed;
    const basis = selected > 0 ? 'Selected quote' : planner > 0 ? 'Planner estimate' : committed > 0 ? 'Actual (no quote)' : 'Pending';
    const forecast = Math.max(estimate, committed);
    const status = alloc === 0 && forecast > 0 ? 'No allocation' : forecast > alloc ? 'Over budget'
      : (alloc > 0 && forecast >= 0.9 * alloc) ? 'Near limit' : committed > 0 ? 'Committed' : estimate > 0 ? 'Estimated' : 'Not started';
    budget.push({ _row: r, group: str(wb.value('Budget', 'A' + r)), category: name, alloc, planner, selected, estimate, basis,
      committed, paid, outstanding: committed - paid, forecast, remaining: alloc - forecast, status, sources,
      notes: str(wb.value('Budget', 'X' + r)), allocIsFormula: !!wb.cell('Budget', 'E' + r)?.f, nameIsFormula: !!wb.cell('Budget', 'B' + r)?.f });
  }
  const sum = (k) => budget.reduce((s, b) => s + b[k], 0);
  const totals = { alloc: sum('alloc'), estimate: sum('estimate'), committed: sum('committed'), paid: sum('paid'),
    outstanding: sum('outstanding'), forecast: sum('forecast') };
  totals.remaining = settings.budget - totals.forecast;
  totals.unallocated = settings.budget - totals.alloc;

  // ---- timeline ----
  const tasks = readRows(wb, 'Timeline').filter((t) => !t._empty).map((t) => ({
    ...t, dueSerial: blank(t.daysBefore) ? null : settings.weddingSerial - num(t.daysBefore),
  }));
  for (const t of tasks) { t.dueISO = t.dueSerial ? serialToISO(t.dueSerial) : ''; t.overdue = t.status !== 'Done' && t.dueSerial && t.dueSerial < todaySerial; }

  // ---- guests & tables ----
  const hasGuests = !!wb.sheets.Guests;
  const guests = hasGuests ? readRows(wb, 'Guests').filter((g) => !g._empty) : [];
  for (const g of guests) { g.party = g.rsvp === 'Declined' ? 0 : 1 + num(g.plus); g.arrivalISO = serialToISO(g.arrival); }
  const tables = wb.sheets.Tables ? readRows(wb, 'Tables').filter((t) => !t._empty) : [];
  for (const t of tables) {
    t.seatsNum = num(t.seats);
    t.guests = guests.filter((g) => g.table === t.name);
    t.assigned = t.guests.reduce((s, g) => s + g.party, 0);
    t.left = t.seatsNum - t.assigned;
  }
  const rsvp = {
    invited: guests.length,
    people: guests.reduce((s, g) => s + g.party, 0),
    attending: guests.filter((g) => g.rsvp === 'Attending').reduce((s, g) => s + g.party, 0),
    declined: guests.filter((g) => g.rsvp === 'Declined').length,
    waiting: guests.filter((g) => !g.rsvp || g.rsvp === 'Pending' || g.rsvp === 'Not sent' || g.rsvp === 'Maybe').length,
    seated: guests.filter((g) => g.table).reduce((s, g) => s + g.party, 0),
  };

  // ---- trip ----
  const trip = [];
  for (let r = 8; r <= 28; r++) {
    const d = wb.value('Trip', 'A' + r);
    if (blank(d)) continue;
    trip.push({ _row: r, offset: num(d), dateISO: serialToISO(settings.weddingSerial + num(d)), where: str(wb.value('Trip', 'D' + r)), plan: str(wb.value('Trip', 'E' + r)) });
  }

  const daysToGo = settings.weddingSerial ? settings.weddingSerial - todaySerial : null;
  return { settings, quotes, payments, budget, totals, tasks, guests, tables, rsvp, trip, daysToGo, hasGuests, todaySerial };
}

function listCol(wb, sheet, col, from, to) {
  const out = [];
  for (let r = from; r <= to; r++) { const v = wb.value(sheet, col + r); if (!blank(v)) out.push(String(v)); }
  return out;
}
