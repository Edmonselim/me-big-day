// Local cache (IndexedDB), offline edit queue and two-way sync with the OneDrive workbook.
import { Workbook, splitRef } from './xlsx.js';
import { compute, layoutFor, firstEmptyRow } from './model.js';
import * as G from './graph.js';

const DB = 'meb', STORE = 'kv';
function idb() {
  return new Promise((res, rej) => {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(STORE);
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
async function kv(mode, fn) {
  const db = await idb();
  return new Promise((res, rej) => {
    const tx = db.transaction(STORE, mode);
    const req = fn(tx.objectStore(STORE));
    tx.oncomplete = () => res(req?.result);
    tx.onerror = () => rej(tx.error);
  });
}
const kvGet = (k) => kv('readonly', (s) => s.get(k));
const kvSet = (k, v) => kv('readwrite', (s) => s.put(v, k));

const QUEUE = 'meb.queue';
const FILE = 'meb.file';

export const state = {
  wb: null, model: null, meta: null, file: JSON.parse(localStorage.getItem(FILE) || 'null'),
  queue: JSON.parse(localStorage.getItem(QUEUE) || '[]'),
  status: 'idle', message: '', lastSync: localStorage.getItem('meb.lastSync') || null,
  listeners: new Set(),
};
const emit = () => state.listeners.forEach((f) => f(state));
export const subscribe = (f) => { state.listeners.add(f); return () => state.listeners.delete(f); };
const saveQueue = () => localStorage.setItem(QUEUE, JSON.stringify(state.queue));

export function setFile(file) { state.file = file; localStorage.setItem(FILE, JSON.stringify(file)); }
export function clearFile() { state.file = null; localStorage.removeItem(FILE); }

function setStatus(status, message = '') { state.status = status; state.message = message; emit(); }

function recompute() { state.model = compute(state.wb); emit(); }

// Load whatever we have offline, with pending edits layered on top.
export async function loadCache() {
  const bytes = await kvGet('workbook');
  if (!bytes) return false;
  state.wb = await Workbook.open(bytes);
  state.meta = await kvGet('meta');
  for (const op of state.queue) state.wb.setLocal(op.sheet, op.ref, op.value);
  recompute();
  return true;
}

// ---- editing API used by the UI ----
let tmpCounter = 0;
export function edit(sheet, changes, { newRow = false } = {}) {
  // changes: { ref: value } or [[ref, value]]
  const entries = Array.isArray(changes) ? changes : Object.entries(changes);
  let tmp = newRow ? `${sheet}#${Date.now()}#${tmpCounter++}` : null;
  if (!tmp && entries.length) {
    // editing a record that was created offline and not yet saved: keep it in the same group
    const row = splitRef(entries[0][0]).row;
    const pending = state.queue.find((o) => o.tmp && o.sheet === sheet && splitRef(o.ref).row === row);
    if (pending) tmp = pending.tmp;
  }
  for (const [ref, value] of entries) {
    state.wb.setLocal(sheet, ref, value);
    // collapse repeated edits of the same cell
    state.queue = state.queue.filter((o) => !(o.sheet === sheet && o.ref === ref && o.tmp === tmp));
    state.queue.push({ sheet, ref, value: value === undefined ? null : value, tmp, ts: Date.now() });
  }
  saveQueue();
  recompute();
  scheduleSync();
}

export function nextRow(sheet) {
  const claimed = new Set();
  return firstEmptyRow(state.wb, sheet, claimed);
}

export function clearRow(sheet, row, cols) {
  edit(sheet, cols.map((c) => [c + row, null]));
}

// ---- sync ----
let timer = null;
export function scheduleSync(delay = 1500) {
  clearTimeout(timer);
  timer = setTimeout(() => sync().catch(() => {}), delay);
}

let running = null;
export function sync(opts = {}) {
  if (running) return running;
  running = doSync(opts).finally(() => { running = null; });
  return running;
}

async function doSync({ force = false } = {}) {
  if (!state.file) return;
  if (!navigator.onLine) { setStatus('offline', state.queue.length ? `${state.queue.length} change(s) waiting – will sync when you're online` : 'Offline – showing your last copy'); return; }
  try {
    if (!state.queue.length) {
      setStatus('syncing', 'Checking for updates…');
      const meta = await G.itemMeta(state.file);
      if (!force && state.meta && meta.eTag === state.meta.eTag && state.wb) { done(meta); return; }
      setStatus('syncing', 'Downloading latest…');
      const { bytes } = await G.download(state.file);
      await adopt(bytes, meta);
      done(meta);
      return;
    }
    for (let attempt = 0; attempt < 3; attempt++) {
      setStatus('syncing', `Saving ${state.queue.length} change(s)…`);
      const { bytes, meta } = await G.download(state.file);
      const wb = await Workbook.open(bytes);
      const ops = state.queue.slice();
      const edits = resolveRows(wb, ops);
      await wb.applyEdits(edits);
      const out = await wb.save();
      try {
        const newMeta = await G.upload(state.file, out, meta.eTag);
        // remove exactly the ops we saved (new ones may have arrived meanwhile)
        state.queue = state.queue.filter((o) => !ops.includes(o));
        saveQueue();
        await adopt(out, newMeta);
        for (const op of state.queue) state.wb.setLocal(op.sheet, op.ref, op.value);
        recompute();
        if (!state.queue.length) { done(newMeta); return; }
      } catch (e) {
        if (e.status === 412) continue; // someone saved in between – merge again
        throw e;
      }
    }
    setStatus('error', 'The file kept changing while saving – will try again shortly.');
    scheduleSync(10000);
  } catch (e) {
    if (e instanceof G.AuthRequired) { setStatus('auth', 'Please sign in again to sync.'); return; }
    if (e.status === 423 || e.code === 'resourceLocked' || e.code === 'notAllowed') {
      setStatus('locked', 'The workbook is open in Excel – close it there and the app will save automatically.');
      scheduleSync(20000);
      return;
    }
    if (e.status === 404) { setStatus('error', "Can't find the workbook in OneDrive – check it wasn't moved or renamed."); return; }
    if (!navigator.onLine || e instanceof TypeError) { setStatus('offline', `${state.queue.length} change(s) waiting – will sync when you're online`); return; }
    setStatus('error', e.message || String(e));
    scheduleSync(30000);
  }
}

function done(meta) {
  state.meta = meta;
  kvSet('meta', meta);
  state.lastSync = new Date().toISOString();
  localStorage.setItem('meb.lastSync', state.lastSync);
  setStatus('synced', '');
}

async function adopt(bytes, meta) {
  state.wb = await Workbook.open(bytes);
  state.meta = meta;
  await kvSet('workbook', bytes);
  await kvSet('meta', meta);
  for (const op of state.queue) state.wb.setLocal(op.sheet, op.ref, op.value);
  recompute();
}

// New rows were picked on this phone; if the other phone already used that row,
// move our new record to the next free row so nothing gets overwritten.
function resolveRows(wb, ops) {
  const claimed = {};
  const remap = {};
  for (const op of ops) {
    if (!op.tmp || remap[op.tmp] !== undefined) continue;
    const { row } = splitRef(op.ref);
    const L = layoutFor(op.sheet);
    claimed[op.sheet] ||= new Set();
    const taken = !!wb.value(op.sheet, L.key + row) || claimed[op.sheet].has(row);
    const target = taken ? firstEmptyRow(wb, op.sheet, claimed[op.sheet]) : row;
    if (target === null) throw new Error(`No empty rows left on the ${op.sheet} tab.`);
    claimed[op.sheet].add(target);
    remap[op.tmp] = target;
  }
  return ops.map((op) => {
    if (!op.tmp) return op;
    const { col } = splitRef(op.ref);
    return { sheet: op.sheet, ref: col + remap[op.tmp], value: op.value };
  });
}

// Local testing only (localhost ?demo): load a workbook without OneDrive.
export async function loadDemo(bytes) {
  state.wb = await Workbook.open(bytes);
  state.file = null;
  state.status = 'synced';
  state.lastSync = new Date().toISOString();
  recompute();
}
export async function demoSave() {
  const ops = state.queue.slice();
  const wb = await Workbook.open(await state.wb.save());
  await wb.applyEdits(resolveRows(wb, ops));
  return wb.save();
}

export async function connect(file) {
  setFile(file);
  state.meta = null;
  await sync({ force: true });
}

export async function downloadCopy() {
  const bytes = await kvGet('workbook');
  return bytes;
}
