// Visual floor plan: drag tables around, tap a table to seat guests.
import { state, edit, nextRow } from './store.js';
import { esc, openSheet, closeSheet, confirmSheet, toast, formHtml, wireForm, readForm, same } from './ui.js';

const W = 1000, H = 1400;
let zoom = Number(localStorage.getItem('meb.zoom') || 1.75);

function geometry(t) {
  const seats = Math.max(1, t.seatsNum || 10);
  if (t.shape === 'Rectangle' || t.shape === 'Long' || t.shape === 'Head') {
    const w = Math.max(120, Math.ceil(seats / 2) * 34 + 30);
    return { kind: 'rect', w, h: t.shape === 'Head' ? 60 : 76 };
  }
  return { kind: 'round', r: 34 + seats * 2.6 };
}

function seatDots(t, g) {
  const n = Math.max(1, t.seatsNum || 0);
  const out = [];
  for (let i = 0; i < n; i++) {
    const filled = i < t.assigned;
    const fill = filled ? (t.left < 0 ? '#cc4b5d' : '#c9a14a') : '#fff';
    let x, y;
    if (g.kind === 'round') {
      const a = (i / n) * Math.PI * 2 - Math.PI / 2;
      x = Math.cos(a) * (g.r + 15); y = Math.sin(a) * (g.r + 15);
    } else if (t.shape === 'Head') {
      const per = n; x = -g.w / 2 + (i + 0.5) * (g.w / per); y = g.h / 2 + 15;
    } else {
      const top = Math.ceil(n / 2), row = i < top ? 0 : 1, idx = row ? i - top : i, cnt = row ? n - top : top;
      x = -g.w / 2 + (idx + 0.5) * (g.w / cnt); y = row ? g.h / 2 + 15 : -g.h / 2 - 15;
    }
    out.push(`<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="9" fill="${fill}" stroke="#d9798f" stroke-width="2"/>`);
  }
  return out.join('');
}

function tableSvg(t) {
  const g = geometry(t);
  const x = Number(t.x) || W / 2, y = Number(t.y) || H / 2;
  const full = t.left <= 0 && t.assigned > 0;
  const fill = t.left < 0 ? '#fbe3e6' : full ? '#f3e3bd' : '#fff';
  const shape = g.kind === 'round'
    ? `<circle class="shape" r="${g.r}" fill="${fill}" stroke="#d9798f" stroke-width="2.5"/>`
    : `<rect class="shape" x="${-g.w / 2}" y="${-g.h / 2}" width="${g.w}" height="${g.h}" rx="12" fill="${fill}" stroke="#d9798f" stroke-width="2.5"/>`;
  return `<g class="tbl" data-row="${t._row}" transform="translate(${x} ${y})">
    ${seatDots(t, g)}${shape}
    <text text-anchor="middle" y="-2" font-family="Iowan Old Style, Georgia, serif" font-size="22" font-weight="700" fill="#3b2a30">${esc(t.name)}</text>
    <text text-anchor="middle" y="22" font-size="17" fill="${t.left < 0 ? '#cc4b5d' : '#6e5a61'}">${t.assigned}/${t.seatsNum}</text></g>`;
}

export function renderSeating(root, m, { openGuest }) {
  const attending = m.guests.filter((g) => g.rsvp !== 'Declined');
  const unseated = attending.filter((g) => !g.table);
  const seats = m.tables.reduce((s, t) => s + t.seatsNum, 0);
  root.innerHTML = `
    <div class="floor-tools">
      <div class="info"><b>${m.rsvp.seated}</b> seated · <b>${unseated.reduce((s, g) => s + g.party, 0)}</b> to seat · <b>${seats}</b> seats at ${m.tables.length} tables</div>
      <button class="icon-btn" data-z="-1" aria-label="Zoom out">−</button>
      <button class="icon-btn" data-z="1" aria-label="Zoom in">+</button>
    </div>
    <div class="floor-wrap"><svg id="floor" viewBox="0 0 ${W} ${H}" width="100%">
      <rect x="300" y="40" width="400" height="110" rx="18" fill="#fde8ed" stroke="#f4b6c2" stroke-dasharray="10 8" stroke-width="3"/>
      <text x="500" y="105" text-anchor="middle" font-size="26" fill="#b4566e" font-family="Iowan Old Style, Georgia, serif">💍 Kosha / Stage</text>
      <rect x="340" y="175" width="320" height="110" rx="18" fill="#fffaf0" stroke="#e9cf8d" stroke-dasharray="10 8" stroke-width="3"/>
      <text x="500" y="240" text-anchor="middle" font-size="24" fill="#9a7426" font-family="Iowan Old Style, Georgia, serif">Dance floor</text>
      ${m.tables.map(tableSvg).join('')}
    </svg></div>
    <div class="section">
      <button class="btn primary" data-addtable>+ Add a table</button>
      <p class="muted" style="font-size:12.5px;text-align:center;margin-top:10px">Drag a table to move it · tap a table to seat guests or change its size and shape.</p>
      ${unseated.length ? `<div class="group-h">Still to seat (${unseated.length} invitations)</div><div class="card">${unseated.map((g) =>
        `<div class="seat-guest"><div class="n"><b>${esc(g.name)}</b>${g.plus ? ` +${g.plus}` : ''}<div class="muted" style="font-size:12.5px">${esc([g.side, g.group, g.rsvp || 'Not sent'].filter(Boolean).join(' · '))}</div></div><button data-pick="${g._row}">Seat…</button></div>`).join('')}</div>` : ''}
    </div>`;

  const svg = root.querySelector('#floor');
  const wrap = root.querySelector('.floor-wrap');
  const applyZoom = () => {
    const cx = wrap.scrollLeft + wrap.clientWidth / 2, cw = wrap.scrollWidth || 1;
    svg.style.width = `${zoom * 100}%`; svg.removeAttribute('width');
    wrap.scrollLeft = (cx / cw) * wrap.scrollWidth - wrap.clientWidth / 2;
  };
  const saved = JSON.parse(sessionStorage.getItem('meb.floorScroll') || 'null');
  svg.style.width = `${zoom * 100}%`; svg.removeAttribute('width');
  requestAnimationFrame(() => {
    if (saved) { wrap.scrollLeft = saved[0]; wrap.scrollTop = saved[1]; }
    else wrap.scrollLeft = (wrap.scrollWidth - wrap.clientWidth) / 2;
  });
  wrap.addEventListener('scroll', () => sessionStorage.setItem('meb.floorScroll', JSON.stringify([wrap.scrollLeft, wrap.scrollTop])), { passive: true });
  root.querySelectorAll('[data-z]').forEach((b) => (b.onclick = () => {
    zoom = Math.min(3, Math.max(0.6, +(zoom + 0.25 * b.dataset.z).toFixed(2)));
    localStorage.setItem('meb.zoom', zoom);
    applyZoom();
  }));

  // ---- drag ----
  const toSvg = (e) => {
    const p = svg.createSVGPoint(); p.x = e.clientX; p.y = e.clientY;
    return p.matrixTransform(svg.getScreenCTM().inverse());
  };
  svg.querySelectorAll('.tbl').forEach((el) => {
    let start = null;
    el.addEventListener('pointerdown', (e) => {
      const row = +el.dataset.row;
      const t = m.tables.find((x) => x._row === row);
      const p = toSvg(e);
      start = { p, x: Number(t.x) || W / 2, y: Number(t.y) || H / 2, moved: false, t };
      el.setPointerCapture(e.pointerId);
      el.classList.add('drag');
      e.preventDefault();
    });
    el.addEventListener('pointermove', (e) => {
      if (!start) return;
      const p = toSvg(e);
      const dx = p.x - start.p.x, dy = p.y - start.p.y;
      if (Math.hypot(dx, dy) > 6) start.moved = true;
      if (start.moved) {
        start.nx = Math.round(Math.min(W - 40, Math.max(40, start.x + dx)));
        start.ny = Math.round(Math.min(H - 40, Math.max(40, start.y + dy)));
        el.setAttribute('transform', `translate(${start.nx} ${start.ny})`);
      }
    });
    const end = () => {
      if (!start) return;
      el.classList.remove('drag');
      const s = start; start = null;
      if (s.moved && s.nx !== undefined) edit('Tables', { ['G' + s.t._row]: s.nx, ['H' + s.t._row]: s.ny });
      else openTable(s.t._row);
    };
    el.addEventListener('pointerup', end);
    el.addEventListener('pointercancel', () => { start = null; el.classList.remove('drag'); });
  });

  root.querySelector('[data-addtable]').onclick = () => {
    const row = nextRow('Tables');
    if (!row) return toast('The Tables tab is full (50 tables).');
    const names = new Set(m.tables.map((t) => t.name));
    let n = m.tables.length + 1;
    while (names.has(`Table ${n}`)) n++;
    // drop it in the visible middle of the floor
    const rect = wrap.getBoundingClientRect();
    const mid = toSvg({ clientX: rect.left + rect.width / 2, clientY: rect.top + rect.height / 2 });
    edit('Tables', { ['B' + row]: `Table ${n}`, ['C' + row]: 'Round', ['D' + row]: 10, ['G' + row]: Math.round(mid.x), ['H' + row]: Math.round(Math.max(380, mid.y)), ['I' + row]: 0 }, { newRow: true });
    toast(`Table ${n} added`);
  };
  root.querySelectorAll('[data-pick]').forEach((b) => (b.onclick = () => pickTable(+b.dataset.pick)));

  function pickTable(guestRow) {
    const g = m.guests.find((x) => x._row === guestRow);
    openSheet({
      title: `Seat ${g.name}`,
      body: `<div class="card">${m.tables.map((t) => `<div class="seat-guest"><div class="n"><b>${esc(t.name)}</b><div class="muted" style="font-size:12.5px">${t.assigned}/${t.seatsNum} seated · ${t.left >= g.party ? `${t.left} free` : t.left > 0 ? `only ${t.left} free` : 'full'}</div></div><button data-t="${esc(t.name)}">${t.left >= g.party ? 'Seat here' : 'Seat anyway'}</button></div>`).join('')}</div>`,
      onMount: (sh) => sh.querySelectorAll('[data-t]').forEach((b) => (b.onclick = () => { edit('Guests', { ['N' + guestRow]: b.dataset.t }); closeSheet(); toast(`${g.name} → ${b.dataset.t}`); })),
    });
  }

  function openTable(row) {
    const t = m.tables.find((x) => x._row === row);
    const fields = [
      { key: 'name', label: 'Table name', col: 'B', group: 'Table' },
      { key: 'shape', label: 'Shape', type: 'select', options: ['Round', 'Rectangle', 'Long', 'Head'], col: 'C' },
      { key: 'seats', label: 'Seats', type: 'number', col: 'D' },
      { key: 'notes', label: 'Notes', type: 'textarea', col: 'J' },
    ];
    const seated = t.guests;
    const candidates = m.guests.filter((g) => !g.table && g.rsvp !== 'Declined');
    openSheet({
      title: t.name, action: 'Save',
      body: `<div class="calc"><b>${t.assigned}</b> of <b>${t.seatsNum}</b> seats taken${t.left < 0 ? ` – <b style="color:var(--bad)">${-t.left} over</b>` : ''}</div>
        <div class="group-h" style="margin-top:0">Seated here</div>
        <div class="card" style="margin-bottom:14px">${seated.map((g) => `<div class="seat-guest"><div class="n"><b>${esc(g.name)}</b>${g.plus ? ` +${g.plus}` : ''}<div class="muted" style="font-size:12.5px">${esc(g.rsvp || 'Not sent')}</div></div><button data-unseat="${g._row}">Remove</button></div>`).join('') || '<div class="seat-guest muted">Nobody yet</div>'}</div>
        ${candidates.length ? `<div class="group-h">Add guests</div><div class="card" style="margin-bottom:14px">${candidates.map((g) => `<div class="seat-guest"><div class="n"><b>${esc(g.name)}</b>${g.plus ? ` +${g.plus}` : ''}<div class="muted" style="font-size:12.5px">${esc([g.side, g.group].filter(Boolean).join(' · '))} · ${g.party} ${g.party === 1 ? 'seat' : 'seats'}</div></div><button data-seat="${g._row}">Add</button></div>`).join('')}</div>` : ''}
        <button class="btn ghost" data-newguest style="margin-bottom:14px">+ New guest at this table</button>
        ${formHtml(fields, (f) => state.wb.value('Tables', f.col + row))}
        <button class="btn danger" data-deltable>Delete table</button>`,
      onMount: (sh) => {
        wireForm(sh);
        sh.querySelectorAll('[data-unseat]').forEach((b) => (b.onclick = () => { edit('Guests', { ['N' + b.dataset.unseat]: null }); b.closest('.seat-guest').remove(); }));
        sh.querySelectorAll('[data-seat]').forEach((b) => (b.onclick = () => {
          edit('Guests', { ['N' + b.dataset.seat]: t.name });
          b.textContent = '✓'; b.disabled = true;
        }));
        sh.querySelector('[data-newguest]').onclick = () => { closeSheet(true); openGuest(null, { table: t.name, rsvp: 'Pending' }); };
        sh.querySelector('[data-deltable]').onclick = () => confirmSheet(`Delete ${t.name}?`, seated.length ? `${seated.length} invitation(s) seated here will become unseated.` : 'The table will be removed from the floor plan.', 'Delete table', () => {
          edit('Tables', 'BCDGHIJ'.split('').map((c) => [c + row, null]));
          if (seated.length) edit('Guests', seated.map((g) => ['N' + g._row, null]));
        });
      },
      onAction: (sh) => {
        let vals;
        try { vals = readForm(sh, fields); } catch (e) { return toast(e.message); }
        if (!vals.name) return toast('Give the table a name');
        if (vals.name !== t.name && m.tables.some((x) => x.name === vals.name)) return toast('Another table already has that name');
        const changes = fields.filter((f) => !same(state.wb.value('Tables', f.col + row), vals[f.key])).map((f) => [f.col + row, vals[f.key]]);
        closeSheet();
        if (changes.length) edit('Tables', changes);
        if (vals.name !== t.name) {
          // keep guests attached when a table is renamed
          const moved = state.model.guests.filter((g) => g.table === t.name);
          if (moved.length) edit('Guests', moved.map((g) => ['N' + g._row, vals.name]));
        }
      },
    });
  }
}
