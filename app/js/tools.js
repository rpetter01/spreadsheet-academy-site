/* Spreadsheet Academy tools: conditional formatting, filter & sort, PivotTable.
   Pure data functions (usable without a DOM) + small UI panels. window.Tools */
(function (root) {
'use strict';
const { parseRef, numToCol, dateOf } = root.XL;
const addr = (c, r) => numToCol(c) + r;
const bounds = (rng) => { const [a, z] = rng.split(':'); const A = parseRef(a), Z = parseRef(z || a); return { r1: A.r, c1: A.c, r2: Z.r, c2: Z.c }; };
const headers = (sheet, rng) => { const b = bounds(rng), o = []; for (let c = b.c1; c <= b.c2; c++) o.push(String(sheet.display(addr(c, b.r1)))); return o; };
/** the practice sheets follow the app language (Region -> Região), but lessons and answers name fields and values in English: look both up */
const L = (s) => (typeof s === 'string' && root.I18N ? root.I18N.word(s) : s);
const hidx = (H, f) => { let i = H.indexOf(f); if (i < 0 && typeof f === 'string') i = H.indexOf(L(f)); return i; };
const lc = (v) => String(v === null || v === undefined ? '' : v).toLowerCase();

/* ── conditional formatting ── */
function addCF(sheet, rule) { sheet.cf.push(Object.assign({}, rule, { a: L(rule.a), b: L(rule.b) })); }

/* ── filter & sort ── */
function matches(v, op, val) {
  const n = Number(val), isNum = val !== '' && !isNaN(n);
  switch (op) {
    case 'eq': return typeof v === 'number' && isNum ? v === n : lc(v) === lc(val);
    case 'ne': return typeof v === 'number' && isNum ? v !== n : lc(v) !== lc(val);
    case 'gt': return typeof v === 'number' && v > n;
    case 'lt': return typeof v === 'number' && v < n;
    case 'ge': return typeof v === 'number' && v >= n;
    case 'le': return typeof v === 'number' && v <= n;
    case 'contains': return lc(v).includes(lc(val));
    default: return true;
  }
}
function filterRows(sheet, { range, field, op, val }) {
  val = L(val);
  const b = bounds(range), ci = b.c1 + hidx(headers(sheet, range), field); if (ci < b.c1) return;
  const hidden = new Set();
  for (let r = b.r1 + 1; r <= b.r2; r++) if (!matches(sheet.value(addr(ci, r)), op, val)) hidden.add(r);
  sheet.hidden = hidden; sheet.filter = { range, field, op, val }; sheet._inv();
}
function filterValues(sheet, { range, field, values }) {
  const b = bounds(range), ci = b.c1 + hidx(headers(sheet, range), field); if (ci < b.c1) return;
  const ok = new Set(values.map(lc)), hidden = new Set();
  for (let r = b.r1 + 1; r <= b.r2; r++) if (!ok.has(lc(sheet.display(addr(ci, r))))) hidden.add(r);
  sheet.hidden = hidden; sheet.filter = { range, field, op: 'in', val: values }; sheet._inv();
}
function uniqueValues(sheet, range, field) {
  const b = bounds(range), ci = b.c1 + hidx(headers(sheet, range), field), seen = new Map();
  for (let r = b.r1 + 1; r <= b.r2; r++) { const v = sheet.value(addr(ci, r)), d = sheet.display(addr(ci, r)); if (d !== '' && !seen.has(lc(d))) seen.set(lc(d), { d, v }); }
  return [...seen.values()].sort((p, q) => (typeof p.v === 'number' && typeof q.v === 'number' ? p.v - q.v : lc(p.d).localeCompare(lc(q.d))));
}
function clearPivotArea(sheet, dest, rows = 16, cols = 8) { const d = parseRef(dest); for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) sheet.setRaw(addr(d.c + j, d.r + i), ''); sheet._inv(); }
function clearFilter(sheet) { sheet.hidden = new Set(); sheet.filter = null; sheet._inv(); }
function sortRange(sheet, { range, field, dir }) {
  const b = bounds(range), ci = b.c1 + hidx(headers(sheet, range), field); if (ci < b.c1) return;
  const rows = [];
  for (let r = b.r1 + 1; r <= b.r2; r++) { const row = []; for (let c = b.c1; c <= b.c2; c++) { const x = sheet.cells[addr(c, r)]; row.push(x ? { ...x } : null); } rows.push({ row, key: sheet.value(addr(ci, r)) }); }
  rows.sort((p, q) => { const a = p.key, z = q.key; const d = typeof a === 'number' && typeof z === 'number' ? a - z : lc(a).localeCompare(lc(z)); return dir === 'desc' ? -d : d; });
  rows.forEach((o, i) => o.row.forEach((x, j) => { const ref = addr(b.c1 + j, b.r1 + 1 + i); if (x) sheet.cells[ref] = x; else delete sheet.cells[ref]; }));
  sheet._inv();
  if (sheet.filter) filterRows(sheet, sheet.filter); // keep the active filter in step with the new row order
}

/* ── PivotTable ── */
const AGG = { sum: 'Sum', count: 'Count', average: 'Average' };
function sortKeys(a, b) { return typeof a === 'number' && typeof b === 'number' ? a - b : lc(a).localeCompare(lc(b)); }
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function buildPivot(sheet, { range, rowField, colField, valField, agg, dest, filter, filters, sort, showAs, fmt, group }) {
  const b = bounds(range), H = headers(sheet, range), ri = b.c1 + hidx(H, rowField), vi = b.c1 + hidx(H, valField), ci = colField ? b.c1 + hidx(H, colField) : null;
  const fls = (filters || []).concat(filter ? [filter] : []).filter((f) => f && f.field && f.value !== null && f.value !== undefined).map((f) => ({ ci: b.c1 + hidx(H, f.field), value: L(f.value) }));
  const isDate = sheet.fmt(addr(ri, b.r1 + 1)) === 'date', grouped = isDate && group;
  const gl = (v) => { if (typeof v !== 'number') return { l: v, k: null }; const d = dateOf(v), m = d.getUTCMonth(); return group === 'month' ? { l: MONTHS[m], k: m } : group === 'quarter' ? { l: 'Qtr' + (Math.floor(m / 3) + 1), k: Math.floor(m / 3) } : { l: String(d.getUTCFullYear()), k: d.getUTCFullYear() }; };
  const recs = [], keyOf = new Map();
  for (let r = b.r1 + 1; r <= b.r2; r++) {
    if (fls.some((f) => lc(sheet.display(addr(f.ci, r))) !== lc(f.value))) continue;
    const rv = sheet.value(addr(ri, r)), g = grouped ? gl(rv) : { l: rv, k: null };
    if (g.k !== null && !keyOf.has(lc(g.l))) keyOf.set(lc(g.l), g.k);
    recs.push({ r: g.l, c: ci ? sheet.value(addr(ci, r)) : null, v: sheet.value(addr(vi, r)) });
  }
  const uniq = (k) => { const m = new Map(); recs.forEach((x) => { if (x[k] !== null && !m.has(lc(x[k]))) m.set(lc(x[k]), x[k]); }); return [...m.values()].sort((p, q) => (k === 'r' && keyOf.has(lc(p)) && keyOf.has(lc(q)) ? keyOf.get(lc(p)) - keyOf.get(lc(q)) : sortKeys(p, q))); };
  const calc = (list) => {
    const nums = list.map((x) => x.v).filter((v) => typeof v === 'number');
    if (agg === 'count') return list.filter((x) => x.v !== null).length;
    if (!nums.length) return 0;
    const t = nums.reduce((p, q) => p + q, 0); return agg === 'average' ? t / nums.length : t;
  };
  const rowTot = (k) => calc(recs.filter((x) => lc(x.r) === lc(k))), colTot = (c) => calc(recs.filter((x) => lc(x.c) === lc(c))), grand = calc(recs);
  let rk = uniq('r'); const ck = ci ? uniq('c') : [null];
  if (sort === 'rowDesc') rk = rk.reverse(); else if (sort === 'valDesc') rk.sort((p, q) => rowTot(q) - rowTot(p)); else if (sort === 'valAsc') rk.sort((p, q) => rowTot(p) - rowTot(q));
  const pct = showAs && showAs !== 'none';
  const shown = (v, ri2, ci2) => { if (!pct) return v; const den = showAs === 'pctRow' ? (ri2 < 0 ? grand : rowTot(rk[ri2])) : showAs === 'pctCol' ? (ci2 < 0 || !ci ? grand : colTot(ck[ci2])) : grand; return den ? v / den : 0; };
  const d = parseRef(dest), at = (i, j) => addr(d.c + j, d.r + i), put = (i, j, v) => sheet.setRaw(at(i, j), v === null || v === undefined ? '' : String(v));
  const fm = pct ? 'percent' : fmt || undefined, putV = (i, j, v) => { put(i, j, v); sheet.setFmt(at(i, j), fm); };
  const label = `${AGG[agg]} of ${valField}`;
  const prev = sheet._pvSize || { r: 0, c: 0 }, rr = Math.max(prev.r, rk.length + 4), cc = Math.max(prev.c, ck.length + 4);
  for (let i = 0; i < rr; i++) for (let j = 0; j < cc; j++) { sheet.setFmt(at(i, j), undefined); sheet.setRaw(at(i, j), ''); }
  sheet._pvSize = { r: rk.length + 4, c: ck.length + 4 };
  if (!ci) {
    put(0, 0, rowField); put(0, 1, label);
    rk.forEach((k, i) => { put(i + 1, 0, k); if (isDate && !grouped) sheet.setFmt(at(i + 1, 0), 'date'); putV(i + 1, 1, shown(rowTot(k), i, -1)); });
    put(rk.length + 1, 0, 'Grand Total'); putV(rk.length + 1, 1, shown(grand, -1, -1));
  } else {
    put(0, 0, label); ck.forEach((k, j) => put(0, j + 1, k)); put(0, ck.length + 1, 'Grand Total');
    rk.forEach((k, i) => {
      put(i + 1, 0, k); if (isDate && !grouped) sheet.setFmt(at(i + 1, 0), 'date');
      ck.forEach((c, j) => putV(i + 1, j + 1, shown(calc(recs.filter((x) => lc(x.r) === lc(k) && lc(x.c) === lc(c))), i, j)));
      putV(i + 1, ck.length + 1, shown(rowTot(k), i, -1));
    });
    put(rk.length + 1, 0, 'Grand Total'); ck.forEach((c, j) => putV(rk.length + 1, j + 1, shown(colTot(c), -1, j))); putV(rk.length + 1, ck.length + 1, shown(grand, -1, -1));
  }
  sheet._inv();
}

/* ── UI ── */
const opt = (list) => list.map(([v, t]) => `<option value="${v}">${t}</option>`).join('');
const STYLES = [['red', 'Light red fill with dark red text'], ['yellow', 'Yellow fill with dark yellow text'], ['green', 'Green fill with dark green text']];

function mountCF(cfg, host, ctx) {
  host.innerHTML = `<div class="tool"><div class="tool-h">Conditional Formatting<small>Home › Conditional Formatting › Highlight Cells Rules</small></div>
    <div class="tool-r"><span class="tool-l">Applies to</span><b data-sel></b><span class="muted small">(your current selection)</span></div>
    <div class="tool-r"><select data-t aria-label="Rule type">${opt([['gt', 'Greater Than…'], ['lt', 'Less Than…'], ['between', 'Between…'], ['eq', 'Equal To…'], ['text', 'Text that Contains…'], ['dup', 'Duplicate Values…']])}</select>
      <input data-a placeholder="value" size="8" aria-label="Value"><input data-b placeholder="and" size="8" hidden aria-label="Second value"><span class="tool-l">with</span><select data-s aria-label="Format">${opt(STYLES)}</select></div>
    <div class="tool-r"><button class="btn btn-primary btn-sm" data-apply>Apply rule</button><button class="btn btn-secondary btn-sm" data-clear>Clear rules</button><span class="tool-msg" data-msg></span></div></div>`;
  const q = (s) => host.querySelector(s), t = q('[data-t]'), a = q('[data-a]'), b = q('[data-b]');
  const sync = () => { b.hidden = t.value !== 'between'; a.hidden = t.value === 'dup'; a.placeholder = t.value === 'text' ? 'text' : 'value'; };
  t.onchange = sync; sync();
  const api = { onSelect() { q('[data-sel]').textContent = ctx.grid.rangeStr(); } }; api.onSelect();
  q('[data-apply]').onclick = () => {
    const type = t.value, A = type === 'text' || type === 'eq' ? a.value.trim() : Number(a.value), B = Number(b.value);
    if (type !== 'dup' && (a.value.trim() === '' || ((type === 'gt' || type === 'lt' || type === 'between') && isNaN(A)))) return (q('[data-msg]').textContent = 'Enter a value first.');
    if (type === 'between' && isNaN(B)) return (q('[data-msg]').textContent = 'Enter both values.');
    addCF(ctx.grid.sheet, { range: ctx.grid.rangeStr(), type, a: A, b: B, style: q('[data-s]').value });
    q('[data-msg]').textContent = `Rule applied to ${ctx.grid.rangeStr()}.`; ctx.grid.afterChange('tool');
  };
  q('[data-clear]').onclick = () => { ctx.grid.sheet.cf = []; q('[data-msg]').textContent = 'Rules cleared.'; ctx.grid.afterChange('tool'); };
  return api;
}

function mountFilter(cfg, host, ctx) {
  const H = headers(ctx.grid.sheet, cfg.range), b = bounds(cfg.range), total = b.r2 - b.r1;
  const fields = opt(H.map((h) => [h, h]));
  host.innerHTML = `<div class="tool"><div class="tool-h">Sort &amp; Filter<small>Data › Filter / Sort · table ${cfg.range}</small></div>
    <div class="tool-r"><span class="tool-l">Filter</span><select data-f>${fields}</select><select data-o>${opt([['eq', 'equals'], ['ne', 'does not equal'], ['gt', 'is greater than'], ['lt', 'is less than'], ['ge', 'is at least'], ['le', 'is at most'], ['contains', 'contains']])}</select><input data-v placeholder="value" size="10" aria-label="Filter value">
      <button class="btn btn-primary btn-sm" data-go>Apply filter</button><button class="btn btn-secondary btn-sm" data-clr>Clear</button></div>
    <div class="tool-r"><span class="tool-l">Sort</span><select data-sf>${fields}</select><button class="btn btn-secondary btn-sm" data-asc>A → Z / Smallest first</button><button class="btn btn-secondary btn-sm" data-desc>Z → A / Largest first</button></div>
    <div class="tool-r"><span class="tool-msg" data-msg></span></div></div>`;
  const q = (s) => host.querySelector(s), msg = () => { const sh = ctx.grid.sheet; q('[data-msg]').textContent = sh.filter ? `${total - sh.hidden.size} of ${total} records found.` : `${total} records.`; }; msg();
  q('[data-go]').onclick = () => { filterRows(ctx.grid.sheet, { range: cfg.range, field: q('[data-f]').value, op: q('[data-o]').value, val: q('[data-v]').value.trim() }); msg(); ctx.grid.afterChange('tool'); };
  q('[data-clr]').onclick = () => { clearFilter(ctx.grid.sheet); msg(); ctx.grid.afterChange('tool'); };
  const sort = (dir) => { sortRange(ctx.grid.sheet, { range: cfg.range, field: q('[data-sf]').value, dir }); ctx.grid.afterChange('tool'); };
  q('[data-asc]').onclick = () => sort('asc'); q('[data-desc]').onclick = () => sort('desc');
  return {};
}

function mountPivot(cfg, host, ctx) {
  const H = headers(ctx.grid.sheet, cfg.range), fields = opt(H.map((h) => [h, h]));
  host.innerHTML = `<div class="tool"><div class="tool-h">PivotTable Fields<small>Insert › PivotTable · source ${cfg.range} · output at ${cfg.dest}</small></div>
    <div class="tool-r"><span class="tool-l">Rows</span><select data-r>${fields}</select><span class="tool-l">Columns</span><select data-c><option value="">(none)</option>${fields}</select></div>
    <div class="tool-r"><span class="tool-l">Values</span><select data-v>${fields}</select><select data-a>${opt([['sum', 'Sum of'], ['count', 'Count of'], ['average', 'Average of']])}</select>
      <button class="btn btn-primary btn-sm" data-go>Create PivotTable</button><button class="btn btn-secondary btn-sm" data-clr>Clear</button><span class="tool-msg" data-msg></span></div></div>`;
  const q = (s) => host.querySelector(s);
  q('[data-go]').onclick = () => {
    if (q('[data-r]').value === q('[data-v]').value && q('[data-a]').value !== 'count') { q('[data-msg]').textContent = 'Pick a numeric field for Values.'; return; }
    buildPivot(ctx.grid.sheet, { range: cfg.range, rowField: q('[data-r]').value, colField: q('[data-c]').value || null, valField: q('[data-v]').value, agg: q('[data-a]').value, dest: cfg.dest });
    q('[data-msg]').textContent = `PivotTable created at ${cfg.dest}.`; ctx.grid.afterChange('tool');
  };
  q('[data-clr]').onclick = () => {
    const d = parseRef(cfg.dest); for (let i = 0; i < 16; i++) for (let j = 0; j < 8; j++) ctx.grid.sheet.setRaw(addr(d.c + j, d.r + i), '');
    q('[data-msg]').textContent = 'Cleared.'; ctx.grid.afterChange('tool');
  };
  return {};
}

const MOUNT = { cf: mountCF, filter: mountFilter, pivot: mountPivot };
root.Tools = { addCF, filterRows, filterValues, uniqueValues, clearPivotArea, headers, bounds, clearFilter, sortRange, buildPivot, register: (k, f) => { MOUNT[k] = f; }, mount: (cfg, host, ctx) => MOUNT[cfg.kind](cfg, host, ctx) };
})(typeof window !== 'undefined' ? window : globalThis);
