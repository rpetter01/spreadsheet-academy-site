/* Spreadsheet Academy charts: chart model, SVG renderer and the Insert Chart panel.
   A chart lives on the sheet: sheet.charts[0] = { type, range, title, xTitle, yTitle, legend, labels, byRow }.
   window.Charts (pure functions, usable without a DOM) + registers the "chart" tool. */
(function (root) {
'use strict';
const { parseRef, numToCol } = root.XL;
const addr = (c, r) => numToCol(c) + r;
const bounds = (rng) => { const [a, z] = rng.split(':'); const A = parseRef(a), Z = parseRef(z || a); return { r1: A.r, c1: A.c, r2: Z.r, c2: Z.c }; };
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']; // Office default chart colours
const TYPES = [['column', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['pie', 'Pie']];

/** Read the chart's source range: first column = categories, first row = series names (when text). */
function chartData(sheet, def) {
  const b = bounds(def.range), g = [];
  for (let r = b.r1; r <= b.r2; r++) { const row = []; for (let c = b.c1; c <= b.c2; c++) row.push({ v: sheet.value(addr(c, r)), d: sheet.display(addr(c, r)) }); g.push(row); }
  const cols = g[0].length, num = (x) => (typeof x.v === 'number' ? x.v : null);
  if (cols === 1) return { categories: g.map((_, i) => String(i + 1)), series: [{ name: 'Series1', values: g.map((r) => num(r[0])) }] };
  const hdr = g[0].slice(1).every((x) => typeof x.v === 'string');
  const body = hdr ? g.slice(1) : g, cats = body.map((r) => r[0].d);
  const names = hdr ? g[0].slice(1).map((x) => x.d) : g[0].slice(1).map((_, i) => 'Series' + (i + 1));
  if (def.byRow) return { categories: names, series: body.map((r, i) => ({ name: cats[i], values: names.map((_, j) => num(r[j + 1])) })) };
  return { categories: cats, series: names.map((n, j) => ({ name: n, values: body.map((r) => num(r[j + 1])) })) };
}

/** A tidy axis: at most ~6 gridlines, rounded to 1 / 2 / 2.5 / 5 steps. */
function niceScale(max) {
  if (!(max > 0)) return { max: 1, step: 0.2 };
  const exp = Math.pow(10, Math.floor(Math.log10(max)));
  for (const mult of [1, 2, 2.5, 5, 10, 20, 25, 50]) { const step = (mult * exp) / 10; if (Math.ceil(max / step - 1e-9) <= 6) return { step, max: Math.ceil(max / step - 1e-9) * step }; }
  return { max, step: max / 5 };
}
const fmt = (n) => (Math.abs(n) >= 1000 ? n.toLocaleString('en-US') : String(+n.toFixed(2)));

function chartSVG(def, data) {
  const W = 640, H = 330, cats = data.categories, S = data.series, ns = S.length, title = (def.title || '').trim();
  const t = (x, y, s, o = {}) => `<text${o.edit ? ` data-edit="${o.edit}" class="ce"` : ''} x="${x}" y="${y}" font-size="${o.size || 12}" fill="${o.fill || '#595959'}" text-anchor="${o.anchor || 'middle'}"${o.rot ? ` transform="rotate(${o.rot} ${x} ${y})"` : ''}${o.bold ? ' font-weight="600"' : ''}>${esc(s)}</text>`;
  let out = `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="${esc(title || def.type + ' chart')}" font-family="Inter, Segoe UI, system-ui, sans-serif"><rect width="${W}" height="${H}" fill="#fff"/>`;
  if (title) out += t(W / 2, 28, title, { size: 17, fill: '#404040', edit: 'title' });
  const legendOn = !!def.legend, top = title ? 46 : 18;
  const legendItems = def.type === 'pie' ? cats : S.map((s) => s.name);
  const legend = (y) => {
    if (!legendOn) return ''; const items = legendItems.map((n, i) => ({ n, w: 26 + String(n).length * 6.6, c: PALETTE[i % PALETTE.length] })); const total = items.reduce((a, i) => a + i.w, 0); let x = (W - total) / 2, s = '';
    items.forEach((i) => { s += `<rect x="${x}" y="${y - 9}" width="11" height="11" rx="2" fill="${i.c}"/>` + t(x + 16, y + 1, i.n, { anchor: 'start', size: 12 }); x += i.w; }); return s;
  };

  if (def.type === 'pie') {
    const vals = (S[0] ? S[0].values : []).map((v) => Math.max(0, v || 0)), sum = vals.reduce((a, b) => a + b, 0) || 1;
    const cx = W / 2 - (legendOn ? 70 : 0), cy = top + (H - top - 18) / 2, R = Math.min(120, (H - top - 30) / 2); let a0 = -Math.PI / 2;
    vals.forEach((v, i) => {
      const a1 = a0 + (v / sum) * 2 * Math.PI, large = a1 - a0 > Math.PI ? 1 : 0, col = PALETTE[i % PALETTE.length];
      const x0 = cx + R * Math.cos(a0), y0 = cy + R * Math.sin(a0), x1 = cx + R * Math.cos(a1), y1 = cy + R * Math.sin(a1);
      out += v / sum >= 0.9999 ? `<circle cx="${cx}" cy="${cy}" r="${R}" fill="${col}" stroke="#fff" stroke-width="2"/>` : `<path d="M${cx},${cy} L${x0},${y0} A${R},${R} 0 ${large} 1 ${x1},${y1} Z" fill="${col}" stroke="#fff" stroke-width="2"/>`;
      if (def.labels && v > 0) { const am = (a0 + a1) / 2; out += t(cx + R * 0.62 * Math.cos(am), cy + R * 0.62 * Math.sin(am) + 4, Math.round((v / sum) * 100) + '%', { fill: '#fff', size: 12, bold: true }); }
      a0 = a1;
    });
    if (legendOn) cats.forEach((n, i) => { const y = cy - (cats.length * 22) / 2 + i * 22 + 11; out += `<rect x="${cx + R + 36}" y="${y - 9}" width="11" height="11" rx="2" fill="${PALETTE[i % PALETTE.length]}"/>` + t(cx + R + 52, y + 1, n, { anchor: 'start' }); });
    return out + '</svg>';
  }

  const all = S.flatMap((s) => s.values.filter((v) => v !== null)), sc = niceScale(Math.max(0, ...all)), horizontal = def.type === 'bar';
  const maxLab = Math.max(...cats.map((c) => String(c).length), 3), leftLab = horizontal ? Math.min(120, 14 + maxLab * 6.8) : 40;
  const left = leftLab + (def.yTitle || def.xTitle ? 22 : 0) + (horizontal ? 0 : 6), right = 24, bottom = 34 + (horizontal ? (def.yTitle ? 22 : 0) : (def.xTitle ? 22 : 0)) + (legendOn ? 28 : 0);
  const px = left, py = top, pw = W - left - right, ph = H - top - bottom, n = cats.length || 1;
  const vx = (v) => px + (v / sc.max) * pw, vy = (v) => py + ph - (v / sc.max) * ph;
  // gridlines + value axis labels
  for (let v = 0; v <= sc.max + 1e-9; v += sc.step) {
    if (horizontal) out += `<line x1="${vx(v)}" y1="${py}" x2="${vx(v)}" y2="${py + ph}" stroke="#D9D9D9"/>` + t(vx(v), py + ph + 16, fmt(v));
    else out += `<line x1="${px}" y1="${vy(v)}" x2="${px + pw}" y2="${vy(v)}" stroke="#D9D9D9"/>` + t(px - 8, vy(v) + 4, fmt(v), { anchor: 'end' });
  }
  const band = (horizontal ? ph : pw) / n;
  cats.forEach((c, i) => { // category labels
    if (horizontal) out += t(px - 8, py + ph - (i + 0.5) * band + 4, c, { anchor: 'end' });
    else out += t(px + (i + 0.5) * band, py + ph + 16, c);
  });
  if (def.type === 'line') {
    S.forEach((s, si) => {
      const col = PALETTE[si % PALETTE.length], pts = s.values.map((v, i) => (v === null ? null : [px + (i + 0.5) * band, vy(v)]));
      out += `<polyline fill="none" stroke="${col}" stroke-width="2.5" stroke-linejoin="round" points="${pts.filter(Boolean).map((p) => p.join(',')).join(' ')}"/>`;
      pts.forEach((p, i) => { if (!p) return; out += `<circle cx="${p[0]}" cy="${p[1]}" r="4" fill="${col}"/>`; if (def.labels) out += t(p[0], p[1] - 9, fmt(s.values[i]), { size: 11 }); });
    });
  } else {
    const gw = band * 0.7, bw = gw / ns;
    S.forEach((s, si) => s.values.forEach((v, i) => {
      if (v === null) return; const col = PALETTE[si % PALETTE.length];
      if (horizontal) { const y = py + ph - (i + 0.5) * band - gw / 2 + (ns - 1 - si) * bw; out += `<rect x="${px}" y="${y}" width="${Math.max(0, vx(v) - px)}" height="${bw}" fill="${col}"/>`; if (def.labels) out += t(vx(v) + 5, y + bw / 2 + 4, fmt(v), { anchor: 'start', size: 11 }); }
      else { const x = px + (i + 0.5) * band - gw / 2 + si * bw; out += `<rect x="${x}" y="${vy(v)}" width="${bw}" height="${Math.max(0, py + ph - vy(v))}" fill="${col}"/>`; if (def.labels) out += t(x + bw / 2, vy(v) - 5, fmt(v), { size: 11 }); }
    }));
  }
  out += `<line x1="${px}" y1="${py + ph}" x2="${px + pw}" y2="${py + ph}" stroke="#BFBFBF"/>` + (horizontal ? `<line x1="${px}" y1="${py}" x2="${px}" y2="${py + ph}" stroke="#BFBFBF"/>` : '');
  // axis titles: category axis (xTitle) and value axis (yTitle)
  const baseY = py + ph + 34;
  if (horizontal) { if (def.xTitle) out += t(14, py + ph / 2, def.xTitle, { rot: -90, size: 12, bold: true, edit: 'xTitle' }); if (def.yTitle) out += t(px + pw / 2, baseY, def.yTitle, { size: 12, bold: true, edit: 'yTitle' }); }
  else { if (def.yTitle) out += t(14, py + ph / 2, def.yTitle, { rot: -90, size: 12, bold: true, edit: 'yTitle' }); if (def.xTitle) out += t(px + pw / 2, baseY, def.xTitle, { size: 12, bold: true, edit: 'xTitle' }); }
  return out + legend(H - 14) + '</svg>';
}

function insertChart(sheet, def) { sheet.charts = [Object.assign({ title: '', xTitle: '', yTitle: '', legend: false, labels: false, byRow: false }, def)]; }
const seriesCount = (sheet, def) => chartData(sheet, def).series.length;

/* ── UI ── */
function mountChart(cfg, host, ctx) {
  const types = TYPES.filter(([k]) => !cfg.types || cfg.types.includes(k)), el = cfg.elements !== false, grid = ctx.grid;
  const prev = document.createElement('div'); prev.className = 'chart-card'; prev.setAttribute('data-notranslate', ''); grid.host.insertAdjacentElement('afterend', prev);
  const cur = () => (grid.sheet.charts && grid.sheet.charts[0]) || null;
  const render = () => {
    const c = cur();
    if (!c) { prev.innerHTML = '<div class="chart-empty">' + (root.I18N ? root.I18N.t('Your chart will appear here.') : 'Your chart will appear here.') + '</div>'; prev.classList.remove('has'); return; }
    prev.classList.add('has'); try { prev.innerHTML = chartSVG(c, chartData(grid.sheet, c)); } catch (e) { prev.innerHTML = '<div class="chart-empty">…</div>'; }
  };
  if (cfg.hidePanel) { host.innerHTML = ''; render(); return { onSelect() {}, onChange: render }; }
  host.innerHTML = `<div class="tool"><div class="tool-h">Insert Chart<small>Insert › Charts</small></div>
    <div class="tool-r"><span class="tool-l">Data</span><b data-sel></b><span class="muted small">(your current selection)</span></div>
    <div class="tool-r"><span class="tool-l">Chart type</span><div class="types" role="radiogroup">${types.map(([k, n], i) => `<button type="button" class="type-btn${i === 0 ? ' on' : ''}" data-type="${k}" role="radio" aria-checked="${i === 0}"><i class="ct ct-${k}"></i><span>${n}</span></button>`).join('')}</div></div>
    ${el ? `<div class="tool-r"><span class="tool-l">Elements</span><input data-title placeholder="Chart title" size="18" aria-label="Chart title"><input data-x placeholder="Category axis title" size="16" aria-label="Category axis title"><input data-y placeholder="Value axis title" size="14" aria-label="Value axis title"></div>
    <div class="tool-r"><span class="tool-l"></span><label class="chk"><input type="checkbox" data-legend> Legend</label><label class="chk"><input type="checkbox" data-labels> Data labels</label><label class="chk"><input type="checkbox" data-byrow> Switch row/column</label></div>` : ''}
    <div class="tool-r"><button class="btn btn-primary btn-sm" data-ins>Insert chart</button>${el ? '<button class="btn btn-secondary btn-sm" data-apply>Apply to chart</button>' : ''}<button class="btn btn-secondary btn-sm" data-rm>Remove chart</button><span class="tool-msg" data-msg></span></div></div>`;
  const q = (s) => host.querySelector(s), msg = (m) => { q('[data-msg]').textContent = m; };
  let type = types[0][0];
  const setType = (k) => { type = k; host.querySelectorAll('.type-btn').forEach((b) => { const on = b.dataset.type === k; b.classList.toggle('on', on); b.setAttribute('aria-checked', on); }); };
  host.querySelectorAll('.type-btn').forEach((b) => { b.onclick = () => setType(b.dataset.type); });
  const load = (c) => { setType(c.type); if (el) { q('[data-title]').value = c.title || ''; q('[data-x]').value = c.xTitle || ''; q('[data-y]').value = c.yTitle || ''; q('[data-legend]').checked = !!c.legend; q('[data-labels]').checked = !!c.labels; q('[data-byrow]').checked = !!c.byRow; } };
  if (cur()) load(cur());
  const form = () => ({ type, title: el ? q('[data-title]').value.trim() : '', xTitle: el ? q('[data-x]').value.trim() : '', yTitle: el ? q('[data-y]').value.trim() : '', legend: el ? q('[data-legend]').checked : false, labels: el ? q('[data-labels]').checked : false, byRow: el ? q('[data-byrow]').checked : false });
  const done = (m) => { render(); msg(m); grid.afterChange('tool'); };
  q('[data-ins]').onclick = () => {
    const rng = grid.rangeStr(), b = bounds(rng); if (b.r1 === b.r2 && b.c1 === b.c2) return msg('Select the data first.');
    const d = Object.assign(form(), { range: rng }), n = seriesCount(grid.sheet, d);
    if (!el) { d.legend = n > 1 || d.type === 'pie'; d.title = n === 1 ? chartData(grid.sheet, d).series[0].name : ''; }
    insertChart(grid.sheet, d); if (el) { /* keep the typed options */ } else load(d); done(`Chart inserted from ${rng}.`);
  };
  const ap = q('[data-apply]'); if (ap) ap.onclick = () => { const c = cur(); if (!c) return msg('Insert a chart first.'); Object.assign(c, form(), { range: c.range }); done('Chart updated.'); };
  q('[data-rm]').onclick = () => { grid.sheet.charts = []; done('Chart removed.'); };
  const api = { onSelect() { q('[data-sel]').textContent = grid.rangeStr(); }, onChange: render }; api.onSelect(); render();
  return api;
}

root.Charts = { chartData, chartSVG, insertChart, niceScale, TYPES };
if (root.Tools && root.Tools.register) root.Tools.register('chart', mountChart);
})(typeof window !== 'undefined' ? window : globalThis);
