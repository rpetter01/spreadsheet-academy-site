/* Spreadsheet Academy ribbon: an Excel-style ribbon (tabs, groups, icons, galleries, dialogs, contextual tabs)
   that drives the chart, conditional formatting, filter/sort and PivotTable exercises. window.Ribbon */
(function (root) {
'use strict';
const { parseRef, numToCol } = root.XL, Charts = root.Charts, Tools = root.Tools;
const addr = (c, r) => numToCol(c) + r;
const bounds = (rng) => { const [a, z] = rng.split(':'); const A = parseRef(a), Z = parseRef(z || a); return { r1: A.r, c1: A.c, r2: Z.r, c2: Z.c }; };
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const LABELS = new Set(); const L = (s) => { LABELS.add(s); return s; };
const tt = (s, p) => (root.I18N ? root.I18N.t(s, p) : (p ? s.replace(/\{(\w+)\}/g, (_, k) => p[k]) : s));

/* ═════════ icons: flat Office-style, 32px ═════════ */
const BL = '#4472C4', OR = '#ED7D31', GD = '#FFC000', GR = '#70AD47', TL = '#5B9BD5', GY = '#8c8c8c', DK = '#404040', XLG = '#217346', LG = '#d9d9d9';
const IC = {}; const def = (n, body) => { IC[n] = body; };
const frame = '<rect x="4" y="4" width="24" height="24" rx="2" fill="#fff" stroke="#8c8c8c"/>';
const axes = '<path d="M4 3v25h25" stroke="#8c8c8c" fill="none"/>';
const glyph = (s, fill = DK, size = 20, y = 23) => `<text x="16" y="${y}" font-size="${size}" font-weight="700" text-anchor="middle" fill="${fill}" font-family="Segoe UI,Arial,sans-serif">${s}</text>`;
const bars = `<rect x="6" y="16" width="5" height="12" fill="${BL}"/><rect x="13" y="8" width="5" height="20" fill="${OR}"/><rect x="20" y="13" width="5" height="15" fill="${GY}"/>`;
def('gen', `<rect x="7" y="4" width="18" height="24" rx="2" fill="#fff" stroke="${GY}"/><path d="M11 11h10M11 16h10M11 21h6" stroke="${LG}" stroke-width="2"/>`);
def('pivot', `${frame}<rect x="4" y="4" width="24" height="7" fill="#c6e0b4"/><rect x="4" y="11" width="8" height="17" fill="#e2efda"/><path d="M12 11v17M4 11h24" stroke="${GY}"/><path d="M17 20h7m-3-3 3 3-3 3" stroke="${XLG}" stroke-width="2" fill="none"/>`);
def('recpivot', `${IC.pivot}<circle cx="24" cy="8" r="5.5" fill="${GD}"/>${glyph('?', '#fff', 9, 11)}`);
def('table', `<rect x="3" y="5" width="26" height="22" fill="#fff" stroke="${GY}"/><rect x="3" y="5" width="26" height="6" fill="${BL}"/><path d="M3 16.5h26M3 22h26M12 11v16M21 11v16" stroke="#b7b7b7"/>`);
def('pictures', `${frame}<circle cx="12" cy="12" r="3" fill="${GD}"/><path d="M5 27l8-9 5 5 4-4 6 8z" fill="${GR}"/>`);
def('shapes', `<circle cx="11" cy="11" r="6" fill="${TL}"/><rect x="15" y="15" width="12" height="12" rx="1" fill="${OR}"/><path d="M5 27l5-9 5 9z" fill="${GR}"/>`);
def('smartart', `<rect x="11" y="3" width="10" height="7" rx="1" fill="${BL}"/><rect x="3" y="21" width="10" height="7" rx="1" fill="${TL}"/><rect x="19" y="21" width="10" height="7" rx="1" fill="${TL}"/><path d="M16 10v5M8 21v-6h16v6" stroke="${GY}" fill="none"/>`);
def('screenshot', `<rect x="4" y="7" width="24" height="18" rx="2" fill="#fff" stroke="${GY}" stroke-dasharray="3 2"/><circle cx="16" cy="16" r="4.5" fill="${TL}"/>`);
def('reccharts', `${bars}<circle cx="24" cy="8" r="5.5" fill="${GD}"/>${glyph('?', '#fff', 9, 11)}`);
def('colchart', `${axes}<rect x="8" y="15" width="5" height="13" fill="${BL}"/><rect x="15" y="8" width="5" height="20" fill="${BL}"/><rect x="22" y="18" width="5" height="10" fill="${BL}"/>`);
def('barchart', `${axes}<rect x="5" y="7" width="13" height="5" fill="${BL}"/><rect x="5" y="14" width="20" height="5" fill="${BL}"/><rect x="5" y="21" width="9" height="5" fill="${BL}"/>`);
def('hier', `<rect x="3" y="4" width="16" height="14" fill="${BL}"/><rect x="20" y="4" width="9" height="8" fill="${OR}"/><rect x="20" y="13" width="9" height="5" fill="${GD}"/><rect x="3" y="19" width="10" height="9" fill="${GR}"/><rect x="14" y="19" width="15" height="9" fill="${TL}"/>`);
def('stat', `${axes}<rect x="6" y="20" width="4" height="8" fill="${BL}"/><rect x="11" y="10" width="4" height="18" fill="${BL}"/><rect x="16" y="6" width="4" height="22" fill="${BL}"/><rect x="21" y="15" width="4" height="13" fill="${BL}"/>`);
def('linechart', `${axes}<path d="M7 22l6-8 5 5 8-12" stroke="${OR}" stroke-width="2.6" fill="none" stroke-linejoin="round"/><path d="M7 25l6-3 5 1 8-7" stroke="${BL}" stroke-width="2.6" fill="none" stroke-linejoin="round"/>`);
def('piechart', `<circle cx="16" cy="16" r="11" fill="${BL}"/><path d="M16 16V5a11 11 0 0 1 10 15z" fill="${OR}"/><path d="M16 16l10 4a11 11 0 0 1-6 6.5z" fill="${GD}"/>`);
def('scatter', `${axes}<circle cx="10" cy="21" r="2.2" fill="${BL}"/><circle cx="15" cy="15" r="2.2" fill="${BL}"/><circle cx="21" cy="17" r="2.2" fill="${BL}"/><circle cx="25" cy="8" r="2.2" fill="${BL}"/>`);
def('combo', `${axes}<rect x="7" y="17" width="5" height="11" fill="${BL}"/><rect x="15" y="12" width="5" height="16" fill="${BL}"/><path d="M8 11l8-4 9 3" stroke="${OR}" stroke-width="2.4" fill="none"/>`);
def('pivotchart', `<rect x="3" y="4" width="26" height="24" rx="2" fill="#fff" stroke="${GY}"/><rect x="6" y="16" width="5" height="9" fill="${BL}"/><rect x="13" y="9" width="5" height="16" fill="${OR}"/><rect x="20" y="13" width="5" height="12" fill="${GY}"/><path d="M3 9h26" stroke="${GY}"/>`);
def('sparkline', `<rect x="3" y="9" width="26" height="14" fill="#fff" stroke="${LG}"/><path d="M5 20l5-6 4 3 5-6 5 4" stroke="${BL}" stroke-width="2" fill="none"/>`);
def('sparkcol', `<rect x="3" y="9" width="26" height="14" fill="#fff" stroke="${LG}"/><rect x="6" y="15" width="3" height="6" fill="${BL}"/><rect x="11" y="11" width="3" height="10" fill="${BL}"/><rect x="16" y="17" width="3" height="4" fill="${BL}"/><rect x="21" y="12" width="3" height="9" fill="${BL}"/>`);
def('sparkwl', `<rect x="3" y="9" width="26" height="14" fill="#fff" stroke="${LG}"/><rect x="6" y="11" width="3" height="5" fill="${BL}"/><rect x="11" y="16" width="3" height="5" fill="${OR}"/><rect x="16" y="11" width="3" height="5" fill="${BL}"/><rect x="21" y="16" width="3" height="5" fill="${OR}"/>`);
def('slicer', `<rect x="5" y="4" width="22" height="24" rx="2" fill="#fff" stroke="${GY}"/><rect x="8" y="9" width="16" height="5" fill="${BL}"/><rect x="8" y="16" width="16" height="5" fill="${LG}"/><rect x="8" y="23" width="16" height="3" fill="${LG}"/>`);
def('timeline', `<rect x="3" y="9" width="26" height="14" rx="2" fill="#fff" stroke="${GY}"/><rect x="9" y="13" width="12" height="6" fill="${BL}"/><path d="M6 16h2M24 16h2" stroke="${GY}"/>`);
def('link', `<path d="M13 19l6-6M12 14l-3 3a4 4 0 0 0 6 6l3-3M20 18l3-3a4 4 0 0 0-6-6l-3 3" stroke="${BL}" stroke-width="3" fill="none" stroke-linecap="round"/>`);
def('textbox', `<rect x="4" y="6" width="24" height="20" fill="#fff" stroke="${GY}"/>${glyph('A', BL, 18, 22)}`);
def('headfoot', `<rect x="6" y="3" width="20" height="26" fill="#fff" stroke="${GY}"/><rect x="8" y="5" width="16" height="4" fill="${LG}"/><rect x="8" y="23" width="16" height="4" fill="${LG}"/><path d="M10 15h12M10 18h12" stroke="${LG}"/>`);
def('wordart', glyph('A', OR, 28, 26));
def('signature', `<path d="M4 26h24" stroke="${GY}"/><path d="M5 22c4-9 6-9 6-3s3 2 5-2 5 2 11-2" stroke="${DK}" stroke-width="2" fill="none"/>`);
def('object', `<rect x="4" y="6" width="19" height="19" fill="#fff" stroke="${GY}"/><rect x="11" y="11" width="17" height="17" fill="${TL}" opacity=".85"/>`);
def('equation', glyph('π', BL, 26, 25));
def('symbol', glyph('Ω', BL, 26, 25));
// home
def('paste', `<rect x="6" y="5" width="16" height="22" rx="2" fill="#e9e9e9" stroke="${GY}"/><rect x="10" y="3" width="8" height="5" rx="1" fill="${DK}"/><rect x="15" y="12" width="13" height="16" fill="#fff" stroke="${BL}"/><path d="M18 17h7M18 21h7" stroke="${BL}"/>`);
def('cut', `<circle cx="10" cy="24" r="3.5" fill="none" stroke="${BL}" stroke-width="2"/><circle cx="22" cy="24" r="3.5" fill="none" stroke="${BL}" stroke-width="2"/><path d="M12 21L22 5M20 21L10 5" stroke="${DK}" stroke-width="2"/>`);
def('copy', `<rect x="6" y="4" width="14" height="18" fill="#fff" stroke="${GY}"/><rect x="12" y="10" width="14" height="18" fill="#fff" stroke="${BL}"/><path d="M15 16h8M15 20h8" stroke="${BL}"/>`);
def('painter', `<rect x="7" y="4" width="16" height="8" rx="1" fill="${GD}"/><path d="M15 12v5M12 17h6v10h-6z" stroke="${GY}" fill="#fff"/>`);
def('bold', glyph('B', DK, 22, 24)); def('italic', `<text x="16" y="24" font-size="22" font-style="italic" font-weight="600" text-anchor="middle" fill="${DK}" font-family="Georgia,serif">I</text>`); def('underline', `${glyph('U', DK, 20, 21)}<path d="M8 27h16" stroke="${DK}" stroke-width="2"/>`);
def('borders', `<rect x="6" y="6" width="20" height="20" fill="none" stroke="${DK}" stroke-dasharray="2 2"/><path d="M6 26h20" stroke="${DK}" stroke-width="2.5"/>`);
def('fill', `<path d="M8 14l8-9 8 9-8 8z" fill="#fff" stroke="${DK}" stroke-width="1.6"/><rect x="5" y="25" width="22" height="4" fill="${GD}"/>`);
def('fontcolor', `${glyph('A', DK, 20, 22)}<rect x="6" y="25" width="20" height="4" fill="#e00"/>`);
def('alignl', `<path d="M6 8h20M6 14h13M6 20h20M6 26h13" stroke="${DK}" stroke-width="2"/>`); def('alignc', `<path d="M6 8h20M10 14h12M6 20h20M10 26h12" stroke="${DK}" stroke-width="2"/>`); def('alignr', `<path d="M6 8h20M13 14h13M6 20h20M13 26h13" stroke="${DK}" stroke-width="2"/>`);
def('wrap', `<path d="M5 9h22M5 16h18a4 4 0 0 1 0 8h-6" stroke="${DK}" stroke-width="2" fill="none"/><path d="M19 21l-3 3 3 3" stroke="${BL}" stroke-width="2" fill="none"/>`);
def('merge', `<rect x="4" y="8" width="24" height="16" fill="#fff" stroke="${GY}"/><path d="M9 16h14M12 13l-3 3 3 3M20 13l3 3-3 3" stroke="${BL}" stroke-width="1.8" fill="none"/>`);
def('currency', glyph('$', XLG, 22, 24)); def('percent', glyph('%', DK, 20, 23)); def('comma', glyph('9,', DK, 17, 22));
def('decinc', `${glyph('.0', DK, 13, 17)}<path d="M5 24h12m-3-3 3 3-3 3" stroke="${BL}" stroke-width="1.8" fill="none"/>`); def('decdec', `${glyph('.0', DK, 13, 17)}<path d="M17 24H5m3-3-3 3 3 3" stroke="${OR}" stroke-width="1.8" fill="none"/>`);
def('cf', `<rect x="3" y="4" width="22" height="22" fill="#fff" stroke="${GY}"/><rect x="3" y="4" width="22" height="5" fill="${LG}"/><rect x="5" y="11" width="7" height="5" fill="#ffc7ce"/><rect x="14" y="11" width="9" height="5" fill="#c6efce"/><rect x="5" y="18" width="9" height="5" fill="#ffeb9c"/><path d="M20 24l8 0m-3-3 3 3-3 3" stroke="${XLG}" stroke-width="2" fill="none"/>`);
def('fmttable', `<rect x="3" y="5" width="26" height="22" fill="#fff" stroke="${GY}"/><rect x="3" y="5" width="26" height="6" fill="${TL}"/><rect x="3" y="16" width="26" height="5" fill="#ddebf7"/><path d="M12 11v16M21 11v16" stroke="#b7b7b7"/>`);
def('cellstyles', `<rect x="4" y="5" width="24" height="7" fill="${TL}"/><rect x="4" y="14" width="24" height="7" fill="#ffc7ce"/><rect x="4" y="23" width="24" height="5" fill="#c6efce"/>`);
def('inscells', `<rect x="4" y="10" width="24" height="16" fill="#fff" stroke="${GY}"/><path d="M16 10v16M4 18h24" stroke="${GY}"/><circle cx="23" cy="8" r="5" fill="${GR}"/><path d="M23 5.5v5M20.5 8h5" stroke="#fff" stroke-width="1.6"/>`);
def('delcells', `<rect x="4" y="10" width="24" height="16" fill="#fff" stroke="${GY}"/><path d="M16 10v16M4 18h24" stroke="${GY}"/><circle cx="23" cy="8" r="5" fill="#d33"/><path d="M20.5 8h5" stroke="#fff" stroke-width="1.8"/>`);
def('fmtcells', `<rect x="4" y="6" width="24" height="20" fill="#fff" stroke="${GY}"/><path d="M4 12h24M13 6v20" stroke="${GY}"/><circle cx="23" cy="22" r="4" fill="${GD}"/>`);
def('autosum', glyph('Σ', OR, 28, 26)); def('filldown', `<rect x="9" y="4" width="14" height="24" fill="#fff" stroke="${GY}"/><path d="M16 9v13m-4-4 4 4 4-4" stroke="${BL}" stroke-width="2" fill="none"/>`);
def('clear', `<path d="M8 20l9-12 8 6-9 12H11z" fill="#f4b6b6" stroke="${DK}"/><path d="M5 27h22" stroke="${GY}"/>`);
def('sortfilter', `${glyph('A', BL, 12, 12)}${glyph('Z', DK, 12, 24)}<path d="M9 14v10m-3-3 3 3 3-3" stroke="${BL}" stroke-width="1.6" fill="none"/><path d="M17 7h12l-4.5 6v7l-3 2v-9z" fill="${TL}" stroke="${DK}" stroke-width="1"/>`);
def('find', `<circle cx="13" cy="13" r="7" fill="#fff" stroke="${BL}" stroke-width="2.4"/><path d="M18 18l9 9" stroke="${DK}" stroke-width="3"/>`);
// data
def('getdata', `<ellipse cx="14" cy="8" rx="9" ry="4" fill="${BL}"/><path d="M5 8v14c0 2 4 4 9 4s9-2 9-4V8" fill="${TL}" stroke="${BL}"/><path d="M22 20h8m-3-3 3 3-3 3" stroke="${XLG}" stroke-width="2" fill="none"/>`);
def('refresh', `<path d="M26 14a10 10 0 0 0-18-4M6 18a10 10 0 0 0 18 4" stroke="${XLG}" stroke-width="3" fill="none"/><path d="M5 5v7h7M27 27v-7h-7" stroke="${XLG}" stroke-width="2.5" fill="none"/>`);
def('queries', `<rect x="4" y="5" width="18" height="22" fill="#fff" stroke="${GY}"/><path d="M8 11h10M8 16h10M8 21h6" stroke="${LG}" stroke-width="2"/><path d="M22 18h7m-3-3 3 3-3 3" stroke="${BL}" stroke-width="2" fill="none"/>`);
def('sortaz', `${glyph('A', BL, 12, 12)}${glyph('Z', DK, 12, 25)}<path d="M23 6v20m-4-4 4 4 4-4" stroke="${BL}" stroke-width="2" fill="none"/>`); def('sortza', `${glyph('Z', DK, 12, 12)}${glyph('A', BL, 12, 25)}<path d="M23 6v20m-4-4 4 4 4-4" stroke="${BL}" stroke-width="2" fill="none"/>`);
def('sort', `${glyph('A', BL, 11, 12)}${glyph('Z', DK, 11, 25)}<rect x="15" y="5" width="13" height="22" fill="#fff" stroke="${GY}"/><path d="M18 11h7M18 16h7M18 21h7" stroke="${BL}"/>`);
def('filter', `<path d="M4 5h24l-9 11v9l-6 3V16z" fill="${TL}" stroke="${DK}" stroke-width="1.4"/>`); def('clearfilter', `<path d="M4 5h20l-8 10v8l-4 2V15z" fill="${TL}" stroke="${DK}"/><path d="M21 19l8 8M29 19l-8 8" stroke="#d33" stroke-width="2.4"/>`);
def('reapply', `<path d="M4 5h20l-8 10v8l-4 2V15z" fill="${TL}" stroke="${DK}"/><path d="M26 17a5 5 0 1 0 1 6" stroke="${XLG}" stroke-width="2.4" fill="none"/>`); def('advanced', `<path d="M4 5h20l-8 10v8l-4 2V15z" fill="${TL}" stroke="${DK}"/><path d="M22 22h8M26 18v8" stroke="${OR}" stroke-width="2.4"/>`);
def('texttocol', `<rect x="3" y="6" width="26" height="20" fill="#fff" stroke="${GY}"/><path d="M12 6v20M21 6v20M3 13h26" stroke="${LG}"/><path d="M6 18h4m6 0h3" stroke="${BL}" stroke-width="2"/>`);
def('flashfill', `<path d="M17 3l-9 15h6l-2 11 10-15h-6z" fill="${GD}" stroke="${OR}"/>`); def('removedups', `<rect x="4" y="5" width="18" height="22" fill="#fff" stroke="${GY}"/><path d="M8 11h10M8 16h10M8 21h10" stroke="${LG}" stroke-width="2"/><circle cx="23" cy="22" r="6" fill="#d33"/><path d="M20 22h6" stroke="#fff" stroke-width="2"/>`);
def('validation', `<rect x="4" y="6" width="20" height="14" fill="#fff" stroke="${GY}"/><path d="M8 13l3 3 5-6" stroke="${XLG}" stroke-width="2.4" fill="none"/><circle cx="24" cy="23" r="5" fill="#d33"/><path d="M22 21l4 4M26 21l-4 4" stroke="#fff" stroke-width="1.6"/>`);
def('consolidate', `<rect x="3" y="5" width="12" height="9" fill="#fff" stroke="${GY}"/><rect x="17" y="5" width="12" height="9" fill="#fff" stroke="${GY}"/><rect x="10" y="19" width="12" height="9" fill="${TL}" stroke="${GY}"/><path d="M9 14v3h14v-3M16 17v2" stroke="${GY}" fill="none"/>`);
def('whatif', `<rect x="4" y="5" width="16" height="20" fill="#fff" stroke="${GY}"/><path d="M8 11h8M8 16h8" stroke="${LG}" stroke-width="2"/>${glyph('?', OR, 22, 28)}`);
def('forecast', `${axes}<path d="M7 22l5-5 5 3 4-9" stroke="${BL}" stroke-width="2.4" fill="none"/><path d="M21 11l7-4" stroke="${OR}" stroke-width="2.4" stroke-dasharray="3 2"/>`);
def('group', `<path d="M6 8h4M6 24h4M10 8v16" stroke="${GY}" stroke-width="2" fill="none"/><rect x="14" y="5" width="14" height="6" fill="${LG}"/><rect x="14" y="14" width="14" height="6" fill="${LG}"/><rect x="14" y="23" width="14" height="5" fill="${LG}"/>`);
def('subtotal', `<rect x="4" y="5" width="24" height="22" fill="#fff" stroke="${GY}"/><path d="M4 20h24" stroke="${GY}"/>${glyph('Σ', OR, 14, 18)}`);
// chart design
def('addelem', `${axes}<rect x="8" y="16" width="5" height="12" fill="${BL}"/><rect x="15" y="10" width="5" height="18" fill="${BL}"/><circle cx="24" cy="9" r="6" fill="${GR}"/><path d="M24 6v6M21 9h6" stroke="#fff" stroke-width="2"/>`);
def('quicklayout', `<rect x="4" y="4" width="24" height="24" fill="#fff" stroke="${GY}"/><rect x="7" y="7" width="18" height="4" fill="${LG}"/><rect x="7" y="14" width="8" height="11" fill="${BL}"/><rect x="17" y="18" width="8" height="7" fill="${OR}"/>`);
def('colors', `<circle cx="11" cy="12" r="6" fill="${BL}"/><circle cx="21" cy="12" r="6" fill="${OR}"/><circle cx="16" cy="21" r="6" fill="${GD}"/>`);
def('switch', `<rect x="3" y="4" width="18" height="14" fill="#fff" stroke="${GY}"/><path d="M3 9h18M9 4v14" stroke="${LG}"/><path d="M22 16a7 7 0 0 1-9 10M22 22v-6h-6" stroke="${XLG}" stroke-width="2.2" fill="none"/>`);
def('selectdata', `<rect x="3" y="5" width="18" height="16" fill="#fff" stroke="${GY}"/><path d="M3 10h18M9 5v16" stroke="${LG}"/><rect x="15" y="14" width="13" height="13" fill="${TL}" opacity=".8"/><path d="M18 21h7M21.5 17.5v7" stroke="#fff" stroke-width="2"/>`);
def('changetype', `${axes}<rect x="8" y="16" width="5" height="12" fill="${BL}"/><rect x="15" y="9" width="5" height="19" fill="${BL}"/><path d="M20 8l8 0m-3-3 3 3-3 3" stroke="${OR}" stroke-width="2.2" fill="none"/>`);
def('movechart', `<rect x="3" y="5" width="18" height="16" fill="#fff" stroke="${GY}"/><rect x="11" y="13" width="18" height="14" fill="${TL}" opacity=".8"/><path d="M20 20h6" stroke="#fff" stroke-width="2"/>`);
// menu thumbnails
def('m-col', `<rect x="1" y="1" width="30" height="30" fill="#fff" stroke="${LG}"/><rect x="5" y="14" width="6" height="14" fill="${BL}"/><rect x="13" y="7" width="6" height="21" fill="${BL}"/><rect x="21" y="17" width="6" height="11" fill="${BL}"/>`);
def('m-bar', `<rect x="1" y="1" width="30" height="30" fill="#fff" stroke="${LG}"/><rect x="4" y="5" width="14" height="6" fill="${BL}"/><rect x="4" y="13" width="23" height="6" fill="${BL}"/><rect x="4" y="21" width="10" height="6" fill="${BL}"/>`);
def('m-line', `<rect x="1" y="1" width="30" height="30" fill="#fff" stroke="${LG}"/><path d="M5 24l7-9 6 5 9-14" stroke="${BL}" stroke-width="2.4" fill="none"/>`);
def('m-linem', `${IC['m-line']}<circle cx="5" cy="24" r="2" fill="${BL}"/><circle cx="12" cy="15" r="2" fill="${BL}"/><circle cx="18" cy="20" r="2" fill="${BL}"/><circle cx="27" cy="6" r="2" fill="${BL}"/>`);
def('m-pie', `<rect x="1" y="1" width="30" height="30" fill="#fff" stroke="${LG}"/><circle cx="16" cy="16" r="11" fill="${BL}"/><path d="M16 16V5a11 11 0 0 1 10 15z" fill="${OR}"/><path d="M16 16l10 4a11 11 0 0 1-6 6.5z" fill="${GD}"/>`);
const icon = (n) => `<svg class="ri" viewBox="0 0 32 32" aria-hidden="true">${IC[n] || IC.gen}</svg>`;

/* ═════════ tabs, groups and buttons ═════════ */
const B = (id, label, ic, o = {}) => Object.assign({ id, label: label ? L(label) : '', ic, size: 'lg', tip: label ? L(label) : '' }, o);
const S = (id, label, ic, o = {}) => Object.assign({ id, label: label ? L(label) : '', ic, size: 'sm', tip: o.tip ? L(o.tip) : label ? L(label) : '' }, o);
const COL = (...items) => ({ col: items });
const SEL = (label, w) => ({ sel: label, w });
const MINI = (id, label, ic) => ({ mini: true, id, label: L(label), ic });
const ROWS = (...rows) => ({ rows });
const GRP = (name, ...items) => ({ name: L(name), items });
const TABS = {
  home: [
    GRP('Clipboard', B('paste', 'Paste', 'paste', { menu: 1 }), COL(S('cut', 'Cut', 'cut'), S('copy', 'Copy', 'copy'), S('painter', 'Format Painter', 'painter'))),
    GRP('Font', ROWS([SEL('Calibri', 100), SEL('11', 40)], [MINI('bold', 'Bold', 'bold'), MINI('italic', 'Italic', 'italic'), MINI('underline', 'Underline', 'underline'), MINI('borders', 'Borders', 'borders'), MINI('fill', 'Fill Color', 'fill'), MINI('fontcolor', 'Font Color', 'fontcolor')])),
    GRP('Alignment', ROWS([MINI('alignl', 'Align Left', 'alignl'), MINI('alignc', 'Center', 'alignc'), MINI('alignr', 'Align Right', 'alignr')], [MINI('wrap', 'Wrap Text', 'wrap'), MINI('merge', 'Merge & Center', 'merge')])),
    GRP('Number', ROWS([SEL('General', 118)], [MINI('currency', 'Accounting Number Format', 'currency'), MINI('percent', 'Percent Style', 'percent'), MINI('comma', 'Comma Style', 'comma'), MINI('decinc', 'Increase Decimal', 'decinc'), MINI('decdec', 'Decrease Decimal', 'decdec')])),
    GRP('Styles', B('cf', 'Conditional Formatting', 'cf', { menu: 1 }), B('fmttable', 'Format as Table', 'fmttable', { menu: 1 }), B('cellstyles', 'Cell Styles', 'cellstyles', { menu: 1 })),
    GRP('Cells', B('inscells', 'Insert', 'inscells', { menu: 1 }), B('delcells', 'Delete', 'delcells', { menu: 1 }), B('fmtcells', 'Format', 'fmtcells', { menu: 1 })),
    GRP('Editing', COL(S('autosum', 'AutoSum', 'autosum', { menu: 1 }), S('fillm', 'Fill', 'filldown', { menu: 1 }), S('clearm', 'Clear', 'clear', { menu: 1 })), B('sortfilter', 'Sort & Filter', 'sortfilter', { menu: 1 }), B('find', 'Find & Select', 'find', { menu: 1 })),
  ],
  insert: [
    GRP('Tables', B('pivot', 'PivotTable', 'pivot', { menu: 1 }), B('recpivot', 'Recommended PivotTables', 'recpivot'), B('table', 'Table', 'table')),
    GRP('Illustrations', B('pictures', 'Pictures', 'pictures', { menu: 1 }), B('shapes', 'Shapes', 'shapes', { menu: 1 }), COL(S('smartart', 'SmartArt', 'smartart'), S('screenshot', 'Screenshot', 'screenshot', { menu: 1 }))),
    GRP('Charts', B('reccharts', 'Recommended Charts', 'reccharts'),
      COL(S('ch-col', '', 'colchart', { menu: 1, tip: 'Insert Column or Bar Chart' }), S('ch-line', '', 'linechart', { menu: 1, tip: 'Insert Line or Area Chart' }), S('ch-pie', '', 'piechart', { menu: 1, tip: 'Insert Pie or Doughnut Chart' })),
      COL(S('ch-hier', '', 'hier', { menu: 1, tip: 'Insert Hierarchy Chart' }), S('ch-stat', '', 'stat', { menu: 1, tip: 'Insert Statistic Chart' }), S('ch-scatter', '', 'scatter', { menu: 1, tip: 'Insert Scatter (X, Y) or Bubble Chart' })),
      COL(S('ch-combo', '', 'combo', { menu: 1, tip: 'Insert Combo Chart' })), B('pivotchart', 'PivotChart', 'pivotchart', { menu: 1 })),
    GRP('Sparklines', COL(S('sp-line', 'Line', 'sparkline'), S('sp-col', 'Column', 'sparkcol'), S('sp-wl', 'Win/Loss', 'sparkwl'))),
    GRP('Filters', B('slicer', 'Slicer', 'slicer'), B('timeline', 'Timeline', 'timeline')),
    GRP('Links', B('link', 'Link', 'link')),
    GRP('Text', B('textbox', 'Text Box', 'textbox'), B('headfoot', 'Header & Footer', 'headfoot'), B('wordart', 'WordArt', 'wordart', { menu: 1 }), COL(S('signature', 'Signature Line', 'signature', { menu: 1 }), S('object', 'Object', 'object'))),
    GRP('Symbols', B('equation', 'Equation', 'equation', { menu: 1 }), B('symbol', 'Symbol', 'symbol')),
  ],
  pagelayout: [
    GRP('Themes', B('themes', 'Themes', 'colors', { menu: 1 }), COL(S('thcolors', 'Colors', 'colors'), S('thfonts', 'Fonts', 'gen'), S('theffects', 'Effects', 'gen'))),
    GRP('Page Setup', B('margins', 'Margins', 'gen', { menu: 1 }), B('orientation', 'Orientation', 'gen', { menu: 1 }), B('size', 'Size', 'gen', { menu: 1 }), B('printarea', 'Print Area', 'gen', { menu: 1 }), B('breaks', 'Breaks', 'gen', { menu: 1 }), B('background', 'Background', 'pictures'), B('printtitles', 'Print Titles', 'gen')),
    GRP('Scale to Fit', COL(S('width', 'Width: Automatic', 'gen'), S('height', 'Height: Automatic', 'gen'), S('scale', 'Scale: 100%', 'gen'))),
    GRP('Arrange', B('bringfwd', 'Bring Forward', 'gen', { menu: 1 }), B('sendback', 'Send Backward', 'gen', { menu: 1 }), B('selpane', 'Selection Pane', 'gen')),
  ],
  formulas: [
    GRP('Function Library', B('insfunc', 'Insert Function', 'gen'), B('autosum2', 'AutoSum', 'autosum', { menu: 1 }), B('recent', 'Recently Used', 'gen', { menu: 1 }), B('financial', 'Financial', 'gen', { menu: 1 }), B('logical', 'Logical', 'gen', { menu: 1 }), B('textf', 'Text', 'gen', { menu: 1 }), B('datetime', 'Date & Time', 'gen', { menu: 1 }), B('lookupf', 'Lookup & Reference', 'gen', { menu: 1 }), B('mathf', 'Math & Trig', 'gen', { menu: 1 })),
    GRP('Defined Names', B('namemgr', 'Name Manager', 'gen'), COL(S('definename', 'Define Name', 'gen', { menu: 1 }), S('useinformula', 'Use in Formula', 'gen', { menu: 1 }), S('createfromsel', 'Create from Selection', 'gen'))),
    GRP('Formula Auditing', COL(S('traceprec', 'Trace Precedents', 'gen'), S('tracedep', 'Trace Dependents', 'gen'), S('removearrows', 'Remove Arrows', 'gen', { menu: 1 })), B('showformulas', 'Show Formulas', 'gen'), B('watch', 'Watch Window', 'gen')),
    GRP('Calculation', B('calcopt', 'Calculation Options', 'gen', { menu: 1 }), B('calcnow', 'Calculate Now', 'gen'), B('calcsheet', 'Calculate Sheet', 'gen')),
  ],
  data: [
    GRP('Get & Transform Data', B('getdata', 'Get Data', 'getdata', { menu: 1 }), COL(S('fromtext', 'From Text/CSV', 'getdata'), S('fromweb', 'From Web', 'getdata'), S('fromtable', 'From Table/Range', 'table')), COL(S('recentsrc', 'Recent Sources', 'queries'), S('existconn', 'Existing Connections', 'queries'))),
    GRP('Queries & Connections', B('refreshall', 'Refresh All', 'refresh', { menu: 1 }), COL(S('queriesconn', 'Queries & Connections', 'queries'), S('properties', 'Properties', 'gen'), S('editlinks', 'Edit Links', 'link'))),
    GRP('Sort & Filter', COL(S('sort-az', 'Sort A to Z', 'sortaz'), S('sort-za', 'Sort Z to A', 'sortza')), B('sort', 'Sort', 'sort'), B('filter', 'Filter', 'filter'), COL(S('clearfilter', 'Clear', 'clearfilter'), S('reapply', 'Reapply', 'reapply'), S('advanced', 'Advanced', 'advanced'))),
    GRP('Data Tools', B('texttocol', 'Text to Columns', 'texttocol'), COL(S('flashfill', 'Flash Fill', 'flashfill'), S('removedups', 'Remove Duplicates', 'removedups'), S('validation', 'Data Validation', 'validation', { menu: 1 })), COL(S('consolidate', 'Consolidate', 'consolidate'), S('relationships', 'Relationships', 'table'), S('whatif', 'What-If Analysis', 'whatif', { menu: 1 }))),
    GRP('Forecast', B('forecast', 'Forecast Sheet', 'forecast')),
    GRP('Outline', B('group', 'Group', 'group', { menu: 1 }), B('ungroup', 'Ungroup', 'group', { menu: 1 }), B('subtotal', 'Subtotal', 'subtotal')),
  ],
  review: [
    GRP('Proofing', B('spelling', 'Spelling', 'gen'), B('thesaurus', 'Thesaurus', 'gen')),
    GRP('Accessibility', B('checkacc', 'Check Accessibility', 'gen', { menu: 1 })),
    GRP('Insights', B('smartlookup', 'Smart Lookup', 'find')),
    GRP('Language', B('translate', 'Translate', 'gen')),
    GRP('Comments', B('newcomment', 'New Comment', 'gen'), B('delcomment', 'Delete', 'delcells'), B('prevcomment', 'Previous', 'gen'), B('nextcomment', 'Next', 'gen'), B('showcomments', 'Show Comments', 'gen')),
    GRP('Protect', B('protectsheet', 'Protect Sheet', 'gen'), B('protectwb', 'Protect Workbook', 'gen'), B('allowedit', 'Allow Edit Ranges', 'gen')),
  ],
  view: [
    GRP('Workbook Views', B('normalv', 'Normal', 'table'), B('pagebreakv', 'Page Break Preview', 'gen'), B('pagelayoutv', 'Page Layout', 'gen'), B('customviews', 'Custom Views', 'gen')),
    GRP('Show', COL(S('rulerv', 'Ruler', 'gen'), S('formulabarv', 'Formula Bar', 'gen'), S('gridlinesv', 'Gridlines', 'gen'), S('headingsv', 'Headings', 'gen'))),
    GRP('Zoom', B('zoom', 'Zoom', 'find'), B('zoom100', '100%', 'gen'), B('zoomsel', 'Zoom to Selection', 'find')),
    GRP('Window', B('newwindow', 'New Window', 'gen'), B('arrangeall', 'Arrange All', 'gen'), B('freeze', 'Freeze Panes', 'gen', { menu: 1 })),
    GRP('Macros', B('macros', 'Macros', 'gen', { menu: 1 })),
  ],
  help: [
    GRP('Help', B('helpb', 'Help', 'gen'), B('contactsupport', 'Contact Support', 'gen'), B('feedbackb', 'Feedback', 'gen'), B('training', 'Show Training', 'gen'), B('community', 'Community', 'gen')),
  ],
  pivotanalyze: [
    GRP('PivotTable', B('pvname', 'PivotTable Name', 'pivot'), B('pvopts', 'Options', 'gen', { menu: 1 })),
    GRP('Active Field', COL(S('pvfield', 'Field Settings', 'gen'), S('pvdrill', 'Drill Down', 'gen'))),
    GRP('Group', B('groupsel', 'Group Selection', 'group'), COL(S('ungroupsel', 'Ungroup', 'group'), S('groupfield', 'Group Field', 'group'))),
    GRP('Filter', B('slicerp', 'Insert Slicer', 'slicer'), B('timelinep', 'Insert Timeline', 'timeline'), B('filterconn', 'Filter Connections', 'gen')),
    GRP('Data', B('refreshp', 'Refresh', 'refresh', { menu: 1 }), B('changesrc', 'Change Data Source', 'table', { menu: 1 })),
    GRP('Actions', B('clearp', 'Clear', 'clear', { menu: 1 }), B('selectp', 'Select', 'gen', { menu: 1 }), B('movep', 'Move PivotTable', 'movechart')),
    GRP('Show', B('fieldlist', 'Field List', 'gen'), B('pmbuttons', '+/- Buttons', 'gen')),
  ],
  chartdesign: [
    GRP('Chart Layouts', B('addelem', 'Add Chart Element', 'addelem', { menu: 1 }), B('quicklayout', 'Quick Layout', 'quicklayout', { menu: 1 })),
    GRP('Chart Styles', B('colors', 'Change Colors', 'colors', { menu: 1 }), { styles: true }),
    GRP('Data', B('switch', 'Switch Row/Column', 'switch'), B('selectdata', 'Select Data', 'selectdata')),
    GRP('Type', B('changetype', 'Change Chart Type', 'changetype', { menu: 1 })),
    GRP('Location', B('movechart', 'Move Chart', 'movechart')),
  ],
  chartformat: [
    GRP('Current Selection', COL(S('cursel', 'Chart Area', 'gen'), S('fmtsel', 'Format Selection', 'gen'), S('resetstyle', 'Reset to Match Style', 'gen'))),
    GRP('Insert Shapes', B('insshapes', 'Shapes', 'shapes', { menu: 1 })),
    GRP('Shape Styles', B('shfill', 'Shape Fill', 'fill', { menu: 1 }), B('shoutline', 'Shape Outline', 'borders', { menu: 1 }), B('sheffects', 'Shape Effects', 'gen', { menu: 1 })),
    GRP('WordArt Styles', B('wafill', 'Text Fill', 'fontcolor', { menu: 1 }), B('waoutline', 'Text Outline', 'gen', { menu: 1 })),
    GRP('Arrange', B('fwd', 'Bring Forward', 'gen'), B('back', 'Send Backward', 'gen')),
    GRP('Size', B('height', 'Height', 'gen'), B('width', 'Width', 'gen')),
  ],
};
const TAB_NAMES = { home: 'Home', insert: 'Insert', pagelayout: 'Page Layout', formulas: 'Formulas', data: 'Data', review: 'Review', view: 'View', help: 'Help', chartdesign: 'Chart Design', chartformat: 'Format', pivotanalyze: 'PivotTable Analyze' };
Object.values(TAB_NAMES).forEach(L); L('File'); L('Tell me what you want to do');

/* ═════════ menus & dialogs: all static strings registered for translation ═════════ */
const M = {
  colbar: [{ head: L('2-D Column') }, { id: 'column', label: L('Clustered Column'), ic: 'm-col' }, { head: L('2-D Bar') }, { id: 'bar', label: L('Clustered Bar'), ic: 'm-bar' }, { sep: 1 }, { label: L('More Column Charts…'), off: 1 }],
  line: [{ head: L('2-D Line') }, { id: 'line', label: L('Line'), ic: 'm-line' }, { id: 'line', label: L('Line with Markers'), ic: 'm-linem' }, { sep: 1 }, { label: L('More Line Charts…'), off: 1 }],
  pie: [{ head: L('2-D Pie') }, { id: 'pie', label: L('Pie'), ic: 'm-pie' }, { label: L('Doughnut'), off: 1 }, { sep: 1 }, { label: L('More Pie Charts…'), off: 1 }],
};
const CFD = {
  gt: { title: L('Greater Than'), text: L('Format cells that are GREATER THAN:'), n: 1 }, lt: { title: L('Less Than'), text: L('Format cells that are LESS THAN:'), n: 1 },
  between: { title: L('Between'), text: L('Format cells that are BETWEEN:'), n: 2 }, eq: { title: L('Equal To'), text: L('Format cells that are EQUAL TO:'), n: 1 },
  text: { title: L('Text That Contains'), text: L('Format cells that contain the text:'), n: 1 }, dup: { title: L('Duplicate Values'), text: L('Format cells that contain:'), n: 0 },
};
const CFS = [['red', L('Light Red Fill with Dark Red Text')], ['yellow', L('Yellow Fill with Dark Yellow Text')], ['green', L('Green Fill with Dark Green Text')], [null, L('Light Red Fill')], [null, L('Red Text')], [null, L('Red Border')], [null, L('Custom Format…')]];
['OK', 'Cancel', 'with', 'and', 'Duplicate', 'Unique', 'values', 'Create PivotTable', 'Choose the data that you want to analyze', 'Select a table or range', 'Table/Range:', 'Choose where you want the PivotTable report to be placed', 'New Worksheet', 'Existing Worksheet', 'Location:',
  'PivotTable Fields', 'Choose fields to add to report:', 'Drag fields between areas below:', 'Filters', 'Columns', 'Rows', 'Values', 'To build a report, choose fields from the list.', 'Move to Column Labels', 'Move to Row Labels', 'Move to Values', 'Remove Field', 'Summarize Values By', 'Sum', 'Count', 'Average',
  'Select Data Source', 'Chart data range:', 'Custom AutoFilter', 'Show rows where:', 'Search', '(Select All)', 'Number Filters', 'Text Filters', 'Equals…', 'Does Not Equal…', 'Greater Than…', 'Greater Than Or Equal To…', 'Less Than…', 'Less Than Or Equal To…', 'Contains…',
  'Sort Smallest to Largest', 'Sort Largest to Smallest', 'Highlight Cells Rules', 'Top/Bottom Rules', 'Data Bars', 'Color Scales', 'Icon Sets', 'New Rule…', 'Clear Rules', 'Manage Rules…', 'Clear Rules from Selected Cells', 'Clear Rules from Entire Sheet', 'A Date Occurring…', 'More Rules…',
  'Custom Sort…', 'Filter', 'Clear', 'Reapply', 'Axes', 'Axis Titles', 'Chart Title', 'Data Labels', 'Data Table', 'Error Bars', 'Gridlines', 'Legend', 'Lines', 'Trendline', 'Up/Down Bars', 'None', 'Above Chart', 'Centered Overlay', 'Primary Horizontal', 'Primary Vertical', 'Right', 'Top', 'Left', 'Bottom', 'Center', 'Inside End', 'Outside End',
  'Column', 'Bar', 'Line', 'Pie', 'This feature is not part of this exercise.', 'Enter a valid range such as A1:B7.', 'Enter a number.', 'Enter a value.'].forEach(L);

/* ═════════ the ribbon controller ═════════ */
const ALWAYS = new Set(['bold', 'currency', 'percent', 'comma', 'autosum']);
const BY_KIND = {
  chart: ['ch-col', 'ch-line', 'ch-pie', 'addelem', 'switch', 'selectdata', 'changetype'],
  pivot: ['pivot', 'refreshp', 'slicerp', 'groupsel', 'ungroupsel', 'groupfield'], cf: ['cf'], filter: ['sort-az', 'sort-za', 'filter', 'clearfilter', 'sortfilter', 'reapply'],
};

function mount(cfg, host, ctx) {
  const kind = cfg.kind, grid = ctx.grid, sheet = () => grid.sheet, chartDef = () => (sheet().charts && sheet().charts[0]) || null;
  const enabled = new Set([...ALWAYS, ...(BY_KIND[kind] || []), ...(cfg.enable || [])]);
  const showRibbon = !cfg.hidePanel;
  let tab = cfg.tab || ({ chart: 'insert', pivot: 'insert', cf: 'home', filter: 'data' }[kind] || 'home');
  if (kind === 'chart' && chartDef()) tab = 'chartdesign';
  let pop = null, afOn = !!(sheet().filter), noteT = null;

  /* ── side card: the chart, or the PivotTable fields pane, shown right next to the sheet ── */
  const side = document.createElement('div');
  if (kind === 'chart') { side.className = 'side-card chart-card'; side.setAttribute('data-notranslate', ''); } else if (kind === 'pivot') side.className = 'side-card pv-pane';
  if (kind === 'chart' || kind === 'pivot') grid.host.insertAdjacentElement('afterend', side);
  const reveal = (el) => { const r = el.getBoundingClientRect(); if (r.top < 70 || r.bottom > window.innerHeight - 96) el.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); };
  const flash = (el) => { el.classList.remove('flash'); void el.offsetWidth; el.classList.add('flash'); };

  const note = (m) => { const n = host.querySelector('.rb-note'); if (!n) return; n.textContent = tt(m); n.classList.add('on'); clearTimeout(noteT); noteT = setTimeout(() => n.classList.remove('on'), 3200); };
  const changed = () => grid.afterChange('tool');
  const tabsFor = () => ['home', 'insert', 'pagelayout', 'formulas', 'data', 'review', 'view', 'help'].concat(chartDef() ? ['chartdesign', 'chartformat'] : []).concat(kind === 'pivot' && pv.created ? ['pivotanalyze'] : []);

  /* ── chart ── */
  function renderChart(doFlash) {
    if (kind !== 'chart') return; const d = chartDef();
    if (!d) { side.innerHTML = `<div class="chart-empty">${tt('Your chart will appear here.')}</div>`; side.classList.remove('has'); return; }
    side.classList.add('has'); side.innerHTML = Charts.chartSVG(d, Charts.chartData(sheet(), d));
    if (doFlash) { flash(side); reveal(side); }
  }
  function editText(prop, el) {
    const d = chartDef(); if (!d) return; const r = el.getBoundingClientRect(), pr = side.getBoundingClientRect();
    const inp = document.createElement('input'); inp.className = 'ce-input'; inp.value = d[prop] || ''; inp.setAttribute('aria-label', prop);
    const w = Math.max(150, r.width + 30); inp.style.width = w + 'px'; inp.style.left = Math.max(4, r.left - pr.left + r.width / 2 - w / 2) + 'px'; inp.style.top = (r.top - pr.top + r.height / 2 - 15) + 'px';
    side.appendChild(inp); inp.focus(); inp.select(); let fin = false;
    const done = (ok) => { if (fin) return; fin = true; const v = inp.value.trim(); inp.remove(); if (ok) { d[prop] = v; renderChart(); changed(); } };
    inp.onkeydown = (e) => { e.stopPropagation(); if (e.key === 'Enter') done(true); else if (e.key === 'Escape') done(false); };
    inp.onblur = () => done(true);
  }
  if (kind === 'chart') side.addEventListener('click', (e) => { const el = e.target.closest('[data-edit]'); if (el) editText(el.dataset.edit, el); });
  const setProp = (p, v, edit) => { const d = chartDef(); if (!d) return; d[p] = v; renderChart(); changed(); if (edit) { const el = side.querySelector(`[data-edit="${p}"]`); if (el) editText(p, el); } };
  function insertChartType(type) {
    let rng = grid.rangeStr(); const b = bounds(rng);
    if (b.r1 === b.r2 && b.c1 === b.c2) return note('Select the data first.');
    const d = { type, range: rng, title: '', xTitle: '', yTitle: '', legend: false, labels: false, byRow: false }, data = Charts.chartData(sheet(), d);
    d.legend = data.series.length > 1 || type === 'pie'; d.title = data.series.length === 1 ? data.series[0].name : '';
    Charts.insertChart(sheet(), d); tab = 'chartdesign'; draw(); renderChart(true); changed();
  }
  const addElemItems = () => {
    const d = chartDef(), pie = d.type === 'pie';
    const toggleAxis = (p) => { if (d[p]) setProp(p, ''); else setProp(p, 'Axis Title', true); };
    const mk = (labels, f) => labels.map((l) => ({ label: L(l), act: () => f(l) }));
    return [
      { label: 'Axes', off: 1 }, { label: 'Axis Titles', off: pie, sub: [{ label: 'Primary Horizontal', check: !!d.xTitle, act: () => toggleAxis('xTitle') }, { label: 'Primary Vertical', check: !!d.yTitle, act: () => toggleAxis('yTitle') }] },
      { label: 'Chart Title', sub: [{ label: 'None', check: !d.title, act: () => setProp('title', '') }, { label: 'Above Chart', check: !!d.title, act: () => { if (!d.title) setProp('title', 'Chart Title', true); } }, { label: 'Centered Overlay', off: 1 }] },
      { label: 'Data Labels', sub: [{ label: 'None', check: !d.labels, act: () => setProp('labels', false) }, ...['Center', 'Inside End', 'Outside End'].map((l) => ({ label: l, check: d.labels, act: () => setProp('labels', true) }))] },
      { label: 'Data Table', off: 1 }, { label: 'Error Bars', off: 1 }, { label: 'Gridlines', off: 1 },
      { label: 'Legend', sub: [{ label: 'None', check: !d.legend, act: () => setProp('legend', false) }, ...['Right', 'Top', 'Left', 'Bottom'].map((l) => ({ label: l, check: d.legend && l === 'Bottom', act: () => setProp('legend', true) }))] },
      { label: 'Lines', off: 1 }, { label: 'Trendline', off: 1 }, { label: 'Up/Down Bars', off: 1 },
    ];
  };

  /* ── conditional formatting ── */
  function cfDialog(type) {
    const k = CFD[type], styles = CFS.map(([v, n]) => `<option value="${v || ''}"${v ? '' : ' disabled'}>${n}</option>`).join('');
    const inputs = k.n === 2 ? '<input data-a size="10" aria-label="value"> <span>and</span> <input data-b size="10" aria-label="value">' : k.n === 1 ? '<input data-a size="14" aria-label="value">' : '<select data-d><option>Duplicate</option><option disabled>Unique</option></select> <span>values</span>';
    openDialog({ title: k.title, body: `<p>${k.text}</p><div class="rb-drow">${inputs}<span>with</span><select data-s>${styles}</select></div>`, ok: (d) => {
      const rng = grid.rangeStr(), st = d.querySelector('[data-s]').value, ra = d.querySelector('[data-a]'), rb = d.querySelector('[data-b]'); let a, b;
      if (k.n >= 1) { const sa = ra.value.trim(); if (sa === '') return tt('Enter a value.'); a = type === 'text' || (type === 'eq' && isNaN(Number(sa))) ? sa : Number(sa); if ((type === 'gt' || type === 'lt' || type === 'between') && isNaN(a)) return tt('Enter a number.'); }
      if (k.n === 2) { b = Number(rb.value); if (rb.value.trim() === '' || isNaN(b)) return tt('Enter a number.'); }
      Tools.addCF(sheet(), { range: rng, type, a, b, style: st }); changed(); return true;
    } });
  }

  /* ── filter & sort ── */
  const af = document.createElement('div'); af.className = 'af-layer';
  function placeArrows() {
    af.innerHTML = ''; if (kind !== 'filter' || !afOn) return;
    if (!af.isConnected) grid.wrap.appendChild(af);
    const b = bounds(cfg.range), w = grid.wrap.getBoundingClientRect(), H = Tools.headers(sheet(), cfg.range);
    for (let c = b.c1; c <= b.c2; c++) {
      const td = grid.td[b.r1][c], r = td.getBoundingClientRect(), field = H[c - b.c1], active = sheet().filter && sheet().filter.field === field;
      const bt = document.createElement('button'); bt.type = 'button'; bt.className = 'af-btn' + (active ? ' on' : ''); bt.title = tt('Filter'); bt.setAttribute('aria-label', 'Filter ' + field);
      bt.style.left = (r.right - w.left - 20) + 'px'; bt.style.top = (r.top - w.top + 4) + 'px'; bt.innerHTML = active ? '<i class="af-f"></i>' : '<i class="af-c"></i>';
      bt.onmousedown = (e) => e.preventDefault(); bt.onclick = (e) => { e.stopPropagation(); filterMenu(field, bt); }; af.appendChild(bt);
    }
  }
  const sortBy = (field, dir) => { Tools.sortRange(sheet(), { range: cfg.range, field, dir }); changed(); };
  const activeField = () => { const b = bounds(cfg.range), c = grid.cur.c; if (c < b.c1 || c > b.c2) return null; return Tools.headers(sheet(), cfg.range)[c - b.c1]; };
  function sortActive(dir) { const f = activeField(); if (!f) return note('Select a cell in the table first.'); sortBy(f, dir); }
  function toggleFilter() { afOn = !afOn; if (!afOn) Tools.clearFilter(sheet()); placeArrows(); changed(); draw(); }
  function customFilter(field, op) {
    const ops = [['eq', 'equals'], ['ne', 'does not equal'], ['gt', 'is greater than'], ['ge', 'is at least'], ['lt', 'is less than'], ['le', 'is at most'], ['contains', 'contains']];
    openDialog({ title: L('Custom AutoFilter'), body: `<p>${tt('Show rows where:')}</p><div class="rb-drow"><b>${esc(field)}</b><select data-o>${ops.map(([k, n]) => `<option value="${k}"${k === op ? ' selected' : ''}>${n}</option>`).join('')}</select><input data-v size="12" aria-label="value"></div>`, ok: (d) => {
      const v = d.querySelector('[data-v]').value.trim(); if (v === '') return tt('Enter a value.');
      Tools.filterRows(sheet(), { range: cfg.range, field, op: d.querySelector('[data-o]').value, val: v }); afOn = true; placeArrows(); changed(); return true;
    } });
  }
  function filterMenu(field, anchor) {
    closePop(); const vals = Tools.uniqueValues(sheet(), cfg.range, field), numeric = vals.length && vals.every((x) => typeof x.v === 'number');
    const cur = sheet().filter && sheet().filter.field === field ? sheet().filter : null; const pick = new Set(cur ? (cur.op === 'in' ? cur.val.map((x) => String(x).toLowerCase()) : vals.filter((x) => { const r = sheet(); return true; }).map(() => '')) : vals.map((x) => x.d.toLowerCase()));
    const m = document.createElement('div'); m.className = 'rb-menu af-menu';
    m.innerHTML = `<button class="rb-mi" data-k="asc"><i class="rb-mck"></i><span>${tt(numeric ? 'Sort Smallest to Largest' : 'Sort A to Z')}</span></button><button class="rb-mi" data-k="desc"><i class="rb-mck"></i><span>${tt(numeric ? 'Sort Largest to Smallest' : 'Sort Z to A')}</span></button><div class="rb-sep"></div>
      <button class="rb-mi${cur ? '' : ' off'}" data-k="clear"><i class="rb-mck"></i><span>${tt('Clear Filter From "{f}"', { f: field })}</span></button>
      <button class="rb-mi" data-k="nf"><i class="rb-mck"></i><span>${tt(numeric ? 'Number Filters' : 'Text Filters')}</span><i class="rb-arrow">▸</i></button><div class="rb-sep"></div>
      <input class="af-search" placeholder="${tt('Search')}" aria-label="Search"><div class="af-list"><label><input type="checkbox" data-all> <span>${tt('(Select All)')}</span></label>${vals.map((x) => `<label><input type="checkbox" data-v="${esc(x.d)}"${pick.has(x.d.toLowerCase()) ? ' checked' : ''}> <span>${esc(x.d)}</span></label>`).join('')}</div>
      <div class="af-foot"><button class="btn btn-primary btn-sm" data-k="ok">${tt('OK')}</button><button class="btn btn-secondary btn-sm" data-k="cancel">${tt('Cancel')}</button></div>`;
    place(m, anchor.getBoundingClientRect(), true); host.appendChild(m); pop = m;
    const boxes = [...m.querySelectorAll('[data-v]')], all = m.querySelector('[data-all]'); const sync = () => { all.checked = boxes.every((b) => b.checked); };
    sync(); all.onchange = () => { boxes.forEach((b) => { if (!b.closest('label').hidden) b.checked = all.checked; }); };
    boxes.forEach((b) => { b.onchange = sync; });
    m.querySelector('.af-search').oninput = (e) => { const q = e.target.value.toLowerCase(); boxes.forEach((b) => { b.closest('label').hidden = !!q && !b.dataset.v.toLowerCase().includes(q); }); };
    m.onclick = (e) => {
      const b = e.target.closest('[data-k]'); if (!b || b.classList.contains('off')) return; const k = b.dataset.k;
      if (k === 'asc' || k === 'desc') { closePop(); sortBy(field, k); }
      else if (k === 'clear') { closePop(); Tools.clearFilter(sheet()); placeArrows(); changed(); }
      else if (k === 'cancel') closePop();
      else if (k === 'ok') { const on = boxes.filter((x) => x.checked).map((x) => x.dataset.v); closePop(); if (on.length === boxes.length) Tools.clearFilter(sheet()); else Tools.filterValues(sheet(), { range: cfg.range, field, values: on }); placeArrows(); changed(); }
      else if (k === 'nf') {
        const r = b.getBoundingClientRect(), items = (numeric ? [['eq', 'Equals…'], ['ne', 'Does Not Equal…'], ['gt', 'Greater Than…'], ['ge', 'Greater Than Or Equal To…'], ['lt', 'Less Than…'], ['le', 'Less Than Or Equal To…']] : [['eq', 'Equals…'], ['ne', 'Does Not Equal…'], ['contains', 'Contains…']]).map(([op, l]) => ({ label: l, act: () => customFilter(field, op) }));
        showMenu(items, r.right - 4, r.top - 4, 1, true);
      }
    };
  }

  /* ── PivotTable ── */
  const pv = { created: false, rows: [], cols: [], vals: [], filters: [], filterVal: null, sort: null, showAs: 'none', nf: null, group: null, slicer: null, range: cfg.range, dest: cfg.dest };
  const AGGN = { sum: 'Sum', count: 'Count', average: 'Average' };
  const SHOWAS = [['none', 'No Calculation'], ['pctGrand', '% of Grand Total'], ['pctCol', '% of Column Total'], ['pctRow', '% of Row Total']];
  const NFS = [['', 'General'], ['currency', 'Currency'], ['comma', 'Number'], ['percent', 'Percentage']];
  ['No Calculation', '% of Grand Total', '% of Column Total', '% of Row Total', 'General', 'Number', 'Percentage', 'Value Field Settings…', 'Summarize values by:', 'Show values as:', 'Number format:', 'Add to Report Filter', 'Add to Row Labels', 'Add to Column Labels', 'Add to Values', 'Move to Report Filter', 'Filter', '(All)', 'Sort A to Z', 'Sort Z to A', 'Sort Largest to Smallest', 'Sort Smallest to Largest', 'Currency'].forEach(L);
  function pivotDialog() {
    openDialog({ title: L('Create PivotTable'), body: `<p><b>${tt('Choose the data that you want to analyze')}</b></p><label class="rb-rad"><input type="radio" checked disabled> ${tt('Select a table or range')}</label><div class="rb-drow"><span>${tt('Table/Range:')}</span><input value="Sheet1!$${cfg.range.split(':').map((x) => x.replace(/^([A-Z]+)(\d+)$/, '$1$$$2')).join(':$')}" readonly size="22"></div>
      <p><b>${tt('Choose where you want the PivotTable report to be placed')}</b></p><label class="rb-rad"><input type="radio" disabled> ${tt('New Worksheet')}</label><label class="rb-rad"><input type="radio" checked disabled> ${tt('Existing Worksheet')}</label><div class="rb-drow"><span>${tt('Location:')}</span><input value="Sheet1!$${cfg.dest.replace(/^([A-Z]+)(\d+)$/, '$1$$$2')}" readonly size="14"></div>`, ok: () => { pv.created = true; renderPivot(); draw(); reveal(side); flash(side); changed(); return true; } });
  }
  const numericField = (f) => { const b = bounds(cfg.range), ci = b.c1 + Tools.headers(sheet(), cfg.range).indexOf(f); return typeof sheet().value(addr(ci, b.r1 + 1)) === 'number' && sheet().fmt(addr(ci, b.r1 + 1)) !== 'date'; };
  const isDateField = (f) => { const b = bounds(cfg.range), ci = b.c1 + Tools.headers(sheet(), cfg.range).indexOf(f); return ci >= b.c1 && sheet().fmt(addr(ci, b.r1 + 1)) === 'date'; };
  function groupDialog() {
    if (!pv.rows[0] || !isDateField(pv.rows[0])) return note('Select a date field in Rows first.');
    const r = (v, t) => `<label class="rb-rad"><input type="radio" name="gb" value="${v}"${pv.group === v ? ' checked' : ''}> ${tt(t)}</label>`;
    openDialog({ title: L('Grouping'), body: `<p>${tt('By:')}</p>${r('month', 'Months')}${r('quarter', 'Quarters')}${r('year', 'Years')}`, ok: (d) => { const c = d.querySelector('[name=gb]:checked'); if (!c) return tt('Pick how to group the dates.'); pv.group = c.value; rebuildPivot(); return true; } });
  }
  function slicerDialog() {
    if (!pv.created) return note('Create a PivotTable first.');
    const H = Tools.headers(sheet(), cfg.range);
    openDialog({ title: L('Insert Slicers'), body: `<p>${tt('Select the field to filter by:')}</p>${H.map((f, i) => `<label class="rb-rad"><input type="radio" name="sf" value="${esc(f)}"${i === 1 ? ' checked' : ''}> ${esc(f)}</label>`).join('')}`, ok: (d) => { pv.slicer = { field: d.querySelector('[name=sf]:checked').value, value: null }; rebuildPivot(); reveal(side); return true; } });
  }
  const dropField = (f) => { pv.rows = pv.rows.filter((x) => x !== f); pv.cols = pv.cols.filter((x) => x !== f); pv.vals = pv.vals.filter((x) => x.f !== f); if (pv.filters[0] === f) { pv.filters = []; pv.filterVal = null; } };
  const moveField = (f, to) => { dropField(f); if (to === 'rows') pv.rows = [f]; else if (to === 'cols') pv.cols = [f]; else if (to === 'filters') { pv.filters = [f]; pv.filterVal = null; } else pv.vals = [{ f, agg: numericField(f) ? 'sum' : 'count' }]; rebuildPivot(); };
  function rebuildPivot() {
    const s = sheet();
    if (pv.rows.length && pv.vals.length) Tools.buildPivot(s, { range: cfg.range, dest: cfg.dest, rowField: pv.rows[0], colField: pv.cols[0] || null, valField: pv.vals[0].f, agg: pv.vals[0].agg, filters: [pv.filters[0] ? { field: pv.filters[0], value: pv.filterVal } : null, pv.slicer ? { field: pv.slicer.field, value: pv.slicer.value } : null].filter(Boolean), sort: pv.sort, showAs: pv.showAs, fmt: pv.nf, group: pv.group });
    else Tools.clearPivotArea(s, cfg.dest);
    renderPivot(); changed();
  }
  function valueSettings() {
    const v = pv.vals[0]; if (!v) return; const sel = (list, cur) => list.map(([k, t]) => `<option value="${k}"${k === cur ? ' selected' : ''}>${tt(t)}</option>`).join('');
    openDialog({ title: L('Value Field Settings'), body: `<p><b>${esc(v.f)}</b></p><div class="rb-drow"><span class="xl-lbl">${tt('Summarize values by:')}</span><select data-ag>${sel(Object.entries(AGGN), v.agg)}</select></div><div class="rb-drow"><span class="xl-lbl">${tt('Show values as:')}</span><select data-sa>${sel(SHOWAS, pv.showAs)}</select></div><div class="rb-drow"><span class="xl-lbl">${tt('Number format:')}</span><select data-nf>${sel(NFS, pv.nf || '')}</select></div>`,
      ok: (d) => { v.agg = d.querySelector('[data-ag]').value; pv.showAs = d.querySelector('[data-sa]').value; pv.nf = d.querySelector('[data-nf]').value || null; rebuildPivot(); return true; } });
  }
  const slicerHtml = () => {
    if (!pv.slicer) return '';
    const items = Tools.uniqueValues(sheet(), cfg.range, pv.slicer.field), on = (x) => pv.slicer.value !== null && String(pv.slicer.value).toLowerCase() === x.d.toLowerCase();
    return `<div class="pv-slicer" role="group" aria-label="${esc(pv.slicer.field)}"><div class="pv-sl-h"><b>${esc(pv.slicer.field)}</b><span><button type="button" data-slclear title="${tt('Clear Filter')}" aria-label="${tt('Clear Filter')}">⨯</button><button type="button" data-slremove title="${tt('Remove')}" aria-label="${tt('Remove')}">🗑</button></span></div><div class="pv-sl-b">${items.map((x) => `<button type="button" class="pv-sl-i${on(x) ? ' on' : ''}" data-v="${esc(x.d)}" aria-pressed="${on(x)}">${esc(x.d)}</button>`).join('')}</div></div>`;
  };
  function renderPivot() {
    if (kind !== 'pivot') return;
    if (!pv.created) { side.classList.remove('has'); side.innerHTML = `<div class="chart-empty">${tt('PivotTable Fields')}</div>`; return; }
    side.classList.add('has'); const H = Tools.headers(sheet(), cfg.range), used = new Set([...pv.rows, ...pv.cols, ...pv.filters, ...pv.vals.map((v) => v.f)]);
    const chip = (f, area, label) => `<div class="pv-chip" data-f="${esc(f)}" data-area="${area}"><span>${esc(label || f)}</span><button type="button" class="pv-dd" aria-label="menu">▾</button></div>`;
    const fval = pv.filterVal === null ? '' : ' (' + pv.filterVal + ')';
    side.innerHTML = `<div class="pv-h">${tt('PivotTable Fields')}</div><p class="pv-s">${tt('Choose fields to add to report:')}</p>
      <div class="pv-fields">${H.map((f) => `<div class="pv-f"><label><input type="checkbox" data-f="${esc(f)}"${used.has(f) ? ' checked' : ''}> <span>${esc(f)}</span></label><button type="button" class="pv-fm" data-f="${esc(f)}" aria-label="${esc(f)}">▾</button></div>`).join('')}</div>
      <p class="pv-s">${tt('Drag fields between areas below:')}</p>
      <div class="pv-areas"><div class="pv-a"><b>${tt('Filters')}</b><div class="pv-box">${pv.filters.map((f) => chip(f, 'filters', f + fval)).join('')}</div></div><div class="pv-a"><b>${tt('Columns')}</b><div class="pv-box">${pv.cols.map((f) => chip(f, 'cols')).join('')}</div></div>
      <div class="pv-a"><b>${tt('Rows')}</b><div class="pv-box">${pv.rows.map((f) => chip(f, 'rows')).join('')}</div></div><div class="pv-a"><b>${tt('Values')}</b><div class="pv-box">${pv.vals.map((v) => chip(v.f, 'vals', tt(AGGN[v.agg] + ' of {f}', { f: v.f }))).join('')}</div></div></div>
      ${slicerHtml()}
      ${pv.rows.length && pv.vals.length ? '' : `<p class="pv-hint">${tt('To build a report, choose fields from the list.')}</p>`}`;
    side.querySelectorAll('.pv-sl-i').forEach((bt) => { bt.onclick = () => { pv.slicer.value = pv.slicer.value !== null && String(pv.slicer.value).toLowerCase() === bt.dataset.v.toLowerCase() ? null : bt.dataset.v; rebuildPivot(); }; });
    const slc = side.querySelector('[data-slclear]'); if (slc) slc.onclick = () => { pv.slicer.value = null; rebuildPivot(); };
    const slx = side.querySelector('[data-slremove]'); if (slx) slx.onclick = () => { pv.slicer = null; rebuildPivot(); };
    side.querySelectorAll('.pv-f input').forEach((cb) => { cb.onchange = () => { const f = cb.dataset.f; if (cb.checked) { if (numericField(f)) { pv.vals = [{ f, agg: 'sum' }]; } else { pv.rows = [f]; } } else dropField(f); rebuildPivot(); }; });
    side.querySelectorAll('.pv-fm').forEach((bt) => { bt.onclick = (e) => { e.stopPropagation(); const f = bt.dataset.f, r = bt.getBoundingClientRect();
      showMenu([{ label: 'Add to Report Filter', act: () => moveField(f, 'filters') }, { label: 'Add to Row Labels', act: () => moveField(f, 'rows') }, { label: 'Add to Column Labels', act: () => moveField(f, 'cols') }, { label: 'Add to Values', act: () => moveField(f, 'vals') }], r.left, r.bottom + 2, 0); }; });
    side.querySelectorAll('.pv-chip').forEach((ch) => { ch.querySelector('.pv-dd').onclick = (e) => { e.stopPropagation(); const f = ch.dataset.f, area = ch.dataset.area, r = e.target.getBoundingClientRect();
      const items = [{ label: 'Move to Report Filter', act: () => moveField(f, 'filters'), off: area === 'filters' }, { label: 'Move to Row Labels', act: () => moveField(f, 'rows'), off: area === 'rows' }, { label: 'Move to Column Labels', act: () => moveField(f, 'cols'), off: area === 'cols' }, { label: 'Move to Values', act: () => moveField(f, 'vals'), off: area === 'vals' }];
      if (area === 'filters') items.unshift({ label: 'Filter', sub: [{ label: '(All)', check: pv.filterVal === null, act: () => { pv.filterVal = null; rebuildPivot(); } }, ...Tools.uniqueValues(sheet(), cfg.range, f).map((x) => ({ label: x.d, check: pv.filterVal !== null && String(pv.filterVal).toLowerCase() === x.d.toLowerCase(), act: () => { pv.filterVal = x.d; rebuildPivot(); } }))] }, { sep: 1 });
      if (area === 'rows' && isDateField(f)) items.unshift({ label: 'Group…', act: groupDialog }, { label: 'Ungroup', act: () => { pv.group = null; rebuildPivot(); }, off: !pv.group }, { sep: 1 });
      if (area === 'rows') items.unshift({ label: 'Sort A to Z', check: pv.sort === 'rowAsc', act: () => { pv.sort = 'rowAsc'; rebuildPivot(); } }, { label: 'Sort Z to A', check: pv.sort === 'rowDesc', act: () => { pv.sort = 'rowDesc'; rebuildPivot(); } }, { label: 'Sort Largest to Smallest', check: pv.sort === 'valDesc', act: () => { pv.sort = 'valDesc'; rebuildPivot(); } }, { label: 'Sort Smallest to Largest', check: pv.sort === 'valAsc', act: () => { pv.sort = 'valAsc'; rebuildPivot(); } }, { sep: 1 });
      if (area === 'vals') { items.unshift({ label: 'Value Field Settings…', act: valueSettings }, { sep: 1 }); items.push({ label: 'Summarize Values By', sub: ['sum', 'count', 'average'].map((a2) => ({ label: AGGN[a2], check: pv.vals[0].agg === a2, act: () => { pv.vals[0].agg = a2; rebuildPivot(); } })) }); }
      items.push({ sep: 1 }, { label: 'Remove Field', act: () => { dropField(f); rebuildPivot(); } });
      showMenu(items, r.left, r.bottom + 2, 0); }; });
  }

  /* ── menus ── */
  function closePop() { host.querySelectorAll('.rb-menu').forEach((m) => m.remove()); pop = null; }
  function place(el, r, below) { el.style.position = 'fixed'; el.style.left = Math.min(r.left, window.innerWidth - 250) + 'px'; el.style.top = (below ? r.bottom + 1 : r.top) + 'px'; el.style.maxHeight = (window.innerHeight - r.bottom - 12) + 'px'; }
  function showMenu(items, x, y, level, side2) {
    if (level === 0) closePop(); else host.querySelectorAll('.rb-menu').forEach((m) => { if (+m.dataset.lv >= level) m.remove(); });
    const m = document.createElement('div'); m.className = 'rb-menu'; m.dataset.lv = level;
    m.innerHTML = items.map((it, i) => it.sep ? '<div class="rb-sep"></div>' : it.head ? `<div class="rb-mh">${tt(it.head)}</div>` : `<button type="button" class="rb-mi${it.off ? ' off' : ''}" data-i="${i}">${it.ic ? icon(it.ic) : it.sw ? `<i class="rb-sw" style="background:${it.sw}"></i>` : `<i class="rb-mck">${it.check ? '✓' : ''}</i>`}<span>${tt(it.label)}</span>${it.sub ? '<i class="rb-arrow">▸</i>' : ''}</button>`).join('');
    m.style.position = 'fixed'; m.style.left = Math.min(x, window.innerWidth - 250) + 'px'; m.style.top = Math.min(y, window.innerHeight - 40 - Math.min(380, items.length * 30)) + 'px';
    host.appendChild(m); pop = m;
    m.onclick = (e) => { const b = e.target.closest('.rb-mi'); if (!b || b.classList.contains('off')) return; const it = items[+b.dataset.i];
      if (it.sub) { const r = b.getBoundingClientRect(); showMenu(it.sub, r.right - 2, r.top - 4, level + 1); return; }
      closePop(); if (it.id && it.ic) insertChartType(it.id); else if (it.act) it.act(); if (!grid.dlgEl) grid.focus(); };
    m.onmouseover = (e) => { const b = e.target.closest('.rb-mi'); if (!b || b.classList.contains('off')) return; const it = items[+b.dataset.i]; if (it.sub) { const r = b.getBoundingClientRect(); showMenu(it.sub, r.right - 2, r.top - 4, level + 1); } else host.querySelectorAll('.rb-menu').forEach((mm) => { if (+mm.dataset.lv > level) mm.remove(); }); };
  }
  function openDialog({ title, body, ok }) {
    closePop(); const ov = document.createElement('div'); ov.className = 'rb-dov';
    ov.innerHTML = `<div class="rb-dlg" role="dialog" aria-modal="true"><div class="rb-dt"><span>${tt(title)}</span><button type="button" class="rb-dx" aria-label="Close">×</button></div><div class="rb-db">${body}</div><div class="rb-err" aria-live="polite"></div><div class="rb-df"><button type="button" class="btn btn-primary btn-sm" data-ok>${tt('OK')}</button><button type="button" class="btn btn-secondary btn-sm" data-cancel>${tt('Cancel')}</button></div></div>`;
    host.appendChild(ov); const first = ov.querySelector('input:not([readonly]):not([disabled]),select'); if (first) first.focus(); else ov.querySelector('[data-ok]').focus();
    const close = () => ov.remove(), go = () => { const r = ok(ov); if (r === true) close(); else if (r) ov.querySelector('.rb-err').textContent = r; };
    ov.querySelector('[data-ok]').onclick = go; ov.querySelector('[data-cancel]').onclick = close; ov.querySelector('.rb-dx').onclick = close;
    ov.onkeydown = (e) => { if (e.key === 'Escape') close(); else if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); go(); } e.stopPropagation(); };
  }

  /* ── button actions ── */
  const SORT_MENU = () => [{ label: 'Sort A to Z', act: () => sortActive('asc') }, { label: 'Sort Z to A', act: () => sortActive('desc') }, { label: 'Custom Sort…', off: 1 }, { sep: 1 }, { label: 'Filter', check: afOn, act: toggleFilter }, { label: 'Clear', act: () => { Tools.clearFilter(sheet()); placeArrows(); changed(); } }, { label: 'Reapply', off: 1 }];
  const CF_MENU = () => {
    const hl = ['gt', 'lt', 'between', 'eq', 'text'].map((k) => ({ label: CFD[k].title + '…', act: () => cfDialog(k) }));
    hl.splice(4, 0, ...[]); hl.push({ label: 'A Date Occurring…', off: 1 }, { label: 'Duplicate Values…', act: () => cfDialog('dup') }, { sep: 1 }, { label: 'More Rules…', off: 1 });
    return [{ label: 'Highlight Cells Rules', sub: hl }, { label: 'Top/Bottom Rules', off: 1 }, { sep: 1 }, { label: 'Data Bars', off: 1 }, { label: 'Color Scales', off: 1 }, { label: 'Icon Sets', off: 1 }, { sep: 1 }, { label: 'New Rule…', off: 1 },
      { label: 'Clear Rules', sub: [{ label: 'Clear Rules from Selected Cells', act: () => { const b = bounds(grid.rangeStr()); sheet().cf = sheet().cf.filter((r) => { const q = bounds(r.range); return q.r2 < b.r1 || q.r1 > b.r2 || q.c2 < b.c1 || q.c1 > b.c2; }); changed(); } }, { label: 'Clear Rules from Entire Sheet', act: () => { sheet().cf = []; changed(); } }] }, { label: 'Manage Rules…', off: 1 }];
  };
  const MENUS = {
    'ch-col': () => M.colbar, 'ch-line': () => M.line, 'ch-pie': () => M.pie, cf: CF_MENU, sortfilter: SORT_MENU,
    addelem: addElemItems,
    changetype: () => ['column', 'bar', 'line', 'pie'].map((t) => ({ label: { column: 'Column', bar: 'Bar', line: 'Line', pie: 'Pie' }[t], check: chartDef().type === t, act: () => setProp('type', t) })),
  };
  const ACTS = {
    bold: () => grid.toggleBold(), currency: () => grid.applyFmt('currency'), percent: () => grid.applyFmt('percent'), comma: () => grid.applyFmt('comma'), autosum: () => grid.autoSum(),
    pivot: pivotDialog, 'sort-az': () => sortActive('asc'), 'sort-za': () => sortActive('desc'), filter: toggleFilter,
    clearfilter: () => { Tools.clearFilter(sheet()); placeArrows(); changed(); },
    switch: () => { const d = chartDef(); setProp('byRow', !d.byRow); },
    selectdata: () => { const d = chartDef(); openDialog({ title: L('Select Data Source'), body: `<div class="rb-drow"><span>${tt('Chart data range:')}</span><input data-r value="${d.range}" size="14" aria-label="range"></div>`, ok: (o) => {
      const v = o.querySelector('[data-r]').value.trim().toUpperCase().replace(/\$/g, ''); if (!/^[A-Z]+\d+:[A-Z]+\d+$/.test(v)) return tt('Enter a valid range such as A1:B7.'); d.range = v; renderChart(); changed(); return true; } }); },
  };
  /* ── general Home / View / Data buttons that work on the sheet ── */
  const PAL = [['No Fill', null, '#ffffff'], ['Yellow', 'yellow', '#fff2a8'], ['Green', 'green', '#c6efce'], ['Blue', 'blue', '#cfe2f7'], ['Orange', 'orange', '#fbd5a8'], ['Red', 'red', '#ffc7ce'], ['Grey', 'grey', '#e4e7e5']];
  const FPAL = [['Automatic', null, '#262626'], ['Red', 'red', '#c00000'], ['Blue', 'blue', '#1f5fbf'], ['Green', 'green', '#1e7a46']];
  const pasteNow = () => { if (grid.clip) grid.pasteData(grid.clip.text); else note('Copy something first (Ctrl + C), then open Paste Special.'); };
  Object.assign(MENUS, {
    fill: () => PAL.map(([l, v, hex]) => ({ label: l, sw: hex, act: () => grid.setStyle('fill', v || false) })),
    fontcolor: () => FPAL.map(([l, v, hex]) => ({ label: l, sw: hex, act: () => grid.setStyle('color', v || false) })),
    borders: () => [{ label: 'All Borders', ic: 'borders', act: () => grid.setStyle('border', 'all') }, { label: 'Bottom Border', ic: 'borders', act: () => grid.setStyle('border', 'bottom') }, { label: 'No Border', ic: 'borders', act: () => grid.setStyle('border', false) }],
    fmtcells: () => [{ label: 'AutoFit Column Width', act: () => grid.autofitSel() }, { label: 'Column Width…', off: 1 }, { label: 'Row Height…', off: 1 }, { sep: 1 },
      { label: 'Hide & Unhide', sub: [{ label: 'Hide Rows', act: () => grid.hide('rows') }, { label: 'Hide Columns', act: () => grid.hide('cols') }, { label: 'Unhide Rows', act: () => grid.unhide('rows') }, { label: 'Unhide Columns', act: () => grid.unhide('cols') }] }],
    inscells: () => [{ label: 'Insert Sheet Rows', act: () => grid.structure('insert', 'rows') }, { label: 'Insert Sheet Columns', act: () => grid.structure('insert', 'cols') }],
    delcells: () => [{ label: 'Delete Sheet Rows', act: () => grid.structure('delete', 'rows') }, { label: 'Delete Sheet Columns', act: () => grid.structure('delete', 'cols') }],
    find: () => [{ label: 'Find…', act: () => grid.openFind(false) }, { label: 'Replace…', act: () => grid.openFind(true) }, { sep: 1 }, { label: 'Go To…', act: () => grid.openGoTo() }, { label: 'Go To Special…', act: () => grid.openSpecial() }],
    paste: () => [{ label: 'Paste', act: pasteNow }, { label: 'Paste Special…', act: () => grid.openPasteSpecial() }],
    clearm: () => [{ label: 'Clear All', act: () => { grid.clearRange(); grid.clearFormats(); } }, { label: 'Clear Formats', act: () => grid.clearFormats() }, { label: 'Clear Contents', act: () => grid.clearRange() }],
    fillm: () => [{ label: 'Down', act: () => grid.fill('down') }, { label: 'Right', act: () => grid.fill('right') }, { sep: 1 }, { label: 'Flash Fill', act: () => grid.flashFill() }],
    freeze: () => [{ label: 'Freeze Panes', act: () => grid.setFreeze('both') }, { label: 'Freeze Top Row', act: () => grid.setFreeze('row') }, { label: 'Freeze First Column', act: () => grid.setFreeze('col') }, { label: 'Unfreeze Panes', act: () => grid.setFreeze('none') }],
  });
  Object.assign(ACTS, {
    refreshp: () => { rebuildPivot(); note('PivotTable refreshed.'); }, slicerp: () => slicerDialog(), groupsel: () => groupDialog(), groupfield: () => groupDialog(), ungroupsel: () => { pv.group = null; rebuildPivot(); },
    italic: () => grid.toggleStyle('italic'), underline: () => grid.toggleStyle('underline'), wrap: () => grid.toggleStyle('wrap'),
    alignl: () => grid.setStyle('align', 'left'), alignc: () => grid.setStyle('align', 'center'), alignr: () => grid.setStyle('align', 'right'),
    cut: () => grid.copyData(true), copy: () => grid.copyData(false), flashfill: () => grid.flashFill(), removedups: () => grid.openRemoveDuplicates(), validation: () => grid.openDataValidation(), definename: () => grid.openDefineName(), namemgr: () => grid.openNameManager(), texttocol: () => grid.openTextToColumns(),
  });
  const DIRECT_MENU = new Set(['ch-col', 'ch-line', 'ch-pie', 'cf', 'sortfilter', 'addelem', 'changetype']);

  /* ── rendering ── */
  const miniH = (it) => `<button type="button" class="rb-btn mini${enabled.has(it.id) ? '' : ' off'}" data-id="${it.id}" title="${tt(it.label)}" aria-label="${tt(it.label)}">${icon(it.ic)}</button>`;
  const itemH = (it) => {
    if (it.col) return `<div class="rb-col">${it.col.map(itemH).join('')}</div>`;
    if (it.rows) return `<div class="rb-rows">${it.rows.map((r) => `<div class="rb-row">${r.map(itemH).join('')}</div>`).join('')}</div>`;
    if (it.sel) return `<span class="rb-sel" style="width:${it.w}px">${it.sel}<i class="rb-dd"></i></span>`;
    if (it.mini) return miniH(it);
    if (it.styles) return '<div class="rb-styles"><i></i><i></i><i></i></div>';
    const on = enabled.has(it.id);
    return `<button type="button" class="rb-btn ${it.size}${on ? '' : ' off'}${it.label ? '' : ' icononly'}" data-id="${it.id}" title="${tt(it.tip || it.label)}"${on ? '' : ' aria-disabled="true"'}>${icon(it.ic)}${it.label ? `<span class="rb-l">${tt(it.label)}</span>` : ''}${it.menu ? '<i class="rb-dd"></i>' : ''}</button>`;
  };
  function draw() {
    if (!showRibbon) return;
    const tabs = tabsFor(); if (!tabs.includes(tab)) tab = 'home';
    host.innerHTML = `<div class="rb" role="tablist" aria-label="Ribbon"><div class="rb-tabs"><button type="button" class="rb-tab file" data-tab="file">${tt('File')}</button>${tabs.map((k) => `<button type="button" class="rb-tab${k === tab ? ' on' : ''}${k.startsWith('chart') || k === 'pivotanalyze' ? ' ctx' : ''}" role="tab" aria-selected="${k === tab}" data-tab="${k}">${tt(TAB_NAMES[k])}</button>`).join('')}<span class="rb-search"><svg viewBox="0 0 16 16" aria-hidden="true"><path d="M8 1a4.5 4.5 0 0 0-2.5 8.2V11h5V9.2A4.5 4.5 0 0 0 8 1zM6 13h4M6.5 15h3" stroke="currentColor" fill="none" stroke-linecap="round"/></svg><span>${tt('Tell me what you want to do')}</span></span><span class="rb-note" role="status"></span></div>
      <div class="rb-body">${TABS[tab].map((g) => `<div class="rb-group"><div class="rb-items">${g.items.map(itemH).join('')}</div><div class="rb-gname">${tt(g.name)}</div></div>`).join('')}</div></div>`;
    host.querySelector('.rb-tabs').onclick = (e) => { const b = e.target.closest('[data-tab]'); if (!b) return; if (b.dataset.tab === 'file') return note('This feature is not part of this exercise.'); tab = b.dataset.tab; closePop(); draw(); };
    host.querySelector('.rb-body').onclick = (e) => {
      const b = e.target.closest('[data-id]'); if (!b) return; const id = b.dataset.id;
      if (b.classList.contains('off')) return note('This feature is not part of this exercise.');
      if (MENUS[id] && !chartMissing(id)) { if (pop && pop.dataset.for === id) return closePop(); const r = b.getBoundingClientRect(); showMenu(MENUS[id](), r.left, r.bottom + 1, 0); if (pop) pop.dataset.for = id; return; }
      if (ACTS[id]) { ACTS[id](); if (!grid.dlgEl) grid.focus(); }
    };
    // bring the first exercise button of this tab into view when the ribbon is wider than the window
    const body = host.querySelector('.rb-body'), want = [...new Set([...(cfg.enable || []), ...(BY_KIND[kind] || [])])].filter((i) => !ALWAYS.has(i));
    for (const i of want) { const el = body.querySelector('[data-id="' + i + '"]'); if (el) { const x = el.getBoundingClientRect().left - body.getBoundingClientRect().left + body.scrollLeft; if (x + el.offsetWidth > body.clientWidth) body.scrollLeft = Math.max(0, x - 40); break; } }
    place_af();
  }
  const chartMissing = (id) => (id === 'addelem' || id === 'changetype') && !chartDef();
  const place_af = () => placeArrows();
  document.addEventListener('mousedown', function outside(e) { if (!host.isConnected) { document.removeEventListener('mousedown', outside); return; } if (pop && !pop.contains(e.target) && !e.target.closest('.rb-btn') && !e.target.closest('.pv-dd') && !e.target.closest('.af-btn')) closePop(); });
  window.addEventListener('scroll', () => { if (pop && host.isConnected) closePop(); }, { passive: true, capture: true });

  if (cfg.prebuilt) { Object.assign(pv, { created: true }, cfg.prebuilt); rebuildPivot(); }
  draw(); renderChart(); renderPivot(); if (afOn) placeArrows();
  return {
    onSelect() {},
    onChange() { if (kind === 'chart') renderChart(); if (kind === 'filter') placeArrows(); if (kind === 'pivot' && pv.created) renderPivot(); },
  };
}

root.Ribbon = { mount, TABS, LABELS, icon, IC };
['chart', 'cf', 'filter', 'pivot', 'ribbon'].forEach((k) => Tools.register(k, (cfg, host, ctx) => mount(Object.assign({}, cfg, { kind: k }), host, ctx)));
})(typeof window !== 'undefined' ? window : globalThis);
