/* Spreadsheet Academy curriculum. Three tracks, 13 lessons, ~80 hands-on steps.
   A step = { title, body, task, keys?, grid, check | quiz, hint, sol }
   Add a lesson by pushing to a track's `lessons` array: no other file needs to change. */
(function (root) {
'use strict';
const XL = root.XL;

/* ───────── authoring helpers ───────── */
const put = (start, rows) => {
  const m = XL.parseRef(start), o = {};
  rows.forEach((row, i) => row.forEach((v, j) => { if (v !== null && v !== undefined && v !== '') o[XL.numToCol(m.c + j) + (m.r + i)] = v; }));
  return o;
};
const merge = (...a) => Object.assign({}, ...a);
const G = (rows, cols, data, extra) => Object.assign({ rows, cols, data: data || {} }, extra);
const c = (s) => `<code>${s}</code>`;
const kbd = (s) => `<kbd>${s}</kbd>`;

/* a text answer is right in English or in the app language (the sheet words follow the language: Revenue / Receita) */
const same = (a, b) => (typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b)) : a === b || (typeof b === 'string' && typeof a === 'string' && !!root.I18N && root.I18N.word(b) === a));
const show = (v) => (v instanceof XL.XlError ? v.code : v === null ? 'empty' : typeof v === 'string' ? `"${v}"` : typeof v === 'boolean' ? String(v).toUpperCase() : v);

/** One cell must equal `exp`. o.formula: must be a formula. o.uses: RegExp the formula must match. */
function cell(ref, exp, o = {}) {
  return ({ sheet }) => {
    const raw = sheet.raw(ref);
    if (raw === '') return { ok: false, quiet: true, msg: `${ref} is still empty.` };
    const v = sheet.value(ref), isF = raw[0] === '=';
    if (v instanceof XL.XlError) return { ok: false, msg: `${ref} returned <b>${v.code}</b>. Check the spelling, commas and brackets in your formula.` };
    if (o.formula && !isF) return { ok: false, msg: same(v, exp) ? `Right result, but ${ref} should be a <b>formula</b> (start with <b>=</b>) so it updates by itself.` : `${ref} should contain a <b>formula</b> starting with =.` };
    if (!same(v, exp)) return { ok: false, msg: `${ref} shows <b>${show(v)}</b>, but it should show <b>${show(typeof exp === 'string' && root.I18N ? root.I18N.word(exp) : exp)}</b>.` };
    if (o.uses && !o.uses.test(XL.canon(raw))) return { ok: false, msg: o.usesMsg || `The result is right, but this exercise wants you to use a specific function. Re-read the task.` };
    return { ok: true };
  };
}
const cells = (map, o) => (ctx) => all(Object.keys(map).map((k) => cell(k, map[k], o)))(ctx);
function all(checks) {
  return (ctx) => {
    const rs = checks.map((f) => f(ctx));
    return rs.find((r) => !r.ok && r.msg && !r.quiet) || rs.find((r) => !r.ok) || { ok: true };
  };
}
const sel = (range) => ({ grid }) => grid.rangeStr() === range ? { ok: true } : { ok: false, quiet: true, msg: `You've selected <b>${grid.rangeStr()}</b>; the target is <b>${range}</b>.` };
const fmtOf = (refs, f, name) => ({ sheet }) => refs.every((r) => sheet.fmt(r) === f) ? { ok: true } : { ok: false, quiet: true, msg: `Not yet: every highlighted cell needs the <b>${name}</b> format.` };

const dataSales = put('A1', [['Rep', 'Region', 'Product', 'Sales'], ['Ana', 'West', 'Laptop', 1200], ['Ben', 'East', 'Phone', 800], ['Cy', 'West', 'Phone', 650], ['Dee', 'North', 'Laptop', 1500], ['Eli', 'East', 'Laptop', 1100], ['Fay', 'West', 'Tablet', 400], ['Gus', 'North', 'Phone', 700], ['Hal', 'East', 'Tablet', 450]]);
const dataSales2 = merge(dataSales, put('E1', [['Qtr'], ['Q1'], ['Q1'], ['Q2'], ['Q2'], ['Q1'], ['Q2'], ['Q1'], ['Q2']]));
const dataLookup = put('A1', [['ID', 'Product', 'Price'], [101, 'Pen', 1.5], [102, 'Notebook', 4.25], [103, 'Backpack', 39], [104, 'Bottle', 12.5], [105, 'Lamp', 27]]);

/* ═════════════════════════ TRACK 1 · FOUNDATIONS ═════════════════════════ */
const L1 = {
  id: 'grid-basics', title: 'Meet the grid', summary: 'Cells, addresses and entering data: the vocabulary of every spreadsheet.', minutes: 5, skills: ['Cells', 'Rows & columns', 'Name Box'],
  steps: [
    { title: 'Everything lives in a cell', body: `A spreadsheet is a grid of <b>cells</b>. Columns are labelled with letters (A, B, C…) and rows with numbers (1, 2, 3…). A cell's <b>address</b> combines both, like ${c('D4')}.`,
      task: `Click cell <b>D4</b>.`, grid: G(8, 7, {}, { cur: 'A1' }), check: sel('D4'), hint: `Find column <b>D</b> along the top and row <b>4</b> down the side. Click where they meet.`, sol: { sel: 'D4' } },
    { title: 'Typing text', body: `Select a cell and just start typing. Press <b>Enter</b> to confirm; the cursor drops to the next row. Text lines up on the left.`,
      task: `Type <b>Revenue</b> in cell <b>A1</b> and press <b>Enter</b>.`, grid: G(8, 7), check: cell('A1', 'Revenue'), hint: `A1 is already selected. Type the word and press Enter.`, sol: { cells: { A1: 'Revenue' } } },
    { title: 'Typing numbers', body: `Numbers line up on the <b>right</b>, so you can tell numbers from text at a glance. Press <b>Enter</b> after each entry and Excel moves down for you.`,
      task: `Enter <b>1200</b>, <b>950</b> and <b>1430</b> in <b>B2</b>, <b>B3</b> and <b>B4</b>.`, grid: G(8, 7, put('A1', [['Month', 'Sales'], ['Jan'], ['Feb'], ['Mar']]), { cur: 'B2' }),
      check: cells({ B2: 1200, B3: 950, B4: 1430 }), hint: `Type 1200, Enter, 950, Enter, 1430, Enter.`, sol: { cells: { B2: 1200, B3: 950, B4: 1430 } } },
    { title: 'Read an address', body: `The column letter always comes first, then the row number.`, task: `Which cell sits in <b>column C</b>, <b>row 5</b>?`,
      grid: G(7, 6, { C5: 'here' }, { cur: 'A1' }), locked: true, quiz: { options: [c('C5'), c('5C'), c('R5C3'), c('C-5')], answer: 0, explain: `Column first, row second: <b>C5</b>.` } },
    { title: 'Jump with the Name Box', body: `The <b>Name Box</b> (left of the formula bar) shows the active cell, and you can type an address into it to jump anywhere instantly. Shortcut: ${kbd('Ctrl')}+${kbd('G')}.`,
      task: `Click the Name Box, type <b>F6</b> and press <b>Enter</b>.`, grid: G(8, 7, {}, { cur: 'A1' }), check: sel('F6'), hint: `The Name Box is the small field in the top-left that currently says A1.`, sol: { sel: 'F6', text: 'Type F6 into the Name Box and press Enter.' } },
  ],
};

const L2 = {
  id: 'first-formulas', title: 'Your first formulas', summary: 'Operators, cell references and why formulas beat calculators.', minutes: 6, skills: ['=', '+ − × ÷', 'Cell references'],
  steps: [
    { title: 'Formulas start with =', body: `A cell can hold a <b>formula</b>. Excel shows the <b>result</b>, not the formula. Every formula begins with an equals sign ${c('=')}.`,
      task: `In <b>B2</b>, type ${c('=2+2')} and press <b>Enter</b>.`, grid: G(8, 7, {}, { cur: 'B2' }), check: cell('B2', 4, { formula: true }), hint: `Don't forget the = at the start.`, sol: { cells: { B2: '=2+2' } } },
    { title: 'Operators & parentheses', body: `Use ${c('+')} ${c('-')} ${c('*')} ${c('/')} and ${c('^')} for add, subtract, multiply, divide and power. Parentheses decide what is calculated first.`,
      task: `In <b>B2</b>, add <b>8</b> and <b>4</b> <i>first</i>, then multiply by <b>5</b>.`, grid: G(8, 7, {}, { cur: 'B2' }), check: cell('B2', 60, { formula: true, uses: /\(/, usesMsg: `Use parentheses so the addition happens first.` }),
      hint: `Without brackets, 8+4*5 would multiply first. Wrap the addition: ${c('=(8+4)*5')}.`, sol: { cells: { B2: '=(8+4)*5' } } },
    { title: 'Point at other cells', body: `Formulas get powerful when they use <b>cell references</b> instead of typed numbers. ${c('=B1*B2')} means "take whatever is in B1 and multiply by whatever is in B2".`,
      task: `In <b>B3</b>, multiply the price (B1) by the quantity (B2) using cell references.`, grid: G(8, 7, put('A1', [['Price', 45], ['Quantity', 6], ['Total']]), { cur: 'B3' }),
      check: cell('B3', 270, { formula: true, uses: /[A-Z]+\d/i, usesMsg: `Use the cell references B1 and B2 instead of typing the numbers.` }), hint: `${c('=B1*B2')}`, sol: { cells: { B3: '=B1*B2' } } },
    { title: 'Formulas stay alive', body: `Here is the payoff: when a referenced cell changes, every formula that depends on it <b>recalculates instantly</b>. Nothing to redo.`,
      task: `Change the price in <b>B1</b> to <b>50</b>. Watch <b>B3</b> update on its own.`, grid: G(8, 7, put('A1', [['Price', 45], ['Quantity', 6], ['Total', '=B1*B2']]), { cur: 'B1' }),
      check: all([cell('B1', 50), cell('B3', 300)]), hint: `Select B1, type 50 and press Enter.`, sol: { cells: { B1: 50 } } },
    { title: 'Order of operations', body: `Excel follows the maths rule you know: multiplication and division happen <b>before</b> addition and subtraction, unless parentheses say otherwise.`,
      task: `What does ${c('=2+3*4')} return?`, grid: G(5, 5, { B2: '=2+3*4' }, { cur: 'B2' }), locked: true, quiz: { options: ['20', '14', '24', '9'], answer: 1, explain: `3×4 happens first (12), then 2+12 = <b>14</b>.` } },
    { title: 'Click instead of typing', body: `While typing a formula you can <b>click a cell</b> and Excel inserts its address for you. It's faster and avoids typos.`,
      task: `In <b>B3</b> work out what's left of the budget: type <b>=</b>, click <b>B1</b>, type <b>-</b>, click <b>B2</b>, press <b>Enter</b>.`, grid: G(8, 7, put('A1', [['Budget', 5000], ['Spent', 3250], ['Left']]), { cur: 'B3' }),
      check: cell('B3', 1750, { formula: true, uses: /B1/i }), hint: `The result should be ${c('=B1-B2')}.`, sol: { cells: { B3: '=B1-B2' } } },
  ],
};

/* ═════════════════════════ REGISTRY ═════════════════════════
   Lessons live in lessons.js (core), lessons2-4.js (extensions). The final lessons file calls
   COURSE.finalize() to arrange them into learning paths. */
const REG = {};
[L1, L2].forEach((l) => { REG[l.id] = l; });

const COURSE = {
  TRACKS: [], LESSONS: [], REG,
  H: { put, merge, G, c, kbd, cell, cells, all, sel, fmtOf, same, show, dataSales, dataSales2, dataLookup },
  add(l) { REG[l.id] = l; },
  finalize(paths) {
    COURSE.TRACKS = paths.map((p) => ({ id: p.id, name: p.name, level: p.level, tone: p.tone, tagline: p.tagline, lessons: p.lessons.map((id) => { if (!REG[id]) throw new Error('Unknown lesson ' + id); return REG[id]; }) }));
    COURSE.TRACKS.forEach((t) => t.lessons.forEach((l, i) => { l.track = t.id; l.index = i + 1; }));
    COURSE.LESSONS = COURSE.TRACKS.flatMap((t) => t.lessons);
  },
};
root.COURSE = COURSE;
})(typeof window !== 'undefined' ? window : globalThis);
