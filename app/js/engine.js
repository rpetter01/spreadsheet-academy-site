/* Spreadsheet Academy formula engine: a small, dependency-free Excel-compatible calculator.
   Exposes window.XL = { Sheet, XlError, Rng, parseRef, colToNum, numToCol, shiftFormula, DOCS } */
(function (root) {
'use strict';

/* ───────────── basics ───────────── */
class XlError { constructor(code) { this.code = code; } toString() { return this.code; } }
const E = { DIV0: '#DIV/0!', VALUE: '#VALUE!', REF: '#REF!', NAME: '#NAME?', NA: '#N/A', NUM: '#NUM!', CIRC: '#CIRC!', ERR: '#ERROR!', CALC: '#CALC!', SPILL: '#SPILL!' };
const DAY0 = Date.UTC(1899, 11, 30);
const serialOf = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - DAY0) / 86400000);
const dateOf = (n) => new Date(DAY0 + Math.floor(n) * 86400000);
const todaySerial = () => { const t = new Date(); return serialOf(t.getFullYear(), t.getMonth() + 1, t.getDate()); };
const isDateRaw = (s) => /^\d{4}-\d{2}-\d{2}$/.test(s.trim()) || /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s.trim());
const err = (c) => new XlError(c);
class Rng {
  constructor(rows) { this.rows = rows; }
  get h() { return this.rows.length; }
  get w() { return this.rows.length ? this.rows[0].length : 0; }
  flat() { return this.rows.flat(); }
}
const colToNum = (s) => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
const numToCol = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
function parseRef(ref) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref);
  return m ? { c: colToNum(m[1]), r: +m[2] } : null;
}

/* ───────────── literals & formatting ───────────── */
function parseLiteral(raw) {
  if (raw === '') return null;
  if (raw[0] === "'") return raw.slice(1);
  const s = raw.trim();
  if (/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(s)) return parseFloat(s);
  if (/^[+-]?(\d+\.?\d*|\.\d+)%$/.test(s)) return parseFloat(s) / 100;
  if (/^\$\d[\d,]*(\.\d+)?$/.test(s)) return parseFloat(s.replace(/[$,]/g, ''));
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) return parseFloat(s.replace(/,/g, ''));
  if (/^[+-]?\d+,(\d{1,2}|\d{4,})$/.test(s)) return parseFloat(s.replace(',', '.'));
  let dm;
  if ((dm = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s))) return serialOf(+dm[1], +dm[2], +dm[3]);
  if ((dm = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s))) return serialOf(+dm[3], +dm[1], +dm[2]);
  if (/^true$/i.test(s)) return true;
  if (/^false$/i.test(s)) return false;
  return raw;
}
const numStr = (n) => {
  if (!isFinite(n)) return E.NUM;
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  return String(+n.toPrecision(12));
};
function formatValue(v, fmt) {
  if (v === null || v === undefined) return '';
  if (v instanceof XlError) return v.code;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return v;
  if (fmt === 'date') return dateOf(v).toISOString().slice(0, 10);
  if (fmt === 'currency') return (v < 0 ? '-' : '') + '$' + Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  if (fmt === 'comma') return v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  if (fmt === 'percent') return (v * 100).toFixed(0) + '%';
  if (!isFinite(v)) return E.NUM;
  if (Number.isInteger(v)) return String(v);
  return String(+v.toPrecision(10));
}

/* ───────────── tokenizer & parser ───────────── */
function tokenize(s) {
  const out = []; let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (/\s/.test(ch)) { i++; continue; }
    if (s.startsWith('#REF!', i)) throw err(E.REF);
    if (ch === '"') {
      let v = ''; i++;
      for (;;) {
        if (i >= s.length) throw err(E.ERR);
        if (s[i] === '"') { if (s[i + 1] === '"') { v += '"'; i += 2; continue; } i++; break; }
        v += s[i++];
      }
      out.push({ t: 'str', v }); continue;
    }
    const rest = s.slice(i); let m;
    if ((m = /^'((?:[^']|'')+)'!/.exec(rest))) { out.push({ t: 'sh', v: m[1].replace(/''/g, "'") }); i += m[0].length; continue; }
    if ((m = /^([A-Za-z_][A-Za-z0-9_.]*)!/.exec(rest))) { out.push({ t: 'sh', v: m[1] }); i += m[0].length; continue; }
    if ((m = /^(\d+\.?\d*|\.\d+)(e[+-]?\d+)?/i.exec(rest))) { out.push({ t: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue; }
    if ((m = /^[A-Za-z_$\u00C0-\u024F][A-Za-z0-9_.$\u00C0-\u024F]*/.exec(rest))) { out.push({ t: 'id', v: m[0] }); i += m[0].length; continue; }
    if ((m = /^(<>|<=|>=|[-+*\/^&=<>(),;:%])/.exec(rest))) { out.push({ t: 'op', v: m[0] === ';' ? ',' : m[0] }); i += m[0].length; continue; }
    throw err(E.ERR);
  }
  return out;
}

function parse(src) {
  // pt-BR / es style: ";" separators go with a decimal comma (1,5)
  if (/;/.test(src.replace(/"(?:[^"]|"")*"/g, ''))) src = src.split(/("(?:[^"]|"")*")/).map((p, i) => (i % 2 ? p : p.replace(/(\d),(?=\d)/g, '$1.'))).join('');
  const toks = tokenize(src); let p = 0;
  const isOp = (v) => toks[p] && toks[p].t === 'op' && toks[p].v === v;
  const expect = (v) => { if (!isOp(v)) throw err(E.ERR); p++; };
  const CMP = ['=', '<>', '<', '>', '<=', '>='];
  function cmp() { let l = cat(); while (toks[p] && toks[p].t === 'op' && CMP.includes(toks[p].v)) { const op = toks[p++].v; l = { t: 'bin', op, l, r: cat() }; } return l; }
  function cat() { let l = add(); while (isOp('&')) { p++; l = { t: 'bin', op: '&', l, r: add() }; } return l; }
  function add() { let l = mul(); while (isOp('+') || isOp('-')) { const op = toks[p++].v; l = { t: 'bin', op, l, r: mul() }; } return l; }
  function mul() { let l = pow(); while (isOp('*') || isOp('/')) { const op = toks[p++].v; l = { t: 'bin', op, l, r: pow() }; } return l; }
  function pow() { let l = unary(); while (isOp('^')) { p++; l = { t: 'bin', op: '^', l, r: unary() }; } return l; }
  function unary() { if (isOp('-')) { p++; return { t: 'un', e: unary() }; } if (isOp('+')) { p++; return unary(); } return post(); }
  function post() { let e = prim(); while (isOp('%')) { p++; e = { t: 'pct', e }; } return e; }
  function prim() {
    const t = toks[p++]; if (!t) throw err(E.ERR);
    if (t.t === 'num') return { t: 'num', v: t.v };
    if (t.t === 'str') return { t: 'str', v: t.v };
    if (t.t === 'op' && t.v === '(') { const e = cmp(); expect(')'); return e; }
    if (t.t === 'id') return ident(t.v);
    if (t.t === 'sh') { const n2 = toks[p++]; if (!n2 || n2.t !== 'id') throw err(E.ERR); return ident(n2.v, t.v); }
    throw err(E.ERR);
  }
  function ident(name, sh) {
    if (isOp('(')) {
      p++; const args = [];
      if (isOp(')')) { p++; return { t: 'fn', name: name.toUpperCase(), args }; }
      for (;;) {
        if (isOp(',') || isOp(')')) args.push({ t: 'empty' }); else args.push(cmp());
        if (isOp(',')) { p++; continue; }
        expect(')'); break;
      }
      return { t: 'fn', name: name.toUpperCase(), args };
    }
    const up = name.toUpperCase();
    if (up === 'TRUE' || up === 'FALSE' || up === 'VERDADEIRO' || up === 'VERDADERO' || up === 'FALSO') return { t: 'bool', v: up === 'TRUE' || up === 'VERDADEIRO' || up === 'VERDADERO' };
    const a = parseRef(name);
    if (a) {
      if (isOp(':')) {
        p++; const t2 = toks[p++]; const b = t2 && t2.t === 'id' ? parseRef(t2.v) : null;
        if (!b) throw err(E.ERR);
        return { t: 'range', a, b, sh };
      }
      return { t: 'ref', c: a.c, r: a.r, sh };
    }
    if (/^\$?[A-Za-z]{1,3}$/.test(name) && isOp(':')) {
      p++; const t2 = toks[p++];
      if (!t2 || t2.t !== 'id' || !/^\$?[A-Za-z]{1,3}$/.test(t2.v)) throw err(E.ERR);
      return { t: 'range', a: { c: colToNum(name.replace('$', '')), r: 1 }, b: { c: colToNum(t2.v.replace('$', '')), r: Infinity }, sh };
    }
    return { t: 'name', v: name };
  }
  const ast = cmp();
  if (p < toks.length) throw err(E.ERR);
  return ast;
}
const AST = new Map();

/* ───────────── coercion ───────────── */
const toNum = (v) => {
  if (v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof XlError) throw v;
  if (v instanceof Rng) return toNum(v.rows[0][0]);
  const s = String(v).trim();
  if (s !== '' && !isNaN(Number(s))) return Number(s);
  throw err(E.VALUE);
};
const toStr = (v) => {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return numStr(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof XlError) throw v;
  if (v instanceof Rng) return toStr(v.rows[0][0]);
  return String(v);
};
const toBool = (v) => {
  if (v instanceof Rng) v = v.rows[0][0];
  if (v === null || v === undefined) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v instanceof XlError) throw v;
  const s = String(v).toUpperCase();
  if (s === 'TRUE') return true; if (s === 'FALSE') return false;
  throw err(E.VALUE);
};
const scalar = (v) => (v instanceof Rng ? v.rows[0][0] : v);
const safe = (f) => { try { return f(); } catch (e) { if (e instanceof XlError) return e; throw e; } };

/* ───────────── evaluation ───────────── */
function compare(op, a, b) {
  if (a instanceof XlError) throw a; if (b instanceof XlError) throw b;
  const dflt = (o) => (typeof o === 'string' ? '' : typeof o === 'boolean' ? false : 0);
  if (a === null) a = dflt(b); if (b === null) b = dflt(a);
  const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  let r;
  if (rank(a) !== rank(b)) r = rank(a) - rank(b);
  else if (typeof a === 'string') { const x = a.toLowerCase(), y = b.toLowerCase(); r = x < y ? -1 : x > y ? 1 : 0; }
  else if (typeof a === 'number' && Math.abs(a - b) <= 1e-12 * Math.max(1, Math.abs(a), Math.abs(b))) r = 0;
  else r = a < b ? -1 : a > b ? 1 : 0;
  switch (op) { case '=': return r === 0; case '<>': return r !== 0; case '<': return r < 0; case '>': return r > 0; case '<=': return r <= 0; default: return r >= 0; }
}
function binop(op, a, b) {
  switch (op) {
    case '+': return toNum(a) + toNum(b);
    case '-': return toNum(a) - toNum(b);
    case '*': return toNum(a) * toNum(b);
    case '/': { const d = toNum(b); if (d === 0) throw err(E.DIV0); return toNum(a) / d; }
    case '^': { const r = Math.pow(toNum(a), toNum(b)); if (!isFinite(r)) throw err(E.NUM); return r; }
    case '&': return toStr(a) + toStr(b);
    default: return compare(op, a, b);
  }
}
const pick = (M, i, j) => M[M.length === 1 ? 0 : i]?.[M[0].length === 1 ? 0 : j];
function lift1(x, f) { return x instanceof Rng ? new Rng(x.rows.map((r) => r.map((v) => safe(() => f(v))))) : f(x); }
function lift2(a, b, f) {
  if (!(a instanceof Rng) && !(b instanceof Rng)) return f(a, b);
  const A = a instanceof Rng ? a.rows : [[a]], B = b instanceof Rng ? b.rows : [[b]];
  const h = Math.max(A.length, B.length), w = Math.max(A[0].length, B[0].length), rows = [];
  for (let i = 0; i < h; i++) {
    const row = [];
    for (let j = 0; j < w; j++) { const x = pick(A, i, j), y = pick(B, i, j); row.push(x === undefined || y === undefined ? err(E.NA) : safe(() => f(x, y))); }
    rows.push(row);
  }
  return new Rng(rows);
}
function rangeOf(n, cx) {
  const r1 = Math.min(n.a.r, n.b.r), r2 = Math.max(n.a.r, n.b.r), c1 = Math.min(n.a.c, n.b.c), c2 = Math.max(n.a.c, n.b.c);
  const S = n.sh ? cx.sheet.other(n.sh) : cx.sheet;
  const maxR = Math.min(r2, Math.max(S.maxRow(), r1)), rows = [];
  for (let r = r1; r <= maxR; r++) { const row = []; for (let c = c1; c <= c2; c++) row.push(S.valueAt(c, r)); rows.push(row); }
  return new Rng(rows);
}
function evalRefString(str, cx) {
  const m = /^(?:(.+)!)?(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?$/.exec(str.trim()); if (!m) throw err(E.REF);
  const sh = m[1] ? m[1].replace(/^'|'$/g, '') : undefined, a = parseRef(m[2]), b = m[3] ? parseRef(m[3]) : null;
  if (!b) { const v = cx.sheet.valueAt(a.c, a.r, sh); if (v instanceof XlError) throw v; return v; }
  return rangeOf({ a, b, sh }, cx);
}
function ev(n, cx) {
  switch (n.t) {
    case 'num': case 'str': case 'bool': return n.v;
    case 'empty': return null;
    case 'ref': { const v = cx.sheet.valueAt(n.c, n.r, n.sh); if (v instanceof XlError) throw v; return v; }
    case 'range': return rangeOf(n, cx);
    case 'name': { const d = cx.sheet.nameRef(n.v); if (!d) throw err(E.NAME); return evalRefString(d, cx); }
    case 'un': return lift1(ev(n.e, cx), (x) => -toNum(x));
    case 'pct': return lift1(ev(n.e, cx), (x) => toNum(x) / 100);
    case 'bin': return lift2(ev(n.l, cx), ev(n.r, cx), (a, b) => binop(n.op, a, b));
    case 'fn': {
      const f = FN[n.name] || FN[ALIAS[nrm(n.name)]]; if (!f) throw err(E.NAME);
      return f.lazy ? f.fn(n.args, cx) : f.fn(n.args.map((a) => ev(a, cx)), cx);
    }
  }
  throw err(E.ERR);
}

/* ───────────── function library ───────────── */
const FN = {};
const def = (names, fn, lazy) => { for (const n of names.split(' ')) FN[n] = { fn, lazy: !!lazy }; };
const rngOf = (x) => (x instanceof Rng ? x : new Rng([[x]]));
const need = (a, n) => { if (a.length < n) throw err(E.VALUE); };
function nums(args) {
  const out = [];
  for (const a of args) {
    if (a instanceof Rng) { for (const v of a.flat()) { if (v instanceof XlError) throw v; if (typeof v === 'number') out.push(v); } }
    else if (a !== null) out.push(toNum(a));
  }
  return out;
}
const sum = (a) => a.reduce((x, y) => x + y, 0);
const roundTo = (x, d, mode) => {
  const k = Math.pow(10, d), s = Math.sign(x), v = Math.abs(x) * k;
  const adj = mode === 'up' ? Math.ceil(+(v.toPrecision(14))) : mode === 'down' ? Math.floor(+(v.toPrecision(14))) : Math.round(+(v.toPrecision(14)));
  return (s * adj) / k;
};
const wild = (pat) => new RegExp('^' + pat.replace(/~([*?])|([.+^${}()|[\]\\])|(\*)|(\?)/g, (m, esc, sp, st, q) => esc ? '\\' + esc : sp ? '\\' + sp : st ? '.*' : '.') + '$', 'i');
function crit(c) {
  c = scalar(c);
  if (typeof c === 'number') return (v) => typeof v === 'number' && Math.abs(v - c) < 1e-12;
  if (typeof c === 'boolean') return (v) => v === c;
  c = c === null ? '' : String(c);
  const m = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/.exec(c), op = m[1] || '=', rest = m[2];
  const n = rest.trim() !== '' && !isNaN(Number(rest)) ? Number(rest) : null;
  if (n !== null) {
    return (v) => {
      if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) v = Number(v);
      if (typeof v !== 'number') return op === '<>';
      return op === '=' ? v === n : op === '<>' ? v !== n : op === '<' ? v < n : op === '>' ? v > n : op === '<=' ? v <= n : v >= n;
    };
  }
  if (rest === '') return op === '<>' ? (v) => v !== null && v !== '' : (v) => v === null || v === '';
  if (op === '=' || op === '<>') {
    const re = wild(rest);
    return (v) => { const hit = (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') && re.test(toStr(v)); return op === '=' ? hit : !hit; };
  }
  const L = rest.toLowerCase();
  return (v) => { if (typeof v !== 'string') return false; const x = v.toLowerCase(); return op === '<' ? x < L : op === '>' ? x > L : op === '<=' ? x <= L : x >= L; };
}
function idxWhere(pairs) {
  const n = pairs[0][0].length, out = [];
  for (let i = 0; i < n; i++) if (pairs.every(([arr, f]) => i < arr.length && f(arr[i]))) out.push(i);
  return out;
}
const numAt = (arr, i) => (typeof arr[i] === 'number' ? arr[i] : (arr[i] instanceof XlError ? (() => { throw arr[i]; })() : null));

// math & stats
def('SUM', (a) => sum(nums(a)));
def('AVERAGE', (a) => { const n = nums(a); if (!n.length) throw err(E.DIV0); return sum(n) / n.length; });
def('MIN', (a) => { const n = nums(a); return n.length ? Math.min(...n) : 0; });
def('MAX', (a) => { const n = nums(a); return n.length ? Math.max(...n) : 0; });
def('MEDIAN', (a) => { const n = nums(a).sort((x, y) => x - y); if (!n.length) throw err(E.NUM); const m = n.length >> 1; return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2; });
def('PRODUCT', (a) => nums(a).reduce((x, y) => x * y, 1));
def('COUNT', (a) => { let c = 0; for (const x of a) { if (x instanceof Rng) c += x.flat().filter((v) => typeof v === 'number').length; else if (x !== null) { try { toNum(x); c++; } catch (e) { /* not a number */ } } } return c; });
def('COUNTA', (a) => a.reduce((c, x) => c + (x instanceof Rng ? x.flat().filter((v) => v !== null).length : x !== null ? 1 : 0), 0));
def('COUNTBLANK', (a) => a.reduce((c, x) => c + rngOf(x).flat().filter((v) => v === null || v === '').length, 0));
def('LARGE', (a) => { need(a, 2); const n = nums([a[0]]).sort((x, y) => y - x), k = toNum(a[1]); if (k < 1 || k > n.length) throw err(E.NUM); return n[k - 1]; });
def('SMALL', (a) => { need(a, 2); const n = nums([a[0]]).sort((x, y) => x - y), k = toNum(a[1]); if (k < 1 || k > n.length) throw err(E.NUM); return n[k - 1]; });
def('RANK', (a) => { need(a, 2); const v = toNum(a[0]), n = nums([a[1]]), asc = a.length > 2 && toNum(a[2]) !== 0; if (!n.includes(v)) throw err(E.NA); return 1 + n.filter((x) => (asc ? x < v : x > v)).length; });
def('ROUND', (a) => { need(a, 1); return roundTo(toNum(a[0]), a.length > 1 ? toNum(a[1]) : 0, 'near'); });
def('ROUNDUP', (a) => { need(a, 1); return roundTo(toNum(a[0]), a.length > 1 ? toNum(a[1]) : 0, 'up'); });
def('ROUNDDOWN', (a) => { need(a, 1); return roundTo(toNum(a[0]), a.length > 1 ? toNum(a[1]) : 0, 'down'); });
def('INT', (a) => Math.floor(toNum(a[0])));
def('ABS', (a) => Math.abs(toNum(a[0])));
def('SQRT', (a) => { const x = toNum(a[0]); if (x < 0) throw err(E.NUM); return Math.sqrt(x); });
def('POWER', (a) => binop('^', a[0], a[1]));
def('MOD', (a) => { need(a, 2); const n = toNum(a[0]), d = toNum(a[1]); if (d === 0) throw err(E.DIV0); return n - d * Math.floor(n / d); });
def('SUMPRODUCT', (a) => {
  const R = a.map(rngOf); if (!R.length) throw err(E.VALUE);
  const h = R[0].h, w = R[0].w; if (R.some((r) => r.h !== h || r.w !== w)) throw err(E.VALUE);
  let t = 0;
  for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) {
    let p = 1;
    for (const r of R) { const v = r.rows[i][j]; if (v instanceof XlError) throw v; p *= typeof v === 'number' ? v : 0; }
    t += p;
  }
  return t;
});

// conditional aggregation
def('COUNTIF', (a) => { need(a, 2); const f = crit(a[1]); return rngOf(a[0]).flat().filter(f).length; });
def('SUMIF', (a) => {
  need(a, 2); const r = rngOf(a[0]).flat(), f = crit(a[1]), s = a.length > 2 ? rngOf(a[2]).flat() : r;
  return sum(r.map((v, i) => (f(v) && typeof s[i] === 'number' ? s[i] : 0)));
});
def('AVERAGEIF', (a) => {
  need(a, 2); const r = rngOf(a[0]).flat(), f = crit(a[1]), s = a.length > 2 ? rngOf(a[2]).flat() : r;
  const hits = r.map((v, i) => (f(v) && typeof s[i] === 'number' ? s[i] : null)).filter((x) => x !== null);
  if (!hits.length) throw err(E.DIV0); return sum(hits) / hits.length;
});
const pairsFrom = (a, from) => { const p = []; for (let i = from; i + 1 < a.length; i += 2) p.push([rngOf(a[i]).flat(), crit(a[i + 1])]); if (!p.length) throw err(E.VALUE); return p; };
def('COUNTIFS', (a) => idxWhere(pairsFrom(a, 0)).length);
def('SUMIFS', (a) => { need(a, 3); const s = rngOf(a[0]).flat(); return sum(idxWhere(pairsFrom(a, 1)).map((i) => numAt(s, i) || 0)); });
def('AVERAGEIFS', (a) => { need(a, 3); const s = rngOf(a[0]).flat(); const v = idxWhere(pairsFrom(a, 1)).map((i) => numAt(s, i)).filter((x) => x !== null); if (!v.length) throw err(E.DIV0); return sum(v) / v.length; });
def('MAXIFS', (a) => { need(a, 3); const s = rngOf(a[0]).flat(); const v = idxWhere(pairsFrom(a, 1)).map((i) => numAt(s, i)).filter((x) => x !== null); return v.length ? Math.max(...v) : 0; });
def('MINIFS', (a) => { need(a, 3); const s = rngOf(a[0]).flat(); const v = idxWhere(pairsFrom(a, 1)).map((i) => numAt(s, i)).filter((x) => x !== null); return v.length ? Math.min(...v) : 0; });

// logic
def('IF', (n, cx) => {
  if (n.length < 1) throw err(E.VALUE);
  if (n.length === 1) return toBool(ev(n[0], cx));   /* as in Excel: IF(test) alone gives TRUE or FALSE */
  if (toBool(ev(n[0], cx))) return n[1].t === 'empty' ? 0 : ev(n[1], cx);
  if (n.length < 3) return false;
  return n[2].t === 'empty' ? 0 : ev(n[2], cx);
}, true);
def('IFS', (n, cx) => { for (let i = 0; i + 1 < n.length; i += 2) if (toBool(ev(n[i], cx))) return ev(n[i + 1], cx); throw err(E.NA); }, true);
def('IFERROR', (n, cx) => { if (n.length < 2) throw err(E.VALUE); try { const r = ev(n[0], cx); return r instanceof XlError ? ev(n[1], cx) : r; } catch (e) { if (e instanceof XlError) return ev(n[1], cx); throw e; } }, true);
def('IFNA', (n, cx) => { try { const r = ev(n[0], cx); return r instanceof XlError && r.code === E.NA ? ev(n[1], cx) : r; } catch (e) { if (e instanceof XlError && e.code === E.NA) return ev(n[1], cx); throw e; } }, true);
const bools = (a) => { const o = []; for (const x of a) { if (x instanceof Rng) x.flat().forEach((v) => { if (v instanceof XlError) throw v; if (typeof v === 'boolean' || typeof v === 'number') o.push(!!v); }); else o.push(toBool(x)); } if (!o.length) throw err(E.VALUE); return o; };
def('AND', (a) => bools(a).every(Boolean));
def('OR', (a) => bools(a).some(Boolean));
def('NOT', (a) => !toBool(a[0]));
def('TRUE', () => true); def('FALSE', () => false);
def('ISBLANK', (a) => a[0] === null);
def('ISNUMBER', (a) => typeof scalar(a[0]) === 'number');
def('ISTEXT', (a) => typeof scalar(a[0]) === 'string');
def('ISERROR', (n, cx) => { try { return ev(n[0], cx) instanceof XlError; } catch (e) { if (e instanceof XlError) return true; throw e; } }, true);
def('ISNA', (n, cx) => { try { const r = ev(n[0], cx); return r instanceof XlError && r.code === E.NA; } catch (e) { if (e instanceof XlError) return e.code === E.NA; throw e; } }, true);

// text
def('LEN', (a) => toStr(a[0]).length);
def('LEFT', (a) => toStr(a[0]).slice(0, a.length > 1 ? Math.max(0, toNum(a[1])) : 1));
def('RIGHT', (a) => { const s = toStr(a[0]), n = a.length > 1 ? Math.max(0, toNum(a[1])) : 1; return n ? s.slice(-n) : ''; });
def('MID', (a) => { need(a, 3); const st = toNum(a[1]), n = toNum(a[2]); if (st < 1 || n < 0) throw err(E.VALUE); return toStr(a[0]).substr(st - 1, n); });
def('UPPER', (a) => toStr(a[0]).toUpperCase());
def('LOWER', (a) => toStr(a[0]).toLowerCase());
def('PROPER', (a) => toStr(a[0]).toLowerCase().replace(/(^|[^a-z0-9'])([a-z])/g, (m, p, c) => p + c.toUpperCase()));
def('TRIM', (a) => toStr(a[0]).trim().replace(/ {2,}/g, ' '));
def('EXACT', (a) => toStr(a[0]) === toStr(a[1]));
def('REPT', (a) => toStr(a[0]).repeat(Math.max(0, toNum(a[1]))));
def('VALUE', (a) => { const s = scalar(a[0]); if (typeof s === 'number') return s; const v = parseLiteral(toStr(s)); if (typeof v === 'number') return v; throw err(E.VALUE); });
def('CONCAT CONCATENATE', (a) => a.map((x) => (x instanceof Rng ? x.flat().map(toStr).join('') : toStr(x))).join(''));
def('TEXTJOIN', (a) => {
  need(a, 3); const d = toStr(a[0]), skip = toBool(a[1]), parts = [];
  for (const x of a.slice(2)) (x instanceof Rng ? x.flat() : [x]).forEach((v) => { const s = toStr(v); if (!(skip && s === '')) parts.push(s); });
  return parts.join(d);
});
def('FIND', (a) => { need(a, 2); const i = toStr(a[1]).indexOf(toStr(a[0]), a.length > 2 ? toNum(a[2]) - 1 : 0); if (i < 0) throw err(E.VALUE); return i + 1; });
def('SEARCH', (a) => { need(a, 2); const i = toStr(a[1]).toLowerCase().indexOf(toStr(a[0]).toLowerCase(), a.length > 2 ? toNum(a[2]) - 1 : 0); if (i < 0) throw err(E.VALUE); return i + 1; });
def('SUBSTITUTE', (a) => {
  need(a, 3); const s = toStr(a[0]), o = toStr(a[1]), n = toStr(a[2]); if (o === '') return s;
  if (a.length > 3) { let k = toNum(a[3]), pos = -1; while (k-- > 0) { pos = s.indexOf(o, pos + 1); if (pos < 0) return s; } return s.slice(0, pos) + n + s.slice(pos + o.length); }
  return s.split(o).join(n);
});
def('TEXT', (a) => {
  need(a, 2); const v = scalar(a[0]), f = toStr(a[1]);
  if (typeof v !== 'number') return toStr(v);
  const pct = f.includes('%'); let n = pct ? v * 100 : v;
  const m = /^([^0#.,]*)([0#,]*)(\.[0#]*)?([\s\S]*)$/.exec(f.replace('%', ''));
  const dec = m[3] ? m[3].length - 1 : 0, minInt = (m[2].match(/0/g) || []).length;
  let s = Math.abs(roundTo(n, dec, 'near')).toFixed(dec), [i, d] = s.split('.');
  i = i.padStart(minInt, '0'); if (i === '0' && minInt === 0 && dec > 0) i = '';
  if (m[2].includes(',')) i = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (n < 0 ? '-' : '') + m[1] + i + (d ? '.' + d : '') + m[4] + (pct ? '%' : '');
});

// lookup
const lc = (v) => (typeof v === 'string' ? v.toLowerCase() : v);
const eqLoose = (a, b) => (typeof a === 'string' && typeof b === 'string' ? a.toLowerCase() === b.toLowerCase() : a === b || (typeof a === 'number' && typeof b === 'number' && Math.abs(a - b) < 1e-12));
function findExact(arr, key, fromEnd) {
  const test = typeof key === 'string' && /[*?~]/.test(key) ? ((re) => (v) => typeof v === 'string' && re.test(v))(wild(key)) : (v) => v !== null && eqLoose(v, key);
  if (fromEnd) { for (let i = arr.length - 1; i >= 0; i--) if (test(arr[i])) return i; return -1; }
  return arr.findIndex(test);
}
function findApprox(arr, key, type) {
  let best = -1; const k = lc(key);
  arr.forEach((v, i) => {
    if (v === null || typeof v !== typeof key) return; const x = lc(v);
    if (type === 1 ? x <= k : x >= k) { if (best < 0 || (type === 1 ? x >= lc(arr[best]) : x <= lc(arr[best]))) best = i; }
  });
  return best;
}
def('VLOOKUP', (a) => {
  need(a, 3); const key = scalar(a[0]), t = rngOf(a[1]), col = toNum(a[2]), approx = a.length > 3 ? toBool(a[3]) : true;
  if (col < 1 || col > t.w) throw err(E.REF);
  const first = t.rows.map((r) => r[0]), i = approx ? findApprox(first, key, 1) : findExact(first, key);
  if (i < 0) throw err(E.NA); const v = t.rows[i][col - 1]; if (v instanceof XlError) throw v; return v;
});
def('HLOOKUP', (a) => {
  need(a, 3); const key = scalar(a[0]), t = rngOf(a[1]), row = toNum(a[2]), approx = a.length > 3 ? toBool(a[3]) : true;
  if (row < 1 || row > t.h) throw err(E.REF);
  const i = approx ? findApprox(t.rows[0], key, 1) : findExact(t.rows[0], key);
  if (i < 0) throw err(E.NA); return t.rows[row - 1][i];
});
def('MATCH', (a) => {
  need(a, 2); const key = scalar(a[0]), arr = rngOf(a[1]).flat(), type = a.length > 2 ? toNum(a[2]) : 1;
  const i = type === 0 ? findExact(arr, key) : findApprox(arr, key, type > 0 ? 1 : -1);
  if (i < 0) throw err(E.NA); return i + 1;
});
def('INDEX', (a) => {
  need(a, 2); const t = rngOf(a[0]); let r = toNum(a[1]), c = a.length > 2 ? toNum(a[2]) : 1;
  if (a.length < 3 && t.h === 1) { c = r; r = 1; }
  if (r < 1 || c < 1 || r > t.h || c > t.w) throw err(E.REF);
  const v = t.rows[r - 1][c - 1]; if (v instanceof XlError) throw v; return v;
});
def('XLOOKUP', (a) => {
  need(a, 3); const key = scalar(a[0]), look = rngOf(a[1]), ret = rngOf(a[2]), mm = a.length > 4 && a[4] !== null ? toNum(a[4]) : 0, sm = a.length > 5 ? toNum(a[5]) : 1;
  const arr = look.flat(); let i = findExact(arr, key, sm === -1);
  if (i < 0 && mm === -1) i = findApprox(arr, key, 1);
  if (i < 0 && mm === 1) i = findApprox(arr, key, -1);
  if (i < 0) { if (a.length > 3 && a[3] !== null) return a[3]; throw err(E.NA); }
  const v = look.w === 1 ? (ret.w === 1 ? ret.rows[i][0] : new Rng([ret.rows[i]])) : (ret.h === 1 ? ret.rows[0][i] : new Rng(ret.rows.map((r) => [r[i]])));
  if (v instanceof XlError) throw v; return v;
});
def('ROWS', (a) => rngOf(a[0]).h);
def('COLUMNS', (a) => rngOf(a[0]).w);

// logic extras
def('XOR', (a) => bools(a).filter(Boolean).length % 2 === 1);
def('SWITCH', (n, cx) => {
  if (n.length < 3) throw err(E.VALUE);
  const v = scalar(ev(n[0], cx)); let i = 1;
  for (; i + 1 < n.length; i += 2) if (compare('=', v, scalar(ev(n[i], cx)))) return ev(n[i + 1], cx);
  if (i < n.length) return ev(n[i], cx);
  throw err(E.NA);
}, true);
def('CHOOSE', (a) => { need(a, 2); const i = toNum(a[0]); if (i < 1 || i >= a.length) throw err(E.VALUE); return a[i]; });
def('ISLOGICAL', (a) => typeof scalar(a[0]) === 'boolean');
def('ISEVEN', (a) => Math.floor(toNum(a[0])) % 2 === 0);
def('ISODD', (a) => Math.abs(Math.floor(toNum(a[0]))) % 2 === 1);
def('NA', () => { throw err(E.NA); });

// rounding & stats extras
const clean = (x) => +x.toPrecision(12);
def('FLOOR', (a) => { need(a, 2); const n = toNum(a[0]), s = toNum(a[1]); if (s === 0) return 0; if (n > 0 && s < 0) throw err(E.NUM); return clean(Math.floor(clean(n / s)) * s); });
def('CEILING', (a) => { need(a, 2); const n = toNum(a[0]), s = toNum(a[1]); if (s === 0) return 0; if (n > 0 && s < 0) throw err(E.NUM); return clean(Math.ceil(clean(n / s)) * s); });
def('TRUNC', (a) => roundTo(toNum(a[0]), a.length > 1 ? toNum(a[1]) : 0, 'down'));
def('ODD', (a) => { const n = toNum(a[0]); let v = Math.ceil(Math.abs(n)); if (v % 2 === 0) v += 1; return (Math.sign(n) || 1) * v; });
def('EVEN', (a) => { const n = toNum(a[0]); let v = Math.ceil(Math.abs(n)); if (v % 2 === 1) v += 1; return (Math.sign(n) || 1) * v; });
def('MODE MODE.SNGL', (a) => {
  const cnt = new Map(); nums(a).forEach((x) => cnt.set(x, (cnt.get(x) || 0) + 1));
  let best = null, bc = 1; for (const [k, v] of cnt) if (v > bc) { best = k; bc = v; }
  if (best === null) throw err(E.NA); return best;
});

// text extras
def('TEXTBEFORE', (a) => { need(a, 2); const s = toStr(a[0]), d = toStr(a[1]); let k = a.length > 2 ? toNum(a[2]) : 1, pos = -1; while (k-- > 0) { pos = s.indexOf(d, pos + 1); if (pos < 0) throw err(E.NA); } return s.slice(0, pos); });
def('TEXTAFTER', (a) => { need(a, 2); const s = toStr(a[0]), d = toStr(a[1]); let k = a.length > 2 ? toNum(a[2]) : 1, pos = -1; while (k-- > 0) { pos = s.indexOf(d, pos + 1); if (pos < 0) throw err(E.NA); } return s.slice(pos + d.length); });

// dates
def('DATE', (a) => { need(a, 3); return serialOf(toNum(a[0]), toNum(a[1]), toNum(a[2])); });
def('TODAY NOW', () => todaySerial());
def('YEAR', (a) => dateOf(toNum(a[0])).getUTCFullYear());
def('MONTH', (a) => dateOf(toNum(a[0])).getUTCMonth() + 1);
def('DAY', (a) => dateOf(toNum(a[0])).getUTCDate());
def('WEEKDAY', (a) => { const w = dateOf(toNum(a[0])).getUTCDay(), t = a.length > 1 ? toNum(a[1]) : 1; return t === 2 ? (w === 0 ? 7 : w) : t === 3 ? (w + 6) % 7 : w + 1; });
const isWork = (d, hol) => { const w = dateOf(d).getUTCDay(); return w !== 0 && w !== 6 && !hol.has(d); };
const holidays = (a) => new Set(a.length ? nums(a).map(Math.floor) : []);
def('NETWORKDAYS', (a) => { need(a, 2); let s0 = Math.floor(toNum(a[0])), e0 = Math.floor(toNum(a[1])); const hol = holidays(a.slice(2)), sign = s0 <= e0 ? 1 : -1; if (sign < 0) [s0, e0] = [e0, s0]; let n = 0; for (let d = s0; d <= e0; d++) if (isWork(d, hol)) n++; return sign * n; });
def('WORKDAY', (a) => { need(a, 2); let d = Math.floor(toNum(a[0])), n = Math.trunc(toNum(a[1])); const hol = holidays(a.slice(2)), step = n < 0 ? -1 : 1; while (n !== 0) { d += step; if (isWork(d, hol)) n -= step; } return d; });
def('PMT', (a) => { need(a, 3); const r = toNum(a[0]), n = toNum(a[1]), pv = toNum(a[2]), fv = a.length > 3 && a[3] !== null ? toNum(a[3]) : 0, t = a.length > 4 && a[4] !== null ? toNum(a[4]) : 0; if (n === 0) throw err(E.NUM); if (r === 0) return -(pv + fv) / n; const k = Math.pow(1 + r, n); return -(r * (pv * k + fv)) / ((1 + r * t) * (k - 1)); });
def('FV', (a) => { need(a, 3); const r = toNum(a[0]), n = toNum(a[1]), p = toNum(a[2]), pv = a.length > 3 && a[3] !== null ? toNum(a[3]) : 0, t = a.length > 4 && a[4] !== null ? toNum(a[4]) : 0; if (r === 0) return -(pv + p * n); const k = Math.pow(1 + r, n); return -(pv * k + p * (1 + r * t) * (k - 1) / r); });
def('PV', (a) => { need(a, 3); const r = toNum(a[0]), n = toNum(a[1]), p = toNum(a[2]), fv = a.length > 3 && a[3] !== null ? toNum(a[3]) : 0, t = a.length > 4 && a[4] !== null ? toNum(a[4]) : 0; if (r === 0) return -(fv + p * n); const k = Math.pow(1 + r, n); return -(fv + p * (1 + r * t) * (k - 1) / r) / k; });
def('NPV', (a) => { need(a, 2); const r = toNum(a[0]), v = nums(a.slice(1)); if (r <= -1) throw err(E.NUM); return v.reduce((t, x, i) => t + x / Math.pow(1 + r, i + 1), 0); });
def('IRR', (a) => {
  need(a, 1); const v = nums([a[0]]); if (!v.some((x) => x > 0) || !v.some((x) => x < 0)) throw err(E.NUM);
  const f = (r) => v.reduce((t, x, i) => t + x / Math.pow(1 + r, i), 0); let lo = -0.9999, hi = 10, flo = f(lo), fhi = f(hi);
  if (flo * fhi > 0) throw err(E.NUM);
  for (let i = 0; i < 200; i++) { const mid = (lo + hi) / 2, fm = f(mid); if (Math.abs(fm) < 1e-12) return mid; if (flo * fm < 0) { hi = mid; fhi = fm; } else { lo = mid; flo = fm; } }
  return (lo + hi) / 2;
});
def('DAYS', (a) => { need(a, 2); return Math.floor(toNum(a[0])) - Math.floor(toNum(a[1])); });
def('EDATE', (a) => { const d = dateOf(toNum(a[0])); return serialOf(d.getUTCFullYear(), d.getUTCMonth() + 1 + toNum(a[1]), d.getUTCDate()); });
def('EOMONTH', (a) => { const d = dateOf(toNum(a[0])); return serialOf(d.getUTCFullYear(), d.getUTCMonth() + 2 + toNum(a[1]), 0); });
def('DATEDIF', (a) => {
  need(a, 3); const s = dateOf(toNum(a[0])), e = dateOf(toNum(a[1])); if (e < s) throw err(E.NUM);
  let m = (e.getUTCFullYear() - s.getUTCFullYear()) * 12 + e.getUTCMonth() - s.getUTCMonth(); if (e.getUTCDate() < s.getUTCDate()) m--;
  const u = toStr(a[2]).toUpperCase();
  if (u === 'Y') return Math.floor(m / 12); if (u === 'M') return m; if (u === 'D') return Math.round((e - s) / 86400000); if (u === 'YM') return m % 12;
  throw err(E.NUM);
});

// dynamic arrays & SUBTOTAL
def('UNIQUE', (a) => {
  const t = rngOf(a[0]), seen = new Set(), rows = [];
  for (const r of t.rows) { const k = JSON.stringify(r.map((v) => (typeof v === 'string' ? v.toLowerCase() : v))); if (!seen.has(k)) { seen.add(k); rows.push(r); } }
  return new Rng(rows);
});
def('FILTER', (a) => {
  need(a, 2); const t = rngOf(a[0]), inc = rngOf(a[1]);
  if (inc.w !== 1 || inc.h !== t.h) throw err(E.VALUE);
  const rows = t.rows.filter((r, i) => { const v = inc.rows[i][0]; if (v instanceof XlError) throw v; return v === true || (typeof v === 'number' && v !== 0); });
  if (!rows.length) { if (a.length > 2 && a[2] !== null) return a[2]; throw err(E.CALC); }
  return new Rng(rows);
});
def('SORT', (a) => {
  const t = rngOf(a[0]), col = a.length > 1 && a[1] !== null ? toNum(a[1]) : 1, asc = !(a.length > 2 && a[2] !== null && toNum(a[2]) === -1);
  const rows = [...t.rows].sort((x, y) => { const p = x[col - 1], q = y[col - 1]; const r = typeof p === 'number' && typeof q === 'number' ? p - q : String(p ?? '').toLowerCase().localeCompare(String(q ?? '').toLowerCase()); return asc ? r : -r; });
  return new Rng(rows);
});
def('SUBTOTAL', (n, cx) => {
  if (n.length < 2) throw err(E.VALUE);
  const code = toNum(ev(n[0], cx)), base = code > 100 ? code - 100 : code, vals = [];
  for (const node of n.slice(1)) {
    if (node.t === 'range') {
      const S = node.sh ? cx.sheet.other(node.sh) : cx.sheet;
      const r1 = Math.min(node.a.r, node.b.r), r2 = Math.min(Math.max(node.a.r, node.b.r), Math.max(S.maxRow(), r1)), c1 = Math.min(node.a.c, node.b.c), c2 = Math.max(node.a.c, node.b.c);
      for (let r = r1; r <= r2; r++) { if (S.hidden.has(r)) continue; for (let c = c1; c <= c2; c++) vals.push(S.valueAt(c, r)); }
    } else vals.push(scalar(ev(node, cx)));
  }
  for (const v of vals) if (v instanceof XlError) throw v;
  const num = vals.filter((v) => typeof v === 'number');
  switch (base) {
    case 1: if (!num.length) throw err(E.DIV0); return sum(num) / num.length;
    case 2: return num.length; case 3: return vals.filter((v) => v !== null).length;
    case 4: return num.length ? Math.max(...num) : 0; case 5: return num.length ? Math.min(...num) : 0; case 9: return sum(num);
    default: throw err(E.VALUE);
  }
}, true);


/* ───────────── localized function names (pt-BR / es) ───────────── */
const nrm = (s) => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase();
const ALIAS = {};
const LOCAL = {};
const al = (en, names) => { const p = names.split('|'); LOCAL[en] = { pt: p[0], es: p[1] || p[0] }; p.forEach((n) => { ALIAS[nrm(n)] = en; }); };
// [English, Português (Brasil) | Español]
[
  ['SUM', 'SOMA|SUMA'], ['AVERAGE', 'MÉDIA|PROMEDIO'], ['MIN', 'MÍNIMO'], ['MAX', 'MÁXIMO'], ['COUNT', 'CONT.NÚM|CONTAR'], ['COUNTA', 'CONT.VALORES|CONTARA'], ['COUNTBLANK', 'CONTAR.VAZIO|CONTAR.BLANCO'],
  ['IF', 'SE|SI'], ['AND', 'E|Y'], ['OR', 'OU|O'], ['NOT', 'NÃO|NO'], ['XOR', 'XOU|XO'], ['IFS', 'SES|SI.CONJUNTO'], ['IFERROR', 'SEERRO|SI.ERROR'], ['IFNA', 'SEND|SI.ND'], ['SWITCH', 'PARÂMETRO|CAMBIAR'],
  ['ISERROR', 'ÉERRO|ESERROR'], ['ISNA', 'ÉNÃO.DISP|ESNOD'], ['ISNUMBER', 'ÉNÚM|ESNUMERO'], ['ISTEXT', 'ÉTEXTO|ESTEXTO'], ['ISBLANK', 'ÉCÉL.VAZIA|ESBLANCO'], ['ISEVEN', 'ÉPAR|ES.PAR'], ['ISODD', 'ÉIMPAR|ES.IMPAR'],
  ['COUNTIF', 'CONT.SE|CONTAR.SI'], ['COUNTIFS', 'CONT.SES|CONTAR.SI.CONJUNTO'], ['SUMIF', 'SOMASE|SUMAR.SI'], ['SUMIFS', 'SOMASES|SUMAR.SI.CONJUNTO'], ['AVERAGEIF', 'MÉDIASE|PROMEDIO.SI'], ['AVERAGEIFS', 'MÉDIASES|PROMEDIO.SI.CONJUNTO'],
  ['MAXIFS', 'MÁXIMOSES|MAX.SI.CONJUNTO'], ['MINIFS', 'MÍNIMOSES|MIN.SI.CONJUNTO'], ['SUMPRODUCT', 'SOMARPRODUTO|SUMAPRODUCTO'],
  ['VLOOKUP', 'PROCV|BUSCARV'], ['HLOOKUP', 'PROCH|BUSCARH'], ['XLOOKUP', 'PROCX|BUSCARX'], ['MATCH', 'CORRESP|COINCIDIR'], ['INDEX', 'ÍNDICE|INDICE'], ['CHOOSE', 'ESCOLHER|ELEGIR'],
  ['LEFT', 'ESQUERDA|IZQUIERDA'], ['RIGHT', 'DIREITA|DERECHA'], ['MID', 'EXT.TEXTO|EXTRAE'], ['LEN', 'NÚM.CARACT|LARGO'], ['UPPER', 'MAIÚSCULA|MAYUSC'], ['LOWER', 'MINÚSCULA|MINUSC'], ['PROPER', 'PRI.MAIÚSCULA|NOMPROPIO'], ['TRIM', 'ARRUMAR|ESPACIOS'],
  ['CONCATENATE', 'CONCATENAR'], ['TEXTJOIN', 'UNIRTEXTO|UNIRCADENAS'], ['FIND', 'PROCURAR|ENCONTRAR'], ['SEARCH', 'LOCALIZAR|HALLAR'], ['SUBSTITUTE', 'SUBSTITUIR|SUSTITUIR'], ['REPT', 'REPT|REPETIR'], ['VALUE', 'VALOR'], ['TEXT', 'TEXTO'],
  ['TEXTBEFORE', 'TEXTOANTES'], ['TEXTAFTER', 'TEXTODEPOIS|TEXTODESPUES'], ['ROUND', 'ARRED|REDONDEAR'], ['ROUNDUP', 'ARREDONDAR.PARA.CIMA|REDONDEAR.MAS'], ['ROUNDDOWN', 'ARREDONDAR.PARA.BAIXO|REDONDEAR.MENOS'],
  ['INT', 'INT|ENTERO'], ['SQRT', 'RAIZ'], ['POWER', 'POTÊNCIA|POTENCIA'], ['MOD', 'MOD|RESIDUO'], ['MEDIAN', 'MED|MEDIANA'], ['MODE', 'MODO|MODA'], ['LARGE', 'MAIOR|K.ESIMO.MAYOR'], ['SMALL', 'MENOR|K.ESIMO.MENOR'], ['RANK', 'ORDEM|JERARQUIA'],
  ['CEILING', 'TETO|MULTIPLO.SUPERIOR'], ['FLOOR', 'ARREDMULTB|MULTIPLO.INFERIOR'], ['ODD', 'ÍMPAR|REDONDEA.IMPAR'], ['EVEN', 'PAR|REDONDEA.PAR'], ['TRUNC', 'TRUNCAR'],
  ['DATE', 'DATA|FECHA'], ['TODAY', 'HOJE|HOY'], ['NOW', 'AGORA|AHORA'], ['YEAR', 'ANO|AÑO'], ['MONTH', 'MÊS|MES'], ['DAY', 'DIA'], ['WEEKDAY', 'DIA.DA.SEMANA|DIASEM'], ['NETWORKDAYS', 'DIATRABALHOTOTAL|DIAS.LAB'], ['WORKDAY', 'DIATRABALHO|DIA.LAB'], ['PMT', 'PGTO|PAGO'], ['FV', 'VF'], ['PV', 'VP|VA'], ['NPV', 'VPL|VNA'], ['IRR', 'TIR'], ['DAYS', 'DIAS'], ['DATEDIF', 'DATADIF|SIFECHA'], ['EDATE', 'DATAM|FECHA.MES'], ['EOMONTH', 'FIMMÊS|FIN.MES'],
  ['UNIQUE', 'ÚNICO|UNICOS'], ['FILTER', 'FILTRO|FILTRAR'], ['SORT', 'CLASSIFICAR|ORDENAR'], ['SUBTOTAL', 'SUBTOTAL|SUBTOTALES'], ['ROWS', 'LINS|FILAS'], ['COLUMNS', 'COLS|COLUMNAS'],
  ['ISLOGICAL', 'ÉLÓGICO|ESLOGICO'], ['PRODUCT', 'MULT|PRODUCTO'], ['ABS', 'ABS'], ['EXACT', 'EXATO|IGUAL'], ['CONCAT', 'CONCAT'],
].forEach(([en, names]) => al(en, names));
/** the name of a function in a language ('pt' | 'es'); English when there is no local name */
const localName = (en, lang) => (LOCAL[en] && LOCAL[en][lang]) || en;
const resolveName = (name) => { const u = nrm(name); return FN[u] ? u : ALIAS[u] || u; };
/** Rewrite a formula typed with localized names / ";" separators into the canonical English form. */
function canon(raw) {
  if (!raw || raw[0] !== '=') return raw;
  return raw.split(/("(?:[^"]|"")*")/).map((p, i) => (i % 2 ? p : p
    .replace(/;/g, ',')
    .replace(/([A-Za-z_\u00C0-\u024F][A-Za-z0-9_.\u00C0-\u024F]*)(?=\s*\()/g, (m) => ALIAS[nrm(m)] || m)
    .replace(/\b(VERDADEIRO|VERDADERO)\b/gi, 'TRUE').replace(/\bFALSO\b/gi, 'FALSE'))).join('');
}

/* ───────────── function docs (autocomplete / signature hints) ───────────── */
const DOCS = {
  XOR: ['XOR(logical1, [logical2], …)', 'TRUE if an odd number of arguments are TRUE.'],
  SWITCH: ['SWITCH(expression, value1, result1, [value2, result2], …, [default])', 'Compares a value to a list and returns the matching result.'],
  CHOOSE: ['CHOOSE(index_num, value1, [value2], …)', 'Picks a value from a list by position.'],
  ISLOGICAL: ['ISLOGICAL(value)', 'TRUE if the value is TRUE or FALSE.'], ISEVEN: ['ISEVEN(number)', 'TRUE if the number is even.'], ISODD: ['ISODD(number)', 'TRUE if the number is odd.'],
  FLOOR: ['FLOOR(number, significance)', 'Rounds a number down to a multiple of significance.'], CEILING: ['CEILING(number, significance)', 'Rounds a number up to a multiple of significance.'],
  TRUNC: ['TRUNC(number, [num_digits])', 'Cuts off decimals without rounding.'], ODD: ['ODD(number)', 'Rounds up to the nearest odd integer.'], EVEN: ['EVEN(number)', 'Rounds up to the nearest even integer.'],
  MODE: ['MODE(number1, [number2], …)', 'Returns the most frequent number.'],
  TEXTBEFORE: ['TEXTBEFORE(text, delimiter, [instance_num])', 'Returns the text before a delimiter.'], TEXTAFTER: ['TEXTAFTER(text, delimiter, [instance_num])', 'Returns the text after a delimiter.'],
  DATE: ['DATE(year, month, day)', 'Builds a date from year, month and day.'], TODAY: ['TODAY()', "Returns today's date."], YEAR: ['YEAR(serial_number)', 'Returns the year of a date.'],
  MONTH: ['MONTH(serial_number)', 'Returns the month (1-12) of a date.'], DAY: ['DAY(serial_number)', 'Returns the day of the month.'], WEEKDAY: ['WEEKDAY(serial_number, [return_type])', 'Returns the day of the week as a number.'],
  NETWORKDAYS: ['NETWORKDAYS(start_date, end_date, [holidays])', 'Counts the working days (Monday to Friday) between two dates, skipping any holidays.'], WORKDAY: ['WORKDAY(start_date, days, [holidays])', 'Returns the date that is a number of working days from a start date.'],
  PMT: ['PMT(rate, nper, pv, [fv], [type])', 'Returns the payment for a loan with constant payments and a constant interest rate.'], FV: ['FV(rate, nper, pmt, [pv], [type])', 'Returns the future value of an investment with regular payments.'],
  PV: ['PV(rate, nper, pmt, [fv], [type])', 'Returns the present value of a series of payments.'], NPV: ['NPV(rate, value1, [value2], …)', 'Returns the net present value of future cash flows at a discount rate.'], IRR: ['IRR(values, [guess])', 'Returns the internal rate of return of a series of cash flows.'],
  DAYS: ['DAYS(end_date, start_date)', 'Returns the number of days between two dates.'], EDATE: ['EDATE(start_date, months)', 'Moves a date by whole months.'], EOMONTH: ['EOMONTH(start_date, months)', 'Returns the last day of a month.'],
  DATEDIF: ['DATEDIF(start_date, end_date, unit)', 'Counts whole years ("Y"), months ("M") or days ("D") between two dates.'],
  UNIQUE: ['UNIQUE(array)', 'Returns the distinct values of a list (spills into the cells below).'], FILTER: ['FILTER(array, include, [if_empty])', 'Returns only the rows that meet a condition (spills).'],
  SORT: ['SORT(array, [sort_index], [sort_order])', 'Sorts a range (spills). Order 1 = ascending, -1 = descending.'], SUBTOTAL: ['SUBTOTAL(function_num, ref1, …)', 'Calculates a total that ignores rows hidden by a filter (9 = SUM, 1 = AVERAGE, 2 = COUNT).'],
  SUM: ['SUM(number1, [number2], …)', 'Adds all the numbers in a range of cells.'],
  AVERAGE: ['AVERAGE(number1, [number2], …)', 'Returns the average (arithmetic mean) of its arguments.'],
  MIN: ['MIN(number1, [number2], …)', 'Returns the smallest number in a set of values.'],
  MAX: ['MAX(number1, [number2], …)', 'Returns the largest number in a set of values.'],
  MEDIAN: ['MEDIAN(number1, [number2], …)', 'Returns the middle value of a set of numbers.'],
  COUNT: ['COUNT(value1, [value2], …)', 'Counts how many cells contain numbers.'],
  COUNTA: ['COUNTA(value1, [value2], …)', 'Counts how many cells are not empty.'],
  COUNTBLANK: ['COUNTBLANK(range)', 'Counts empty cells in a range.'],
  COUNTIF: ['COUNTIF(range, criteria)', 'Counts the cells that meet one condition.'],
  COUNTIFS: ['COUNTIFS(range1, criteria1, [range2, criteria2], …)', 'Counts the cells that meet several conditions.'],
  SUMIF: ['SUMIF(range, criteria, [sum_range])', 'Adds the cells that meet one condition.'],
  SUMIFS: ['SUMIFS(sum_range, range1, criteria1, [range2, criteria2], …)', 'Adds the cells that meet several conditions.'],
  AVERAGEIF: ['AVERAGEIF(range, criteria, [average_range])', 'Averages the cells that meet one condition.'],
  AVERAGEIFS: ['AVERAGEIFS(average_range, range1, criteria1, [range2, criteria2], …)', 'Averages the cells that meet several conditions.'],
  MAXIFS: ['MAXIFS(max_range, range1, criteria1, …)', 'Largest value among cells meeting the conditions.'],
  MINIFS: ['MINIFS(min_range, range1, criteria1, …)', 'Smallest value among cells meeting the conditions.'],
  SUMPRODUCT: ['SUMPRODUCT(array1, [array2], …)', 'Multiplies matching items in arrays, then adds the results.'],
  PRODUCT: ['PRODUCT(number1, [number2], …)', 'Multiplies all the numbers given.'],
  LARGE: ['LARGE(array, k)', 'Returns the k-th largest value.'],
  SMALL: ['SMALL(array, k)', 'Returns the k-th smallest value.'],
  RANK: ['RANK(number, ref, [order])', 'Returns the rank of a number in a list.'],
  ROUND: ['ROUND(number, num_digits)', 'Rounds a number to a number of digits.'],
  ROUNDUP: ['ROUNDUP(number, num_digits)', 'Rounds a number up, away from zero.'],
  ROUNDDOWN: ['ROUNDDOWN(number, num_digits)', 'Rounds a number down, toward zero.'],
  INT: ['INT(number)', 'Rounds a number down to the nearest integer.'],
  ABS: ['ABS(number)', 'Returns the absolute value of a number.'],
  SQRT: ['SQRT(number)', 'Returns a positive square root.'],
  POWER: ['POWER(number, power)', 'Raises a number to a power.'],
  MOD: ['MOD(number, divisor)', 'Returns the remainder after division.'],
  IF: ['IF(logical_test, [value_if_true], [value_if_false])', 'Returns one value if a condition is true and another if it is false.'],
  IFS: ['IFS(test1, value1, [test2, value2], …)', 'Checks several conditions and returns the value of the first true one.'],
  IFERROR: ['IFERROR(value, value_if_error)', 'Returns a fallback value if the formula produces an error.'],
  IFNA: ['IFNA(value, value_if_na)', 'Returns a fallback value if the result is #N/A.'],
  AND: ['AND(logical1, [logical2], …)', 'TRUE only if all arguments are TRUE.'],
  OR: ['OR(logical1, [logical2], …)', 'TRUE if any argument is TRUE.'],
  NOT: ['NOT(logical)', 'Reverses TRUE and FALSE.'],
  ISBLANK: ['ISBLANK(value)', 'TRUE if the cell is empty.'],
  ISNUMBER: ['ISNUMBER(value)', 'TRUE if the value is a number.'],
  ISTEXT: ['ISTEXT(value)', 'TRUE if the value is text.'],
  ISERROR: ['ISERROR(value)', 'TRUE if the value is any error.'],
  ISNA: ['ISNA(value)', 'TRUE if the value is #N/A.'],
  LEN: ['LEN(text)', 'Returns the number of characters in text.'],
  LEFT: ['LEFT(text, [num_chars])', 'Returns the first characters of text.'],
  RIGHT: ['RIGHT(text, [num_chars])', 'Returns the last characters of text.'],
  MID: ['MID(text, start_num, num_chars)', 'Returns characters from the middle of text.'],
  UPPER: ['UPPER(text)', 'Converts text to UPPERCASE.'],
  LOWER: ['LOWER(text)', 'Converts text to lowercase.'],
  PROPER: ['PROPER(text)', 'Capitalizes The First Letter Of Each Word.'],
  TRIM: ['TRIM(text)', 'Removes extra spaces from text.'],
  EXACT: ['EXACT(text1, text2)', 'TRUE if two texts are identical (case-sensitive).'],
  REPT: ['REPT(text, number_times)', 'Repeats text a given number of times.'],
  VALUE: ['VALUE(text)', 'Converts text that looks like a number into a number.'],
  CONCAT: ['CONCAT(text1, [text2], …)', 'Joins text from several cells or ranges.'],
  CONCATENATE: ['CONCATENATE(text1, [text2], …)', 'Joins several pieces of text.'],
  TEXTJOIN: ['TEXTJOIN(delimiter, ignore_empty, text1, …)', 'Joins text using a separator.'],
  FIND: ['FIND(find_text, within_text, [start_num])', 'Finds text inside text (case-sensitive).'],
  SEARCH: ['SEARCH(find_text, within_text, [start_num])', 'Finds text inside text (not case-sensitive).'],
  SUBSTITUTE: ['SUBSTITUTE(text, old_text, new_text, [instance_num])', 'Replaces old text with new text.'],
  TEXT: ['TEXT(value, format_text)', 'Formats a number as text, e.g. "0.00" or "0%".'],
  VLOOKUP: ['VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])', 'Looks down the first column of a table and returns a value from the same row.'],
  HLOOKUP: ['HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])', 'Looks across the top row of a table and returns a value from the same column.'],
  MATCH: ['MATCH(lookup_value, lookup_array, [match_type])', 'Returns the position of an item in a range.'],
  INDEX: ['INDEX(array, row_num, [column_num])', 'Returns the value at a position inside a range.'],
  XLOOKUP: ['XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode])', 'Finds a value in one range and returns the matching value from another.'],
  ROWS: ['ROWS(array)', 'Returns the number of rows in a range.'],
  COLUMNS: ['COLUMNS(array)', 'Returns the number of columns in a range.'],
};

/* ───────────── references shifting (fill / copy / paste) ───────────── */
function shiftFormula(raw, dr, dc) {
  return raw.split(/("(?:[^"]|"")*")/).map((part, i) => (i % 2 ? part : part.replace(
    /(?<![A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/g,
    (m, d1, col, d2, row) => {
      let c = colToNum(col), r = +row;
      if (!d1) c += dc; if (!d2) r += dr;
      if (c < 1 || r < 1) return '#REF!';
      return d1 + numToCol(c) + d2 + r;
    }))).join('');
}
function autoClose(text) {
  if (text[0] !== '=') return text;
  let depth = 0, inS = false;
  for (const ch of text) { if (ch === '"') inS = !inS; else if (!inS) { if (ch === '(') depth++; else if (ch === ')') depth--; } }
  return inS ? text : text + ')'.repeat(Math.max(0, depth));
}

/* ───────────── structure edits: formula rewriting for inserted / deleted rows and columns ───────────── */
const STYLE_KEYS = ['bold', 'italic', 'underline', 'align', 'fill', 'color', 'wrap', 'border'];
const REF_OR_RANGE = /(?<![A-Za-z0-9_.$!])(?:(\$?[A-Za-z]{1,3}\$?\d+):(\$?[A-Za-z]{1,3}\$?\d+)|(\$?[A-Za-z]{1,3}\$?\d+))(?![A-Za-z0-9_(])/g;
function adjustFormula(raw, axis, op, at, n) {
  if (raw[0] !== '=') return raw;
  const pt = (i) => (op === 'ins' ? (i >= at ? i + n : i) : i < at ? i : i >= at + n ? i - n : null);
  const rg = (a, b) => (op === 'ins' ? [pt(a), pt(b)] : [a < at ? a : a >= at + n ? a - n : at, b < at ? b : b >= at + n ? b - n : at - 1]);
  const parse = (t) => { const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(t); return { d1: m[1], c: colToNum(m[2]), d2: m[3], r: +m[4] }; };
  const show = (p) => p.d1 + numToCol(p.c) + p.d2 + p.r;
  const k = axis === 'r' ? 'r' : 'c';
  return raw.split(/("(?:[^"]|"")*")/).map((seg, i) => (i % 2 ? seg : seg.replace(REF_OR_RANGE, (m, a, b, one) => {
    if (one) { const p = parse(one), v = pt(p[k]); if (v === null) return '#REF!'; p[k] = v; return show(p); }
    const p = parse(a), q = parse(b), [x, y] = rg(p[k], q[k]); if (y < x) return '#REF!'; p[k] = x; q[k] = y; return show(p) + ':' + show(q);
  }))).join('');
}
const textOf = (v) => (v === null || v === undefined ? '' : v instanceof XlError ? v.code : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v));
const escRe = (t) => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/* Flash Fill: find a text transformation that explains the examples, then apply it to the empty cells */
function flashCandidates(ex) {
  const words = (t) => t.trim().split(/\s+/).filter(Boolean), cap = (w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
  const list = [
    ['same', (a) => a], ['upper', (a) => a.toUpperCase()], ['lower', (a) => a.toLowerCase()], ['proper', (a) => words(a).map(cap).join(' ')],
    ['first word', (a) => words(a)[0]], ['last word', (a) => { const w = words(a); return w[w.length - 1]; }], ['second word', (a) => words(a)[1]],
    ['all but last', (a) => { const w = words(a); return w.length > 1 ? w.slice(0, -1).join(' ') : undefined; }],
    ['initials', (a) => words(a).map((w) => w[0].toUpperCase()).join('')], ['initials.', (a) => words(a).map((w) => w[0].toUpperCase() + '.').join('')],
    ['last, first', (a) => { const w = words(a); return w.length > 1 ? w[w.length - 1] + ', ' + w.slice(0, -1).join(' ') : undefined; }],
    ['first last', (a) => { const m = /^(.*?),\s*(.+)$/.exec(a); return m ? m[2] + ' ' + m[1] : undefined; }],
    ['digits', (a) => a.replace(/\D/g, '')], ['letters', (a) => a.replace(/[^A-Za-z]/g, '')], ['no spaces', (a) => a.replace(/\s+/g, '')],
  ];
  for (const ch of ['@', ' ', ',', '-', '_', '/', '.', '(', ')']) {
    list.push(['before ' + ch, (a) => { const i = a.indexOf(ch); return i < 0 ? undefined : a.slice(0, i); }]);
    list.push(['after ' + ch, (a) => { const i = a.indexOf(ch); return i < 0 ? undefined : a.slice(i + 1).trim(); }]);
    list.push(['last after ' + ch, (a) => { const i = a.lastIndexOf(ch); return i < 0 ? undefined : a.slice(i + 1).trim(); }]);
  }
  list.push(['between ()', (a) => { const m = /\(([^)]*)\)/.exec(a); return m ? m[1] : undefined; }]);
  const e0 = ex[0]; if (e0) {
    const src = e0.s[0], out = e0.out;
    const i = src.indexOf(out); if (i >= 0 && out) { list.push(['substring', (a) => (a.length >= i + out.length ? a.slice(i, i + out.length) : undefined)]); }
    const j = out.indexOf(src); if (j >= 0 && src) { const pre = out.slice(0, j), suf = out.slice(j + src.length); list.push(['wrap', (a) => pre + a + suf]); }
    if (e0.s.length > 1) { const a = e0.s[0], b = e0.s[1], ia = out.indexOf(a), ib = out.indexOf(b);
      if (a && b && ia >= 0 && ib >= 0 && ia + a.length <= ib) { const sep = out.slice(ia + a.length, ib), pre = out.slice(0, ia), suf = out.slice(ib + b.length); list.push(['join', (x, y) => pre + x + sep + y + suf]); }
      if (a && b && ia >= 0 && ib >= 0 && ib + b.length <= ia) { const sep = out.slice(ib + b.length, ia), pre = out.slice(0, ib), suf = out.slice(ia + a.length); list.push(['join reversed', (x, y) => pre + y + sep + x + suf]); }
    }
  }
  return ex[0] && ex[0].s.length > 1 ? list.filter((x) => x[0].startsWith('join')) : list;
}

/* ───────────── Sheet ───────────── */
class Sheet {
  constructor(rows = 10, cols = 8) {
    this.rows = rows; this.cols = cols; this.cells = {}; this._c = new Map(); this._st = new Set();
    this.validation = []; this.hidden = new Set(); this.hiddenCols = new Set(); this.colW = {}; this.freeze = { rows: 0, cols: 0 }; this.cf = []; this.charts = []; this.filter = null; this.book = null; this.name = 'Sheet1'; this._spill = new Map(); this._spillOK = false;
  }
  _inv() {
    this._c.clear(); this._spillOK = false;
    if (this.book) for (const s of Object.values(this.book.sheets)) if (s !== this) { s._c.clear(); s._spillOK = false; }
  }
  other(name) {
    const b = this.book; if (b) { const k = Object.keys(b.sheets).find((n) => n.toLowerCase() === String(name).toLowerCase()); if (k) return b.sheets[k]; }
    throw err(E.REF);
  }
  load(data) {
    for (const k of Object.keys(data || {})) { const raw = String(data[k]); this.cells[k] = { raw }; if (raw[0] !== '=' && isDateRaw(raw)) this.cells[k].fmt = 'date'; }
    this._inv(); return this;
  }
  raw(ref) { const c = this.cells[ref]; return c ? c.raw : ''; }
  fmt(ref) { const c = this.cells[ref]; return c ? c.fmt || null : null; }
  setFmt(ref, f) { const c = this.cells[ref] || (this.cells[ref] = { raw: '' }); c.fmt = f; this._inv(); }
  setBold(ref, b) { const c = this.cells[ref] || (this.cells[ref] = { raw: '' }); c.bold = b; }
  setRaw(ref, raw) {
    const c = this.cells[ref] || (this.cells[ref] = { raw: '' });
    c.raw = raw; if (raw === '' && !this.hasStyle(c)) delete this.cells[ref];
    this._inv();
  }
  commit(ref, text) {
    text = autoClose(text);
    const t = text.trim(); let fmt = null;
    if (text[0] !== '=') { if (/^[+-]?(\d+\.?\d*|\.\d+)%$/.test(t)) fmt = 'percent'; else if (/^\$\d[\d,]*(\.\d+)?$/.test(t)) fmt = 'currency'; else if (isDateRaw(t)) fmt = 'date'; }
    else if (/^=\s*(DATE|TODAY|EDATE|EOMONTH|WORKDAY)\(/i.test(text)) fmt = 'date';
    this.setRaw(ref, text);
    if (fmt && this.cells[ref] && !this.cells[ref].fmt) this.cells[ref].fmt = fmt;
  }
  copyStyle(from, to) { const s = this.cells[from], d = this.cells[to]; if (!s && !d) return; const c = d || (this.cells[to] = { raw: '' }); c.fmt = s ? s.fmt : undefined; for (const k of STYLE_KEYS) { if (s && s[k]) c[k] = s[k]; else delete c[k]; } }
  maxRow() { let m = this.rows; for (const k in this.cells) { const p = parseRef(k); if (p && p.r > m && this.cells[k].raw !== '') m = p.r; } return m; }
  usedEnd() { let r = 1, c = 1; for (const k in this.cells) { if (this.cells[k].raw === '') continue; const p = parseRef(k); if (p) { r = Math.max(r, p.r); c = Math.max(c, p.c); } } return { r, c }; }
  _ensureSpill() {
    if (this._spillOK) return; this._spillOK = true; this._spill = new Map();
    for (const k of Object.keys(this.cells)) {
      const raw = this.cells[k].raw; if (raw[0] !== '=' || !/\b(UNIQUE|FILTER|SORT)\s*\(/i.test(raw)) continue;
      const p = parseRef(k); let v;
      try { let ast = AST.get(raw); if (!ast) { ast = parse(raw.slice(1)); AST.set(raw, ast); } this._st.add(k); try { v = ev(ast, { sheet: this }); } finally { this._st.delete(k); } } catch (e) { continue; }
      if (!(v instanceof Rng) || v.h * v.w < 2) continue;
      let blocked = false; const out = [];
      for (let i = 0; i < v.h; i++) for (let j = 0; j < v.w; j++) {
        if (!i && !j) continue; const ref = numToCol(p.c + j) + (p.r + i);
        if (this.raw(ref) !== '' || this._spill.has(ref)) blocked = true; out.push([ref, v.rows[i][j]]);
      }
      if (blocked) { this._c.set(k, err(E.SPILL)); continue; }
      this._c.set(k, v.rows[0][0]); for (const [r, x] of out) this._spill.set(r, x);
    }
  }
  cfFor(ref) {
    if (!this.cf.length) return null;
    const p = parseRef(ref);
    for (const rule of this.cf) {
      const [a, z] = rule.range.split(':'), A = parseRef(a), Z = parseRef(z || a);
      if (p.r < A.r || p.r > Z.r || p.c < A.c || p.c > Z.c) continue;
      const v = this.value(ref); let ok = false;
      switch (rule.type) {
        case 'gt': ok = typeof v === 'number' && v > rule.a; break;
        case 'lt': ok = typeof v === 'number' && v < rule.a; break;
        case 'eq': ok = v !== null && (typeof v === 'number' ? v === Number(rule.a) : String(v).toLowerCase() === String(rule.a).toLowerCase()); break;
        case 'between': ok = typeof v === 'number' && v >= Math.min(rule.a, rule.b) && v <= Math.max(rule.a, rule.b); break;
        case 'text': ok = typeof v === 'string' && v.toLowerCase().includes(String(rule.a).toLowerCase()); break;
        case 'dup': {
          if (v === null) break; let n = 0;
          for (let r = A.r; r <= Z.r; r++) for (let c = A.c; c <= Z.c; c++) { const w = this.valueAt(c, r); if (w !== null && eqLoose(w, v)) n++; }
          ok = n > 1; break;
        }
        default: break;
      }
      if (ok) return rule.style;
    }
    return null;
  }
  value(ref) {
    const cell = this.cells[ref];
    if (!cell || cell.raw === '') { this._ensureSpill(); return this._spill.has(ref) ? this._spill.get(ref) : null; }
    if (cell.raw[0] === '=') this._ensureSpill();
    if (this._c.has(ref)) return this._c.get(ref);
    if (this._st.has(ref)) return err(E.CIRC);
    const raw = cell.raw; let v;
    if (raw[0] === '=') {
      this._st.add(ref);
      try {
        let ast = AST.get(raw); if (!ast) { ast = parse(raw.slice(1)); AST.set(raw, ast); }
        v = scalar(ev(ast, { sheet: this })); if (v === null) v = 0;
      } catch (e) { v = e instanceof XlError ? e : err(E.ERR); } finally { this._st.delete(ref); }
    } else v = parseLiteral(raw);
    this._c.set(ref, v); return v;
  }
  valueAt(c, r, sh) { return sh ? this.other(sh).valueAt(c, r) : this.value(numToCol(c) + r); }
  display(ref) { return formatValue(this.value(ref), this.fmt(ref)); }
  kind(ref) { const v = this.value(ref); return v === null ? 'empty' : v instanceof XlError ? 'err' : typeof v === 'number' ? 'num' : typeof v === 'boolean' ? 'bool' : 'text'; }
  snapshot() { return JSON.stringify({ cells: this.cells, hidden: [...this.hidden], hiddenCols: [...this.hiddenCols], colW: this.colW, freeze: this.freeze, cf: this.cf, charts: this.charts, names: this._names(), validation: this.validation }); }
  restore(s) { const o = JSON.parse(s); if (o.cells === undefined) { this.cells = o; } else { this.cells = o.cells; this.hidden = new Set(o.hidden); this.hiddenCols = new Set(o.hiddenCols); this.colW = o.colW; this.freeze = o.freeze; this.cf = o.cf; this.charts = o.charts; this.validation = o.validation || []; if (this.book) this.book.names = o.names || {}; else this._nm = o.names || {}; } this._inv(); }

  /* ── styles ── */
  /** lesson set-up: starting styles, column widths, hidden rows / columns, frozen panes */
  setup(g) {
    if (g.styles) for (const ref of Object.keys(g.styles)) for (const [key, val] of Object.entries(g.styles[ref])) this.setStyle(ref, key, val);
    if (g.colW) Object.assign(this.colW, g.colW);
    if (g.hidden) g.hidden.forEach((r) => this.hidden.add(r));
    if (g.hiddenCols) g.hiddenCols.forEach((c) => this.hiddenCols.add(c));
    if (g.freeze) Object.assign(this.freeze, g.freeze);
    if (g.names) for (const k of Object.keys(g.names)) this.defineName(k, g.names[k]);
    if (g.validation) g.validation.forEach((v) => this.validation.push(Object.assign({}, v)));
    this._inv(); return this;
  }
  hasStyle(c) { return !!(c.fmt || STYLE_KEYS.some((k) => c[k])); }
  setStyle(ref, key, val) {
    const c = this.cells[ref] || (this.cells[ref] = { raw: '' });
    if (val === undefined || val === null || val === false) delete c[key]; else c[key] = val;
    if (c.raw === '' && !this.hasStyle(c)) delete this.cells[ref];
  }
  style(ref, key) { const c = this.cells[ref]; return c ? c[key] : undefined; }
  clearStyles(ref) { const c = this.cells[ref]; if (!c) return; for (const k of STYLE_KEYS) delete c[k]; delete c.fmt; if (c.raw === '') delete this.cells[ref]; this._inv(); }
  /** the width a column needs to show its longest entry (px), used by AutoFit */
  neededWidth(col) {
    let w = 0; const L = numToCol(col);
    for (let r = 1; r <= this.rows; r++) { const ref = L + r, c = this.cells[ref]; if (!c || c.raw === '') continue; if (c.wrap) continue; w = Math.max(w, this.display(ref).length * (c.bold ? 9 : 8.2)); }
    return Math.max(60, Math.ceil(w + 22));
  }
  autofit(col) { this.colW[col] = this.neededWidth(col); }
  isCut(col) { const need = this.neededWidth(col), have = this.colW[col] || 104; return need > have + 1; }

  /* ── insert / delete rows and columns (formulas, hidden rows, rules and charts follow) ── */
  _edit(axis, op, at, n) {
    const next = {}, lim = axis === 'r' ? this.rows : this.cols, k = axis === 'r' ? 'r' : 'c';
    for (const ref of Object.keys(this.cells)) {
      const p = parseRef(ref); const i = p[k]; let ni;
      if (op === 'ins') ni = i >= at ? i + n : i; else { if (i >= at && i < at + n) continue; ni = i >= at + n ? i - n : i; }
      if (ni > lim) continue;
      const np = { r: p.r, c: p.c }; np[k] = ni; next[numToCol(np.c) + np.r] = Object.assign({}, this.cells[ref]);
    }
    for (const ref of Object.keys(next)) { const c = next[ref]; if (c.raw && c.raw[0] === '=') c.raw = adjustFormula(c.raw, axis, op, at, n); }
    this.cells = next;
    const mapIdx = (i) => (op === 'ins' ? (i >= at ? i + n : i) : i >= at && i < at + n ? null : i >= at + n ? i - n : i);
    const mapSet = (set) => { const o = new Set(); for (const i of set) { const v = mapIdx(i); if (v !== null && v <= lim) o.add(v); } return o; };
    if (axis === 'r') this.hidden = mapSet(this.hidden); else { this.hiddenCols = mapSet(this.hiddenCols); const w = {}; for (const i of Object.keys(this.colW)) { const v = mapIdx(+i); if (v !== null) w[v] = this.colW[i]; } this.colW = w; }
    const fix = (rng) => { const f = adjustFormula('=' + rng, axis, op, at, n); return f === '=#REF!' ? null : f.slice(1); };
    this.cf = this.cf.map((r) => { const g = fix(r.range); return g ? Object.assign({}, r, { range: g }) : null; }).filter(Boolean);
    this.charts = this.charts.map((r) => { const g = fix(r.range); return g ? Object.assign({}, r, { range: g }) : null; }).filter(Boolean);
    this.validation = this.validation.map((r) => { const g = fix(r.range); return g ? Object.assign({}, r, { range: g }) : null; }).filter(Boolean);
    const nm = this._names(); for (const k of Object.keys(nm)) { if (/^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/.test(nm[k].ref)) { const g = fix(nm[k].ref); if (g) nm[k].ref = g; else delete nm[k]; } }
    this._inv();
  }
  insertRows(at, n = 1) { this._edit('r', 'ins', at, n); }
  deleteRows(at, n = 1) { this._edit('r', 'del', at, n); }
  insertCols(at, n = 1) { this._edit('c', 'ins', at, n); }
  deleteCols(at, n = 1) { this._edit('c', 'del', at, n); }
  hideRows(a, b) { for (let r = a; r <= b; r++) this.hidden.add(r); this._inv(); }
  showRows(a, b) { for (let r = a; r <= b; r++) this.hidden.delete(r); this._inv(); }
  hideCols(a, b) { for (let c = a; c <= b; c++) this.hiddenCols.add(c); }
  showCols(a, b) { for (let c = a; c <= b; c++) this.hiddenCols.delete(c); }

  /* ── find, replace and Go To Special ── */
  _scan(range) {
    const out = []; let b = { r1: 1, c1: 1, r2: this.rows, c2: this.cols };
    if (range) { const [x, y] = range.split(':'), p = parseRef(x), q = parseRef(y || x); b = { r1: Math.min(p.r, q.r), r2: Math.max(p.r, q.r), c1: Math.min(p.c, q.c), c2: Math.max(p.c, q.c) }; }
    for (let r = b.r1; r <= b.r2; r++) for (let c = b.c1; c <= b.c2; c++) out.push(numToCol(c) + r);
    return out;
  }
  findAll(find, o = {}) {
    if (find === '') return [];
    const q = o.matchCase ? find : find.toLowerCase();
    return this._scan(o.range).filter((ref) => {
      const c = this.cells[ref]; if (!c || c.raw === '') return false;
      const hay = o.matchCase ? c.raw : c.raw.toLowerCase();
      return o.whole ? hay === q : hay.includes(q);
    });
  }
  replaceAll(find, repl, o = {}) {
    const hits = this.findAll(find, o); const re = new RegExp(escRe(find), o.matchCase ? 'g' : 'gi');
    for (const ref of hits) { const c = this.cells[ref]; const next = o.whole ? repl : c.raw.replace(re, () => repl); this.commit(ref, next); }
    this._inv(); return hits.length;
  }
  selectSpecial(kind, range) {
    let box = range; if (!box || !box.includes(':')) { const u = this.usedEnd(); box = 'A1:' + numToCol(u.c) + u.r; }
    return this._scan(box).filter((ref) => {
      const raw = (this.cells[ref] || { raw: '' }).raw;
      if (kind === 'blanks') return raw === '';
      if (kind === 'formulas') return raw[0] === '=';
      if (kind === 'constants') return raw !== '' && raw[0] !== '=';
      if (kind === 'errors') return raw !== '' && this.value(ref) instanceof XlError;
      return false;
    });
  }

  /* ── copy blocks and Paste Special ── */
  copyBlock(r1, c1, r2, c2) {
    const rows = [];
    for (let r = r1; r <= r2; r++) { const row = []; for (let c = c1; c <= c2; c++) { const ref = numToCol(c) + r, cell = this.cells[ref] || { raw: '' }; const st = {}; for (const k of STYLE_KEYS) if (cell[k]) st[k] = cell[k]; row.push({ raw: cell.raw, fmt: cell.fmt || null, st, val: this.value(ref), sr: r, sc: c }); } rows.push(row); }
    return { rows, r1, c1, r2, c2 };
  }
  /** mode: all | values | formulas | formats ;  op: add | subtract | multiply | divide ;  transpose: boolean */
  pasteBlock(block, r, c, o = {}) {
    const mode = o.mode || 'all', put = [];
    block.rows.forEach((row, i) => row.forEach((cell, j) => { const dr = o.transpose ? r + j : r + i, dc = o.transpose ? c + i : c + j; if (dr >= 1 && dc >= 1 && dr <= this.rows && dc <= this.cols) put.push([cell, dr, dc]); }));
    for (const [cell, dr, dc] of put) {
      const ref = numToCol(dc) + dr;
      if (mode === 'all' || mode === 'formulas') {
        const raw = cell.raw[0] === '=' ? shiftFormula(cell.raw, dr - cell.sr, dc - cell.sc) : cell.raw;
        if (!o.op) this.setRaw(ref, raw);
      }
      if (mode === 'values') { if (!o.op) this.setRaw(ref, textOf(cell.val)); }
      if (o.op && mode !== 'formats') {
        const cur = this.value(ref), a = typeof cur === 'number' ? cur : (typeof cur === 'string' && cur.trim() !== '' && !isNaN(Number(cur.replace(/,/g, ''))) ? Number(cur.replace(/,/g, '')) : null), b = typeof cell.val === 'number' ? cell.val : null;
        if (a !== null && b !== null) { const v = o.op === 'add' ? a + b : o.op === 'subtract' ? a - b : o.op === 'multiply' ? a * b : b === 0 ? null : a / b; if (v !== null) this.setRaw(ref, String(v)); }
      }
      if (mode === 'all' || mode === 'formats') { const d = this.cells[ref] || (this.cells[ref] = { raw: '' }); d.fmt = cell.fmt || undefined; for (const k of STYLE_KEYS) { if (cell.st[k]) d[k] = cell.st[k]; else delete d[k]; } if (d.raw === '' && !this.hasStyle(d)) delete this.cells[ref]; }
    }
    this._inv(); return put.length;
  }

  /* ── named ranges ── */
  _names() { return this.book ? (this.book.names || (this.book.names = {})) : (this._nm || (this._nm = {})); }
  nameRef(name) { const d = this._names()[String(name).toLowerCase()]; return d ? d.ref : null; }
  listNames() { return Object.values(this._names()).sort((a, b) => a.name.localeCompare(b.name)); }
  /** returns null when defined, or 'invalid' / 'cell' for a bad name */
  defineName(name, ref) {
    if (!/^[A-Za-z_\u00C0-\u024F][A-Za-z0-9_.\u00C0-\u024F]*$/.test(name)) return 'invalid';
    if (parseRef(name) || /^[rRcC]$/.test(name) || /^(TRUE|FALSE)$/i.test(name)) return 'cell';
    this._names()[name.toLowerCase()] = { name, ref: String(ref).replace(/\$/g, '').toUpperCase().replace(/^([^!]+)!/, (m0, sh) => sh + '!') }; this._inv(); return null;
  }
  clearValidation(range) { const b = this._box(range); this.validation = this.validation.filter((v) => { const q = this._box(v.range); return !(q.r1 >= b.r1 && q.r2 <= b.r2 && q.c1 >= b.c1 && q.c2 <= b.c2); }); }
  deleteName(name) { delete this._names()[String(name).toLowerCase()]; this._inv(); }

  /* ── data validation (drop-down lists, whole numbers, decimals) ── */
  setValidation(range, rule) { this.validation = this.validation.filter((v) => v.range !== range); if (rule) this.validation.push(Object.assign({ range }, rule)); }
  validationFor(ref) {
    const p = parseRef(ref);
    for (let i = this.validation.length - 1; i >= 0; i--) { const b = this._box(this.validation[i].range); if (p.r >= b.r1 && p.r <= b.r2 && p.c >= b.c1 && p.c <= b.c2) return this.validation[i]; }
    return null;
  }
  listItems(rule) {
    const src = String(rule.source || '').trim();
    if (src[0] === '=') { const nm = src.slice(1).trim(), ref = (this.nameRef(nm) || nm).replace(/\$/g, ''); if (!/^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/.test(ref)) return []; return this._scan(ref.includes(':') ? ref : ref + ':' + ref).filter((r) => (this.cells[r] || { raw: '' }).raw !== '').map((r) => this.display(r)); }
    return src.split(',').map((x) => x.trim()).filter(Boolean);
  }
  /** null when the entry is allowed, otherwise the kind of problem */
  validate(ref, text) {
    const v = this.validationFor(ref), t = String(text).trim(); if (!v || t === '' || t[0] === '=') return null;
    if (v.type === 'list') return this.listItems(v).some((i) => i.toLowerCase() === t.toLowerCase()) ? null : 'list';
    const n = Number(t.replace(/,/g, '')); if (isNaN(n)) return 'number'; if (v.type === 'whole' && !Number.isInteger(n)) return 'whole';
    return n < v.min || n > v.max ? 'range' : null;
  }

  /* ── Remove Duplicates and Text to Columns ── */
  _box(range) { const [x, y] = range.split(':'), p = parseRef(x), q = parseRef(y || x); return { r1: Math.min(p.r, q.r), r2: Math.max(p.r, q.r), c1: Math.min(p.c, q.c), c2: Math.max(p.c, q.c) }; }
  /** Keeps the first row of every group that matches on the chosen columns (case-insensitive, spaces matter), moves the rest up and blanks the bottom of the range. cols: absolute column numbers, null = all. */
  removeDuplicates(range, cols, header) {
    const b = this._box(range), first = header ? b.r1 + 1 : b.r1, use = cols && cols.length ? cols : null, seen = new Set(), kept = [];
    let removed = 0;
    for (let r = first; r <= b.r2; r++) {
      const key = []; for (let c = b.c1; c <= b.c2; c++) if (!use || use.includes(c)) key.push(this.display(numToCol(c) + r).toLowerCase());
      const k = JSON.stringify(key);
      if (seen.has(k)) { removed++; continue; } seen.add(k);
      const row = []; for (let c = b.c1; c <= b.c2; c++) { const cell = this.cells[numToCol(c) + r]; row.push(cell ? Object.assign({}, cell) : null); } kept.push(row);
    }
    for (let r = first; r <= b.r2; r++) { const row = kept[r - first]; for (let c = b.c1; c <= b.c2; c++) { const ref = numToCol(c) + r, cell = row ? row[c - b.c1] : null; if (cell) this.cells[ref] = cell; else delete this.cells[ref]; } }
    this._inv(); return { removed, unique: kept.length };
  }
  /** Splits the text of one column into the columns to its right. delim is a string; the pieces are entered like typed values (numbers become numbers). */
  textToColumns(col, r1, r2, delim) {
    if (delim === '' || delim == null) return 0; let n = 0;
    for (let r = r1; r <= r2; r++) {
      const cell = this.cells[numToCol(col) + r]; if (!cell || cell.raw === '' || cell.raw[0] === '=') continue;
      const parts = (cell.raw[0] === "'" ? cell.raw.slice(1) : cell.raw).split(delim); if (parts.length < 2) continue;
      parts.forEach((p, i) => { const c = col + i; if (c > this.cols) return; const ref = numToCol(c) + r; if (i === 0) cell.raw = p; else this.commit(ref, p.trim() === p ? p : p.trim()); });
      if (parts[0] !== undefined) this.commit(numToCol(col) + r, parts[0]); n++;
    }
    this._inv(); return n;
  }

  /* ── Flash Fill: learn the pattern from the filled cells of a column, complete the empty ones ── */
  flashFill(col) {
    if (col < 2) return 0;
    const srcCols = [col - 1]; if (col > 2) srcCols.unshift(col - 2);
    const txt = (c, r) => { const ref = numToCol(c) + r, cell = this.cells[ref]; return cell && cell.raw !== '' ? (cell.raw[0] === "'" ? cell.raw.slice(1) : cell.raw[0] === '=' ? this.display(ref) : cell.raw) : ''; };
    const rowsWith = []; for (let r = 1; r <= this.rows; r++) if (txt(col - 1, r) !== '') rowsWith.push(r);
    const all = rowsWith.filter((r) => txt(col, r) !== '').map((r) => ({ r, out: txt(col, r) }));
    if (!all.length) return 0;
    const sources = (r, withTwo) => (withTwo ? srcCols.map((c) => txt(c, r)) : [txt(col - 1, r)]);
    const attempt = (examples) => {
      for (const withTwo of [true, false]) {
        if (withTwo && srcCols.length < 2) continue;
        const exs = examples.map((e) => ({ s: sources(e.r, withTwo), out: e.out })); if (withTwo && exs.some((e) => e.s.some((x) => x === '' || !e.out.includes(x)))) continue;
        for (const [, f] of flashCandidates(exs)) {
          let ok = true; for (const e of exs) { let v; try { v = f(...e.s); } catch (err2) { v = undefined; } if (v !== e.out) { ok = false; break; } }
          if (ok) return { f, withTwo };
        }
      }
      return null;
    };
    let hit = attempt(all); if (!hit && all[0].r === 1 && all.length > 1) hit = attempt(all.slice(1));
    if (!hit) return 0; let n = 0;
    for (const r of rowsWith) {
      if (txt(col, r) !== '') continue; const s = sources(r, hit.withTwo); if (hit.withTwo && s.some((x) => x === '')) continue;
      let v; try { v = hit.f(...s); } catch (err2) { v = undefined; } if (v === undefined || v === null) continue;
      this.setRaw(numToCol(col) + r, String(v)); n++;
    }
    this._inv(); return n;
  }
}

root.XL = { localName, nrm, Sheet, XlError, Rng, STYLE_KEYS, adjustFormula, parseRef, colToNum, numToCol, shiftFormula, autoClose, DOCS, formatValue, serialOf, dateOf, todaySerial, isDateRaw, canon, resolveName };
})(typeof window !== 'undefined' ? window : globalThis);
