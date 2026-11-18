/* Spreadsheet Academy spreadsheet view: keyboard-first grid with formula bar, name box, fill, autosum,
   copy/paste, undo and formula hints. window.Grid */
(function (root) {
'use strict';
const { numToCol, colToNum, parseRef, shiftFormula, DOCS } = root.XL;
const addr = (c, r) => numToCol(c) + r;
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const ARROWS = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };

function sigAt(text) {
  const st = []; let inS = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') { inS = !inS; continue; } if (inS) continue;
    if (ch === '(') { const m = /([A-Za-z][A-Za-z0-9.]*)$/.exec(text.slice(0, i)); st.push({ name: m ? m[1].toUpperCase() : '', i: 0 }); }
    else if ((ch === ',' || ch === ';') && st.length) st[st.length - 1].i++;
    else if (ch === ')') st.pop();
  }
  return st.length ? st[st.length - 1] : null;
}

/** the cell and range references in a formula being typed (outside quotes): [{ start, end, text, r1, c1, r2, c2 }] */
function refTokens(text, maxR, maxC) {
  const out = []; if (text[0] !== '=') return out;
  const inQ = []; let q = false; for (let i = 0; i < text.length; i++) { if (text[i] === '"') { q = !q; inQ[i] = true; } else inQ[i] = q; }
  const re = /(?<![A-Za-z0-9_.!'$])(\$?[A-Za-z]{1,3}\$?\d{1,7})(?::(\$?[A-Za-z]{1,3}\$?\d{1,7}))?(?![A-Za-z0-9_(])/g; let m;
  while ((m = re.exec(text))) {
    if (inQ[m.index]) continue;
    const one = (s) => { const mm = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(s); return { c: colToNum(mm[1].toUpperCase()), r: +mm[2] }; };
    const a = one(m[1]), b = m[2] ? one(m[2]) : a, r1 = Math.min(a.r, b.r), r2 = Math.max(a.r, b.r), c1 = Math.min(a.c, b.c), c2 = Math.max(a.c, b.c);
    if (r1 < 1 || c1 < 1 || r2 > maxR || c2 > maxC) continue;
    out.push({ start: m.index, end: m.index + m[0].length, text: m[0], r1, c1, r2, c2 });
  }
  return out;
}
const REF_COLORS = ['#2F6FDE', '#D6392B', '#8E3FC9', '#1E9A5C', '#E08A00', '#17A2B8'];

class Grid {
  constructor(host, o) {
    this.host = host; this.sheet = o.sheet; this.o = o;
    this.sheets = o.sheets || null; this.activeName = o.activeName || (this.sheets ? Object.keys(this.sheets)[0] : null);
    this.locked = !!o.locked; this.mouse = o.mouse !== false; this.nameBoxOn = o.nameBox !== false;
    this.cur = { r: 1, c: 1 }; this.ext = { r: 1, c: 1 };
    this.editing = false; this.mode = 'edit'; this.undo = []; this.redo = []; this.clip = null; this.drag = false;
    this.hint = { items: [], sel: 0, kind: null }; this.extra = null;
    this.ac = new AbortController();
    this.build();
    if (o.cur) { const p = parseRef(o.cur); if (p) { this.cur = { ...p }; this.ext = { ...p }; } }
    this.refresh(); this.paint();
  }

  /* ── DOM ── */
  build() {
    const { rows, cols } = this.sheet;
    let h = '<thead><tr><th class="corner" data-all="1"></th>';
    for (let c = 1; c <= cols; c++) h += `<th class="ch" data-col="${c}">${numToCol(c)}<i class="rz" data-rz="${c}"></i></th>`;
    h += '</tr></thead><tbody>';
    for (let r = 1; r <= rows; r++) {
      h += `<tr><th class="rh" data-row="${r}">${r}</th>`;
      for (let c = 1; c <= cols; c++) h += `<td data-r="${r}" data-c="${c}"><span></span></td>`;
      h += '</tr>';
    }
    h += '</tbody>';
    this.host.innerHTML = `<div class="xl" data-nomouse="${!this.mouse}">
      <div class="xl-bar">
        <input class="xl-name" spellcheck="false" aria-label="Name box" ${this.nameBoxOn ? '' : 'disabled'}>
        <span class="xl-fx" aria-hidden="true"><i>fx</i></span>
        <input class="xl-formula" spellcheck="false" autocomplete="off" aria-label="Formula bar">
      </div>
      <div class="xl-scroll" tabindex="0" role="grid" aria-label="Spreadsheet">
        <div class="xl-wrap"><table class="xl-table">${h}</table>
          <div class="xl-sel"></div><div class="xl-refs" aria-hidden="true"></div><button type="button" class="xl-dd" hidden tabindex="-1" aria-label="Open list">▾</button>
          <div class="xl-mirror" hidden aria-hidden="true"></div><input class="xl-editor" spellcheck="false" autocomplete="off" hidden aria-label="Cell editor">
          <div class="xl-hint" hidden></div>
        </div>
      </div>${this.sheets ? `<div class="xl-tabs" role="tablist">${Object.keys(this.sheets).map((n) => `<button class="xl-tab${n === this.activeName ? ' on' : ''}" data-name="${n.replace(/"/g, '&quot;')}" role="tab">${n}</button>`).join('')}</div>` : ''}</div>`;
    const q = (s) => this.host.querySelector(s);
    this.root = q('.xl'); this.scroller = q('.xl-scroll'); this.wrap = q('.xl-wrap'); this.table = q('.xl-table');
    this.selBox = q('.xl-sel'); this.refsBox = q('.xl-refs'); this.mirror = q('.xl-mirror'); this.ddBtn = q('.xl-dd'); this.editor = q('.xl-editor'); this.bar = q('.xl-formula'); this.nameBox = q('.xl-name'); this.hintBox = q('.xl-hint');
    this.table.style.minWidth = 44 + cols * 104 + 'px';
    this.td = []; this.table.querySelectorAll('td').forEach((td) => { const r = +td.dataset.r, c = +td.dataset.c; (this.td[r] || (this.td[r] = []))[c] = td; });
    this.ch = [...this.table.querySelectorAll('th.ch')]; this.rh = [...this.table.querySelectorAll('th.rh')];
    this.bind();
  }

  bind() {
    const sig = { signal: this.ac.signal };
    const tabs = this.root.querySelector('.xl-tabs');
    if (tabs) tabs.addEventListener('mousedown', (e) => { const b = e.target.closest('.xl-tab'); if (!b) return; e.preventDefault(); if (this.editing) this.commit(0, 0); this.switchTo(b.dataset.name); });
    this.scroller.addEventListener('keydown', (e) => this.onKey(e));
    this.ddBtn.addEventListener('mousedown', (e) => e.preventDefault());
    this.ddBtn.addEventListener('click', () => { this.focus(); this.ddList ? this.closeList() : this.openList(); });
    this.table.addEventListener('mousedown', (e) => this.onMouseDown(e));
    this.table.addEventListener('mouseover', (e) => { if (!this.drag) return; const td = e.target.closest('td'); if (td) { this.ext = { r: +td.dataset.r, c: +td.dataset.c }; this.paint(); } });
    this.table.addEventListener('dblclick', (e) => { const rz = e.target.closest('.rz'); if (rz && this.mouse) { this.autofitCols([+rz.dataset.rz]); return; } const td = e.target.closest('td'); if (td && this.mouse) this.beginEdit(this.sheet.raw(this.curRef()), 'edit'); });
    document.addEventListener('mouseup', () => { if (this.drag) { this.drag = false; this.emitSelect(); } }, sig);
    this.editor.addEventListener('keydown', (e) => this.onEditorKey(e));
    this.bar.addEventListener('keydown', (e) => this.onEditorKey(e));
    this.editor.addEventListener('input', () => this.onEditInput(this.editor));
    this.bar.addEventListener('input', () => this.onEditInput(this.bar));
    this.bar.addEventListener('focus', () => { if (!this.editing && !this.locked) this.beginEdit(this.sheet.raw(this.curRef()), 'edit', this.bar); });
    this.editor.addEventListener('blur', () => this.onBlur());
    this.bar.addEventListener('blur', () => this.onBlur());
    this.nameBox.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') { e.preventDefault(); this.goTo(this.nameBox.value); }
      else if (e.key === 'Escape') { this.nameBox.value = this.rangeStr(); this.focus(); }
    });
    this.nameBox.addEventListener('focus', () => this.nameBox.select());
    this.scroller.addEventListener('copy', (e) => this.onCopy(e, false));
    this.scroller.addEventListener('cut', (e) => this.onCopy(e, true));
    this.scroller.addEventListener('paste', (e) => this.onPaste(e));
  }
  destroy() { this.ac.abort(); this.closeDialog(); this.closeList(); this.host.innerHTML = ''; }
  focus() { this.scroller.focus({ preventScroll: true }); }
  setLocked(v) { this.locked = v; if (v && this.editing) this.cancelEdit(); this.root.classList.toggle('is-locked', v); }

  /* ── selection ── */
  curRef() { return addr(this.cur.c, this.cur.r); }
  bbox() { return { r1: Math.min(this.cur.r, this.ext.r), r2: Math.max(this.cur.r, this.ext.r), c1: Math.min(this.cur.c, this.ext.c), c2: Math.max(this.cur.c, this.ext.c) }; }
  rangeStr() { const b = this.bbox(); const a = addr(b.c1, b.r1), z = addr(b.c2, b.r2); return a === z ? a : `${a}:${z}`; }
  select(str) {
    const [a, z] = str.toUpperCase().split(':'); const p = parseRef(a), q = z ? parseRef(z) : p; if (!p || !q) return false;
    this.cur = this.clampPos(p); this.ext = this.clampPos(q); this.sel(); return true;
  }
  goTo(str) {
    const t = str.trim(), ref = this.sheet.nameRef(t);
    if (ref && !ref.includes('!') && this.select(ref)) { this.focus(); return true; }
    if (this.select(t)) { this.focus(); return true; }
    if (t && !this.locked && !this.extra) {
      this.pushUndo(); const bad = this.sheet.defineName(t, this.rangeStr());
      if (!bad) { this.afterChange('name'); this.focus(); this.toast(root.I18N ? root.I18N.t('Name {n} now refers to {r}.', { n: t, r: this.rangeStr() }) : 'Name ' + t + ' now refers to ' + this.rangeStr()); return true; }
      this.undo.pop(); this.toast(bad === 'cell' ? 'A name cannot look like a cell address such as A1.' : 'Names must start with a letter or underscore and contain no spaces.');
    }
    this.nameBox.value = this.rangeStr(); this.focus(); return false;
  }
  clampPos(p) { return { r: Math.max(1, Math.min(this.sheet.rows, p.r)), c: Math.max(1, Math.min(this.sheet.cols, p.c)) }; }
  sel() { this.extra = null; this.paint(); this.emitSelect(); }
  emitSelect() { this.o.onSelect && this.o.onSelect(this); }
  has(r, c) { return this.sheet.raw(addr(c, r)) !== ''; }
  inb(r, c) { return r >= 1 && c >= 1 && r <= this.sheet.rows && c <= this.sheet.cols; }
  jump(pos, dr, dc) {
    let { r, c } = pos; const nr = r + dr, nc = c + dc;
    if (!this.inb(nr, nc)) return pos;
    if (this.has(r, c) && this.has(nr, nc)) { while (this.inb(r + dr, c + dc) && this.has(r + dr, c + dc)) { r += dr; c += dc; } return { r, c }; }
    r = nr; c = nc;
    while (!this.has(r, c)) { if (!this.inb(r + dr, c + dc)) return { r, c }; r += dr; c += dc; }
    return { r, c };
  }
  region() {
    const { r, c } = this.cur; const seen = new Set(); const stack = [[r, c]]; let r1 = r, r2 = r, c1 = c, c2 = c; let any = false;
    const ok = (a, b) => this.inb(a, b) && this.has(a, b);
    if (!ok(r, c)) { for (let i = -1; i <= 1; i++) for (let j = -1; j <= 1; j++) if (ok(r + i, c + j)) { stack.length = 0; stack.push([r + i, c + j]); any = true; } if (!any) return null; }
    while (stack.length) {
      const [a, b] = stack.pop(); const k = a * 1000 + b; if (seen.has(k) || !ok(a, b)) continue; seen.add(k);
      r1 = Math.min(r1, a); r2 = Math.max(r2, a); c1 = Math.min(c1, b); c2 = Math.max(c2, b);
      for (let i = -1; i <= 1; i++) for (let j = -1; j <= 1; j++) if (i || j) stack.push([a + i, b + j]);
    }
    return { r1, r2, c1, c2 };
  }
  setBox(b) { this.cur = { r: b.r1, c: b.c1 }; this.ext = { r: b.r2, c: b.c2 }; }

  /* ── painting ── */
  refresh() {
    for (let r = 1; r <= this.sheet.rows; r++) for (let c = 1; c <= this.sheet.cols; c++) {
      const ref = addr(c, r), td = this.td[r][c], cell = this.sheet.cells[ref];
      td.firstChild.textContent = this.sheet.display(ref);
      td.dataset.k = this.sheet.kind(ref);
      const st = (k, v) => { if (v) td.dataset[k] = v === true ? '1' : v; else delete td.dataset[k]; };
      st('b', cell && cell.bold); st('i', cell && cell.italic); st('u', cell && cell.underline); st('al', cell && cell.align); st('fill', cell && cell.fill);
      st('fc', cell && cell.color); st('wrap', cell && cell.wrap); st('bd', cell && cell.border);
      td.hidden = this.sheet.hiddenCols.has(c);
      const cf = this.sheet.cfFor(ref); if (cf) td.dataset.cf = cf; else delete td.dataset.cf;
    }
    this.ch.forEach((th, i) => { const w = this.sheet.colW[i + 1]; th.style.width = w ? w + 'px' : ''; th.hidden = this.sheet.hiddenCols.has(i + 1); });
    this.root.dataset.fr = this.sheet.freeze.rows ? '1' : ''; this.root.dataset.fc = this.sheet.freeze.cols ? '1' : '';
    const trs = this.table.tBodies[0].rows; for (let r = 1; r <= this.sheet.rows; r++) trs[r - 1].hidden = this.sheet.hidden.has(r);
    if (!this.editing) this.bar.value = this.sheet.raw(this.curRef());
  }
  paint() {
    const b = this.bbox();
    for (let r = 1; r <= this.sheet.rows; r++) for (let c = 1; c <= this.sheet.cols; c++) {
      const td = this.td[r][c], inside = r >= b.r1 && r <= b.r2 && c >= b.c1 && c <= b.c2;
      td.classList.toggle('in', this.extra ? this.extra.has(addr(c, r)) : inside && !(b.r1 === b.r2 && b.c1 === b.c2));
      td.classList.toggle('cur', r === this.cur.r && c === this.cur.c);
    }
    this.ch.forEach((th, i) => th.classList.toggle('hl', i + 1 >= b.c1 && i + 1 <= b.c2));
    this.rh.forEach((th, i) => th.classList.toggle('hl', i + 1 >= b.r1 && i + 1 <= b.r2));
    const a = this.td[b.r1][b.c1], z = this.td[b.r2][b.c2], w = this.wrap.getBoundingClientRect(), ra = a.getBoundingClientRect(), rz = z.getBoundingClientRect();
    this.selBox.hidden = !!this.extra;
    Object.assign(this.selBox.style, { left: ra.left - w.left - 1 + 'px', top: ra.top - w.top - 1 + 'px', width: rz.right - ra.left + 1 + 'px', height: rz.bottom - ra.top + 1 + 'px' });
    this.updateDropdown(ra, w);
    this.root.classList.toggle('multi', b.r1 !== b.r2 || b.c1 !== b.c2);
    if (!this.editing) { this.nameBox.value = this.rangeStr(); this.bar.value = this.sheet.raw(this.curRef()); }
    this.reveal(this.td[this.cur.r][this.cur.c]);
  }
  reveal(td) {
    const s = this.scroller, a = td.getBoundingClientRect(), b = s.getBoundingClientRect(), hh = 30, hw = 46;
    if (a.bottom > b.bottom) s.scrollTop += a.bottom - b.bottom + 2; else if (a.top < b.top + hh) s.scrollTop -= b.top + hh - a.top + 2;
    if (a.right > b.right) s.scrollLeft += a.right - b.right + 2; else if (a.left < b.left + hw) s.scrollLeft -= b.left + hw - a.left + 2;
  }

  /* ── mouse ── */
  onMouseDown(e) {
    const el = e.target.closest('td,th'); if (!el) return;
    if (e.target.closest('.rz')) { e.preventDefault(); this.focus(); return this.startResize(+e.target.closest('.rz').dataset.rz, e); }
    this.extra = null;
    if (this.editing) {
      if (el.tagName === 'TD' && this.canInsertRef()) { e.preventDefault(); this.insertRef(addr(+el.dataset.c, +el.dataset.r)); return; }
      this.commit(0, 0);
    }
    e.preventDefault();
    if (!this.mouse) { this.focus(); this.o.onBlocked && this.o.onBlocked(); return; }
    this.focus();
    if (el.tagName === 'TD') {
      const p = { r: +el.dataset.r, c: +el.dataset.c };
      if (e.shiftKey) this.ext = p; else { this.cur = p; this.ext = { ...p }; }
      this.drag = true; this.paint();
    } else if (el.dataset.col) { const c = +el.dataset.col; this.cur = { r: 1, c }; this.ext = { r: this.sheet.rows, c }; this.sel(); }
    else if (el.dataset.row) { const r = +el.dataset.row; this.cur = { r, c: 1 }; this.ext = { r, c: this.sheet.cols }; this.sel(); }
    else if (el.dataset.all) { this.cur = { r: 1, c: 1 }; this.ext = { r: this.sheet.rows, c: this.sheet.cols }; this.sel(); }
  }

  /* ── keyboard (navigation mode) ── */
  onKey(e) {
    if (this.editing) return;
    const k = e.key, ctrl = e.ctrlKey || e.metaKey, sh = e.shiftKey, alt = e.altKey;
    const go = () => { e.preventDefault(); e.stopPropagation(); };
    if (this.ddList) {
      const els = [...this.ddList.children], at = els.findIndex((x) => x.classList.contains('on'));
      if (k === 'Escape') { go(); return this.closeList(); }
      if (k === 'ArrowDown' || k === 'ArrowUp') { go(); const n = (at + (k === 'ArrowDown' ? 1 : els.length - 1)) % els.length; els.forEach((x, i) => x.classList.toggle('on', i === n)); return; }
      if (k === 'Enter') { go(); if (at >= 0) els[at].click(); else this.closeList(); return; }
    }
    if (alt && k === 'ArrowDown' && !this.ddBtn.hidden) { go(); return this.openList(); }
    if (ARROWS[k]) {
      go(); const [dr, dc] = ARROWS[k];
      if (ctrl) { if (sh) this.ext = this.jump(this.ext, dr, dc); else { this.cur = this.jump(this.cur, dr, dc); this.ext = { ...this.cur }; } }
      else if (sh) this.ext = this.clampPos({ r: this.ext.r + dr, c: this.ext.c + dc });
      else { this.cur = this.skipHiddenCols(this.skipHidden(this.clampPos({ r: this.cur.r + dr, c: this.cur.c + dc }), dr), dc); this.ext = { ...this.cur }; }
      return this.sel();
    }
    if (k === 'Enter') { go(); return this.moveBy(sh ? -1 : 1, 0); }
    if (k === 'Tab') { go(); return this.moveBy(0, sh ? -1 : 1); }
    if (k === 'F2') { go(); return this.beginEdit(this.sheet.raw(this.curRef()), 'edit'); }
    if (k === 'F5' || (ctrl && !alt && k.toLowerCase() === 'g')) { go(); if (this.o.goto) return this.openGoTo(); if (this.nameBoxOn) this.nameBox.focus(); return; }
    if (ctrl && alt && k.toLowerCase() === 'v') { go(); return this.openPasteSpecial(); }
    if (k === 'Delete' || k === 'Backspace') { go(); return this.clearRange(); }
    if (k === 'Home') { go(); const p = ctrl ? { r: 1, c: 1 } : { r: this.cur.r, c: 1 }; if (sh) this.ext = p; else { this.cur = p; this.ext = { ...p }; } return this.sel(); }
    if (k === 'End' && ctrl) { go(); const p = this.sheet.usedEnd(); if (sh) this.ext = p; else { this.cur = p; this.ext = { ...p }; } return this.sel(); }
    if (ctrl && !alt) {
      const key = k.toLowerCase();
      if (sh && e.code === 'Digit4') { go(); return this.applyFmt('currency'); }
      if (sh && e.code === 'Digit5') { go(); return this.applyFmt('percent'); }
      if (sh && e.code === 'Digit3') { go(); return this.applyFmt('date'); }
      if (sh && e.code === 'Digit1') { go(); return this.applyFmt('comma'); }
      if (sh && e.code === 'Backquote') { go(); return this.applyFmt(null); }
      if (key === 'a') { go(); const reg = this.region(), b = this.bbox(); const same = reg && reg.r1 === b.r1 && reg.r2 === b.r2 && reg.c1 === b.c1 && reg.c2 === b.c2;
        this.setBox(reg && !same && (b.r1 === b.r2 && b.c1 === b.c2) ? reg : { r1: 1, r2: this.sheet.rows, c1: 1, c2: this.sheet.cols }); return this.sel(); }
      if (k === ' ') { go(); const c = this.cur.c; this.cur = { r: 1, c }; this.ext = { r: this.sheet.rows, c }; return this.sel(); }
      if (key === 'd') { go(); return this.fill('down'); }
      if (key === 'r') { go(); return this.fill('right'); }
      if (key === 'c' && !sh) { go(); return this.copyData(false); }
      if (key === 'x' && !sh) { go(); return this.copyData(true); }
      if (key === 'v' && !sh && this.clip) { go(); return this.pasteData(this.clip.text); }
      if ((k === 'PageDown' || k === 'PageUp') && this.sheets) { go(); const n = Object.keys(this.sheets), i = n.indexOf(this.activeName) + (k === 'PageDown' ? 1 : -1); if (i >= 0 && i < n.length) this.switchTo(n[i]); return; }
      if (key === 'z') { go(); return this.doUndo(); }
      if (key === 'y') { go(); return this.doRedo(); }
      if (key === 'b') { go(); return this.toggleBold(); }
      if (key === 'i') { go(); return this.toggleStyle('italic'); }
      if (key === 'u') { go(); return this.toggleStyle('underline'); }
      if (key === 'e') { go(); return this.flashFill(); }
      if (key === 'h') { go(); return this.openFind(true); }
      if (key === 'f') { go(); return this.openFind(false); }
      if (sh && (e.code === 'Equal' || k === '+')) { go(); return this.insertEntire(); }
      if (!sh && (e.code === 'Minus' || e.code === 'NumpadSubtract')) { go(); return this.deleteEntire(); }
      if (e.code === 'Digit9') { go(); return sh ? this.unhide('rows') : this.hide('rows'); }
      if (e.code === 'Digit0') { go(); return sh ? this.unhide('cols') : this.hide('cols'); }
      return; // let copy / cut / paste through
    }
    if (alt && (k === '=' || e.code === 'Equal')) { go(); return this.autoSum(); }
    if (sh && k === ' ' && !ctrl) { go(); const r = this.cur.r; this.cur = { r, c: 1 }; this.ext = { r, c: this.sheet.cols }; return this.sel(); }
    if (k.length === 1 && !ctrl && !alt && !e.metaKey) { go(); if (!this.locked) this.beginEdit(k, 'enter'); else this.o.onBlocked && this.o.onBlocked('locked'); }
  }
  skipHiddenCols(p, dc) { const d = dc || 1; while (this.sheet.hiddenCols.has(p.c) && p.c + d >= 1 && p.c + d <= this.sheet.cols) p = { r: p.r, c: p.c + d }; return p; }
  skipHidden(p, dr) { const d = dr || 1; while (this.sheet.hidden.has(p.r) && p.r + d >= 1 && p.r + d <= this.sheet.rows) p = { r: p.r + d, c: p.c }; return p; }
  moveBy(dr, dc) { this.cur = this.skipHidden(this.clampPos({ r: this.cur.r + dr, c: this.cur.c + dc }), dr); this.ext = { ...this.cur }; this.sel(); }

  /* ── editing ── */
  beginEdit(text, mode, src) {
    if (this.locked) { this.o.onBlocked && this.o.onBlocked('locked'); return; }
    this.editing = true; this.mode = mode;
    const td = this.td[this.cur.r][this.cur.c], w = this.wrap.getBoundingClientRect(), t = td.getBoundingClientRect();
    Object.assign(this.editor.style, { left: t.left - w.left - 1 + 'px', top: t.top - w.top - 1 + 'px', minWidth: t.width + 1 + 'px', height: t.height + 1 + 'px' });
    this.editor.hidden = false; this.editor.value = text; this.bar.value = text; this.fitEditor();
    const input = src || this.editor; input.focus();
    if (mode === 'edit') input.setSelectionRange(text.length, text.length);
    this.root.classList.add('is-editing'); this.updateHint(); this.paintRefs();
  }
  endEdit() { this.editing = false; this.paintRefs(); this.editor.hidden = true; this.hintBox.hidden = true; this.hint.items = []; this.root.classList.remove('is-editing'); }
  /** like Excel: while a formula is typed, every cell or range it refers to gets a coloured border on the grid, and the same colour in the formula text */
  paintRefs() {
    const box = this.refsBox, mir = this.mirror, ed = this.editor; box.innerHTML = '';
    const text = this.editing ? ed.value : '';
    if (!this.editing || text[0] !== '=') { mir.hidden = true; mir.innerHTML = ''; ed.classList.remove('refmode'); return; }
    const toks = refTokens(text, this.td.length - 1, (this.td[1] || []).length - 1), w = this.wrap.getBoundingClientRect(), colour = {};
    toks.forEach((t) => {
      const key = t.text.toUpperCase().replace(/\$/g, ''); if (!(key in colour)) colour[key] = REF_COLORS[Object.keys(colour).length % REF_COLORS.length];
      t.col = colour[key];
      const a = this.td[t.r1][t.c1].getBoundingClientRect(), b = this.td[t.r2][t.c2].getBoundingClientRect(), d = document.createElement('div');
      d.className = 'xl-ref'; d.style.cssText = '--rc:' + t.col + ';left:' + (a.left - w.left - 1) + 'px;top:' + (a.top - w.top - 1) + 'px;width:' + (b.right - a.left + 1) + 'px;height:' + (b.bottom - a.top + 1) + 'px';
      box.appendChild(d);
    });
    let html = '', at = 0; toks.forEach((t) => { html += esc(text.slice(at, t.start)) + '<span style="color:' + t.col + '">' + esc(t.text) + '</span>'; at = t.end; }); html += esc(text.slice(at));
    mir.innerHTML = html; ['left', 'top', 'width', 'height', 'minWidth'].forEach((k) => { mir.style[k] = ed.style[k]; });
    mir.hidden = false; ed.classList.toggle('refmode', toks.length > 0 && !ed.hidden);
    if (!toks.length) mir.hidden = true;
  }
  fitEditor() {
    this.editor.style.width = Math.max(this.td[this.cur.r][this.cur.c].offsetWidth + 1, this.editor.value.length * 8.5 + 26) + 'px';
    if (this.editor.scrollWidth <= this.editor.clientWidth) this.editor.scrollLeft = 0;
  }
  cancelEdit() { if (!this.editing) return; this.endEdit(); this.bar.value = this.sheet.raw(this.curRef()); this.focus(); }
  commit(dr, dc) {
    if (!this.editing) return;
    const text = this.editor.value; this.endEdit(); const ref = this.curRef();
    if (this.sheet.validate(ref, text)) { window.Sound && window.Sound.play('blocked'); return this.rejectEntry(text); }
    window.Sound && window.Sound.play('enter');
    this.pushUndo(); this.sheet.commit(ref, text); this.refresh();
    if (dr || dc) this.moveBy(dr, dc); else this.paint();
    this.focus(); this.o.onChange && this.o.onChange({ kind: 'commit', ref }); this.emitSelect();
  }
  onBlur() { setTimeout(() => { if (this.editing && document.activeElement !== this.editor && document.activeElement !== this.bar) this.commit(0, 0); }, 0); }
  onEditInput(src) {
    const other = src === this.editor ? this.bar : this.editor; other.value = src.value; this.fitEditor(); this.updateHint(); this.paintRefs();
  }
  activeInput() { return document.activeElement === this.bar ? this.bar : this.editor; }
  onEditorKey(e) {
    const k = e.key, inp = e.currentTarget; e.stopPropagation();
    if (!this.editing) { if (k === 'Escape') this.focus(); return; }
    const listOpen = this.hint.kind === 'list' && this.hint.items.length;
    if (listOpen && (k === 'ArrowDown' || k === 'ArrowUp')) { e.preventDefault(); const n = this.hint.items.length; this.hint.sel = (this.hint.sel + (k === 'ArrowDown' ? 1 : n - 1)) % n; this.renderHint(); return; }
    if (k === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); this.commitAll(); }
    else if (k === 'Enter') { e.preventDefault(); this.commit(e.shiftKey ? -1 : 1, 0); }
    else if (k === 'Tab') { e.preventDefault(); if (listOpen) this.acceptHint(); else this.commit(0, e.shiftKey ? -1 : 1); }
    else if (k === 'Escape') { e.preventDefault(); this.cancelEdit(); }
    else if (ARROWS[k] && !e.shiftKey && !e.ctrlKey) {
      if (this.mode === 'enter' && inp.value[0] !== '=') { e.preventDefault(); const [dr, dc] = ARROWS[k]; this.commit(dr, dc); }
      else this.mode = 'edit';
    } else if (k === 'F4') { e.preventDefault(); this.cycleAbs(inp); }
    else if (k === 'ArrowUp' || k === 'ArrowDown') { /* default caret behaviour */ }
  }
  canInsertRef() { const t = this.editor.value; if (t[0] !== '=') return false; const pos = this.activeInput().selectionStart; return /[=+\-*\/^&(,<>:]$/.test(t.slice(0, pos)); }
  insertRef(ref) {
    const inp = this.activeInput(), s = inp.selectionStart, e = inp.selectionEnd;
    inp.setRangeText(ref, s, e, 'end'); this.onEditInput(inp); inp.focus();
  }
  cycleAbs(inp) {
    const t = inp.value, pos = inp.selectionStart; const re = /\$?[A-Za-z]{1,3}\$?\d+/g; let m;
    while ((m = re.exec(t))) {
      if (pos >= m.index && pos <= m.index + m[0].length) {
        const mm = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(m[0]); const abs = [mm[1] && mm[3], !mm[1] && mm[3], mm[1] && !mm[3]];
        const next = abs[0] ? ['', '$'] : abs[1] ? ['$', ''] : abs[2] ? ['', ''] : ['$', '$'];
        const rep = next[0] + mm[2] + next[1] + mm[4];
        inp.setRangeText(rep, m.index, m.index + m[0].length, 'end'); this.onEditInput(inp); return;
      }
    }
  }

  /* ── formula hints ── */
  updateHint() {
    const inp = this.activeInput(), t = inp.value; this.hint.items = []; this.hint.kind = null;
    if (!this.editing || t[0] !== '=') { this.hintBox.hidden = true; return; }
    const pos = inp.selectionStart, before = t.slice(0, pos), after = t.slice(pos);
    const m = /(?:^=|[^A-Za-z0-9_.$"])([A-Za-z][A-Za-z0-9.]*)$/.exec(before);
    if (m && !after.startsWith('(')) {
      const lg = this.lang(), XL = root.XL, up = XL.nrm(m[1]), names = Object.keys(DOCS);
      let list = lg === 'en' ? names.filter((n) => n.startsWith(up)) : names.filter((n) => XL.nrm(XL.localName(n, lg)).startsWith(up));
      if (lg !== 'en') { if (!list.length) list = names.filter((n) => n.startsWith(up)); else list.sort((a, b) => XL.localName(a, lg).localeCompare(XL.localName(b, lg))); }
      list = list.slice(0, 6);
      if (list.length) { this.hint = { items: list, sel: 0, kind: 'list', typed: m[1] }; return this.renderHint(); }
    }
    const s = sigAt(before);
    const rn = s && root.XL.resolveName(s.name);
    if (s && DOCS[rn]) { this.hint = { items: [], sel: 0, kind: 'sig', fn: rn, idx: s.i }; return this.renderHint(); }
    this.hintBox.hidden = true;
  }
  /** the app language: 'en' | 'pt' | 'es' */
  lang() { return (root.I18N && root.I18N.lang) || 'en'; }
  /** a function name as the learner sees it (SOMA in Português, SUMA in Español) */
  fname(n) { const lg = this.lang(); return lg === 'en' ? n : root.XL.localName(n, lg); }
  /** an argument such as value_if_true shown as valor_se_verdadeiro; digits and brackets are kept */
  farg(a) { const lg = this.lang(), M = root.FNARGS && root.FNARGS[lg]; if (!M) return a; return a.replace(/[a-z_]+/g, (w) => M[w] || w); }
  renderHint() {
    const h = this.hint, box = this.hintBox, lg = this.lang(), tr = (s) => (lg === 'en' || !root.I18N ? s : root.I18N.tx(s)), sep = lg === 'en' ? ', ' : '; ';
    if (h.kind === 'list') {
      box.innerHTML = h.items.map((n, i) => `<div class="hi${i === h.sel ? ' on' : ''}"><b>${this.fname(n)}</b><span>${esc(tr(DOCS[n][1]))}</span></div>`).join('') + `<div class="hk">${esc(tr('Tab to complete'))}</div>`;
    } else {
      const [sig, desc] = DOCS[h.fn], open = sig.indexOf('('), inner = sig.slice(open + 1, sig.lastIndexOf(')')).split(', ');
      const hi = Math.min(h.idx, inner.length - 1);
      box.innerHTML = `<div class="hs">${this.fname(h.fn)}(${inner.map((a, i) => { const s = this.farg(a); return i === hi || (a.includes('…') && h.idx >= i) ? `<b>${esc(s)}</b>` : esc(s); }).join(sep)})</div><div class="hd">${esc(tr(desc))}</div>`;
    }
    const t = this.td[this.cur.r][this.cur.c].getBoundingClientRect();
    box.hidden = false; const bh = box.offsetHeight, bw = box.offsetWidth;
    const top = t.bottom + 6 + bh > window.innerHeight - 8 ? Math.max(8, t.top - bh - 6) : t.bottom + 6;
    box.style.left = Math.max(8, Math.min(t.left, window.innerWidth - bw - 8)) + 'px'; box.style.top = top + 'px';
  }
  acceptHint() {
    const inp = this.activeInput(), pos = inp.selectionStart, name = this.hint.items[this.hint.sel], typed = this.hint.typed;
    inp.setRangeText(this.fname(name) + '(', pos - typed.length, pos, 'end'); this.onEditInput(inp);
  }

  /* ── data operations ── */
  pushUndo() { this.undo.push(this.sheet.snapshot()); if (this.undo.length > 60) this.undo.shift(); this.redo.length = 0; }
  doUndo() { if (this.locked || !this.undo.length) return; this.redo.push(this.sheet.snapshot()); this.sheet.restore(this.undo.pop()); this.afterChange(); }
  doRedo() { if (this.locked || !this.redo.length) return; this.undo.push(this.sheet.snapshot()); this.sheet.restore(this.redo.pop()); this.afterChange(); }
  afterChange(kind = 'edit') { this.refresh(); this.paint(); this.o.onChange && this.o.onChange({ kind }); this.emitSelect(); }
  eachCell(fn) { if (this.extra) { for (const ref of this.extra) { const p = parseRef(ref); fn(p.r, p.c, ref); } return; } const b = this.bbox(); for (let r = b.r1; r <= b.r2; r++) for (let c = b.c1; c <= b.c2; c++) fn(r, c, addr(c, r)); }
  clearRange() {
    if (this.locked) return this.o.onBlocked && this.o.onBlocked('locked');
    this.pushUndo(); this.eachCell((r, c, ref) => this.sheet.setRaw(ref, '')); this.afterChange();
  }
  applyFmt(f) { if (this.locked) return; this.pushUndo(); this.eachCell((r, c, ref) => this.sheet.setFmt(ref, f)); this.afterChange('format'); }
  toggleBold() { if (this.locked) return; const on = !(this.sheet.cells[this.curRef()] || {}).bold; this.pushUndo(); this.eachCell((r, c, ref) => this.sheet.setBold(ref, on)); this.afterChange('format'); }
  copyCell(sr, sc, dr, dc) {
    const sref = addr(sc, sr), dref = addr(dc, dr), raw = this.sheet.raw(sref);
    this.sheet.setRaw(dref, raw[0] === '=' ? shiftFormula(raw, dr - sr, dc - sc) : raw); this.sheet.copyStyle(sref, dref);
  }
  fill(dir) {
    if (this.locked) return this.o.onBlocked && this.o.onBlocked('locked');
    const b = this.bbox(); this.pushUndo();
    if (dir === 'down') {
      for (let c = b.c1; c <= b.c2; c++) {
        if (b.r1 === b.r2) { if (b.r1 > 1) this.copyCell(b.r1 - 1, c, b.r1, c); }
        else for (let r = b.r1 + 1; r <= b.r2; r++) this.copyCell(b.r1, c, r, c);
      }
    } else {
      for (let r = b.r1; r <= b.r2; r++) {
        if (b.c1 === b.c2) { if (b.c1 > 1) this.copyCell(r, b.c1 - 1, r, b.c1); }
        else for (let c = b.c1 + 1; c <= b.c2; c++) this.copyCell(r, b.c1, r, c);
      }
    }
    this.afterChange('fill');
  }
  autoSum() {
    if (this.locked) return this.o.onBlocked && this.o.onBlocked('locked');
    const b = this.bbox(), isNum = (r, c) => this.inb(r, c) && typeof this.sheet.value(addr(c, r)) === 'number';
    if (b.r1 === b.r2 && b.c1 === b.c2) {
      const { r, c } = this.cur; let rr = r - 1; while (rr >= 1 && isNum(rr, c)) rr--;
      if (rr < r - 1) return this.beginEdit(`=SUM(${addr(c, rr + 1)}:${addr(c, r - 1)})`, 'edit');
      let cc = c - 1; while (cc >= 1 && isNum(r, cc)) cc--;
      if (cc < c - 1) return this.beginEdit(`=SUM(${addr(cc + 1, r)}:${addr(c - 1, r)})`, 'edit');
      return this.beginEdit('=SUM(', 'edit');
    }
    if (b.r2 + 1 > this.sheet.rows) return;
    this.pushUndo();
    for (let c = b.c1; c <= b.c2; c++) this.sheet.setRaw(addr(c, b.r2 + 1), `=SUM(${addr(c, b.r1)}:${addr(c, b.r2)})`);
    this.afterChange('fill');
  }
  copyData(cut) {
    const b = this.bbox(), disp = [];
    for (let r = b.r1; r <= b.r2; r++) { const dd = []; for (let c = b.c1; c <= b.c2; c++) dd.push(this.sheet.display(addr(c, r))); disp.push(dd.join('\t')); }
    const block = this.sheet.copyBlock(b.r1, b.c1, b.r2, b.c2), rows = block.rows;
    const text = disp.join('\n'); this.clip = { text, rows, block, cut: cut && !this.locked, box: b };
    try { if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).catch(() => {}); } catch (e) { /* clipboard unavailable */ }
    if (this.clip.cut) { this.pushUndo(); this.eachCell((r, c, ref) => this.sheet.setRaw(ref, '')); this.afterChange('cut'); }
  }
  pasteData(text) {
    if (this.locked) return this.o.onBlocked && this.o.onBlocked('locked');
    this.pushUndo();
    if (this.clip && this.clip.text === text) {
      const b = this.bbox(), one = this.clip.rows.length === 1 && this.clip.rows[0].length === 1;
      if (one && (b.r1 !== b.r2 || b.c1 !== b.c2)) { for (let r = b.r1; r <= b.r2; r++) for (let c = b.c1; c <= b.c2; c++) this.sheet.pasteBlock(this.clip.block, r, c); }
      else this.sheet.pasteBlock(this.clip.block, this.cur.r, this.cur.c);
    } else {
      text.replace(/\r/g, '').replace(/\n$/, '').split('\n').forEach((line, i) => line.split('\t').forEach((v, j) => {
        const r = this.cur.r + i, c = this.cur.c + j; if (this.inb(r, c)) this.sheet.commit(addr(c, r), v);
      }));
    }
    this.afterChange('paste');
  }
  onCopy(e, cut) { if (this.editing) return; e.preventDefault(); this.copyData(cut); e.clipboardData.setData('text/plain', this.clip.text); }
  onPaste(e) { if (this.editing) return; e.preventDefault(); this.pasteData(e.clipboardData.getData('text/plain')); }
  /* ── styles and structure (also driven by the ribbon) ── */
  setStyle(key, val) { if (this.locked) return; this.pushUndo(); this.eachCell((r, c, ref) => this.sheet.setStyle(ref, key, val)); this.afterChange('format'); }
  toggleStyle(key) { if (this.locked) return; const on = !this.sheet.style(this.curRef(), key); this.setStyle(key, on ? true : false); }
  clearFormats() { if (this.locked) return; this.pushUndo(); this.eachCell((r, c, ref) => this.sheet.clearStyles(ref)); this.afterChange('format'); }
  autofitCols(cols) { if (this.locked) return; this.pushUndo(); cols.forEach((c) => this.sheet.autofit(c)); this.afterChange('format'); }
  autofitSel() { const b = this.bbox(), cols = []; for (let c = b.c1; c <= b.c2; c++) cols.push(c); this.autofitCols(cols); }
  startResize(col, e) {
    const th = this.ch[col - 1], start = th.getBoundingClientRect().width, x0 = e.clientX; this.pushUndo();
    const mv = (ev) => { this.sheet.colW[col] = Math.max(40, Math.round(start + ev.clientX - x0)); th.style.width = this.sheet.colW[col] + 'px'; this.paint(); };
    const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); this.afterChange('format'); };
    document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
  }
  entireKind() { const b = this.bbox(); if (b.c1 === 1 && b.c2 === this.sheet.cols) return 'rows'; if (b.r1 === 1 && b.r2 === this.sheet.rows) return 'cols'; return null; }
  structure(op, kind) {
    const b = this.bbox(), rows = kind === 'rows'; this.pushUndo();
    const at = rows ? b.r1 : b.c1, n = rows ? b.r2 - b.r1 + 1 : b.c2 - b.c1 + 1;
    if (op === 'insert') (rows ? this.sheet.insertRows(at, n) : this.sheet.insertCols(at, n)); else (rows ? this.sheet.deleteRows(at, n) : this.sheet.deleteCols(at, n));
    this.extra = null; this.afterChange('structure');
  }
  insertEntire() { if (this.locked) return this.o.onBlocked && this.o.onBlocked('locked'); const k = this.entireKind(); if (k) return this.structure('insert', k); this.chooseStructure('Insert', 'insert'); }
  deleteEntire() { if (this.locked) return this.o.onBlocked && this.o.onBlocked('locked'); const k = this.entireKind(); if (k) return this.structure('delete', k); this.chooseStructure('Delete', 'delete'); }
  chooseStructure(title, op) {
    this.dialog({ title, body: '<label class="rb-rad"><input type="radio" name="k" value="rows" checked> Entire row</label><label class="rb-rad"><input type="radio" name="k" value="cols"> Entire column</label>',
      ok: (d) => { this.structure(op, d.querySelector('[name=k]:checked').value); return true; } });
  }
  hide(kind) {
    if (this.locked) return; const b = this.bbox(); this.pushUndo();
    if (kind === 'rows') { this.sheet.hideRows(b.r1, b.r2); this.cur = this.skipHidden({ r: this.cur.r, c: this.cur.c }, 1); this.ext = { ...this.cur }; } else { this.sheet.hideCols(b.c1, b.c2); this.cur = this.skipHiddenCols({ r: this.cur.r, c: this.cur.c }, 1); this.ext = { ...this.cur }; }
    this.afterChange('hide');
  }
  unhide(kind) { if (this.locked) return; const b = this.bbox(); this.pushUndo(); if (kind === 'rows') this.sheet.showRows(b.r1, b.r2); else this.sheet.showCols(b.c1, b.c2); this.afterChange('hide'); }
  setFreeze(mode) {
    this.pushUndo(); const f = this.sheet.freeze;
    if (mode === 'none') { f.rows = 0; f.cols = 0; } else if (mode === 'row') f.rows = 1; else if (mode === 'col') f.cols = 1; else if (mode === 'both') { f.rows = this.cur.r > 1 ? this.cur.r - 1 : 1; f.cols = this.cur.c > 1 ? this.cur.c - 1 : 1; }
    this.afterChange('freeze');
  }
  selectRefs(refs) {
    if (!refs.length) return false; const ps = refs.map((x) => ({ x, p: parseRef(x) })).sort((a, z) => a.p.r - z.p.r || a.p.c - z.p.c);
    this.extra = new Set(refs); this.cur = { ...ps[0].p }; this.ext = { ...ps[0].p }; this.paint(); this.emitSelect(); return true;
  }
  /** Ctrl+Enter: put the typed entry into every selected cell (formulas adjust like a copy would) */
  commitAll() {
    if (!this.editing) return; const text = this.editor.value, from = { ...this.cur }; this.endEdit(); window.Sound && window.Sound.play('enter'); this.pushUndo();
    this.eachCell((r, c, ref) => this.sheet.commit(ref, text[0] === '=' ? shiftFormula(text, r - from.r, c - from.c) : text));
    this.refresh(); this.paint(); this.focus(); this.o.onChange && this.o.onChange({ kind: 'commit', ref: this.curRef() }); this.emitSelect();
  }
  flashFill() {
    if (this.locked) return this.o.onBlocked && this.o.onBlocked('locked');
    this.pushUndo(); const n = this.sheet.flashFill(this.cur.c);
    if (!n) { this.undo.pop(); return this.toast('Flash Fill needs a pattern: type the first result next to your data, then try again.'); }
    this.afterChange('flash');
  }
  toast(msg) { this.o.onBlocked && this.o.onBlocked('info', msg); }

  /* ── small Excel-style dialogs ── */
  dialog({ title, body, ok, okLabel, cancelLabel, extraButtons }) {
    this.closeDialog(); const T = (x) => (root.I18N ? root.I18N.t(x) : x);
    const ov = document.createElement('div'); ov.className = 'rb-dov xl-dov';
    ov.innerHTML = `<div class="rb-dlg" role="dialog" aria-modal="true"><div class="rb-dt"><span>${esc(T(title))}</span><button type="button" class="rb-dx" aria-label="Close">×</button></div><div class="rb-db">${body}</div><div class="rb-err" aria-live="polite"></div>
      <div class="rb-df">${(extraButtons || []).map((b, i) => `<button type="button" class="btn btn-secondary btn-sm" data-x="${i}">${esc(T(b.label))}</button>`).join('')}${ok ? `<button type="button" class="btn btn-primary btn-sm" data-ok>${esc(T(okLabel || 'OK'))}</button>` : ''}<button type="button" class="btn btn-secondary btn-sm" data-cancel>${esc(T(cancelLabel || 'Cancel'))}</button></div></div>`;
    document.body.appendChild(ov); this.dlgEl = ov;
    const first = ov.querySelector('input[type=text],input[type=search],input:not([type]),input:checked,button[data-ok]'); if (first) first.focus();
    const err = (m) => { ov.querySelector('.rb-err').textContent = m ? T(m) : ''; };
    const close = () => { ov.remove(); if (this.dlgEl === ov) this.dlgEl = null; this.focus(); };
    const go = () => { if (!ok) return close(); const r = ok(ov, err); if (r === true) close(); else if (r) err(r); };
    const okb = ov.querySelector('[data-ok]'); if (okb) okb.onclick = go;
    ov.querySelector('[data-cancel]').onclick = close; ov.querySelector('.rb-dx').onclick = close;
    ov.querySelectorAll('[data-x]').forEach((b) => { b.onclick = () => (extraButtons[+b.dataset.x].run(ov, err, close)); });
    ov.onkeydown = (e) => { e.stopPropagation(); if (e.key === 'Escape') close(); else if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); go(); } };
    return ov;
  }
  closeDialog() { if (this.dlgEl) { this.dlgEl.remove(); this.dlgEl = null; } }
  scopeRange() { const b = this.bbox(); return b.r1 === b.r2 && b.c1 === b.c2 ? null : this.rangeStr(); }
  openFind(replace) {
    if (this.locked) return this.o.onBlocked && this.o.onBlocked('locked');
    const rows = replace ? '<div class="rb-drow"><span class="xl-lbl">Replace with:</span><input type="text" data-r size="22" aria-label="Replace with"></div>' : '';
    this.dialog({ title: replace ? 'Find and Replace' : 'Find', cancelLabel: 'Close',
      body: `<div class="rb-drow"><span class="xl-lbl">Find what:</span><input type="text" data-f size="22" aria-label="Find what"></div>${rows}<label class="rb-rad"><input type="checkbox" data-case> Match case</label><label class="rb-rad"><input type="checkbox" data-whole> Match entire cell contents</label><p class="xl-res" aria-live="polite"></p>`,
      extraButtons: [
        ...(replace ? [{ label: 'Replace All', run: (d, err) => {
          const f = d.querySelector('[data-f]').value; if (f === '') return err('Type what to find.');
          this.pushUndo(); const n = this.sheet.replaceAll(f, d.querySelector('[data-r]').value, { matchCase: d.querySelector('[data-case]').checked, whole: d.querySelector('[data-whole]').checked, range: this.scopeRange() });
          if (!n) this.undo.pop(); else this.afterChange('replace'); err(''); d.querySelector('.xl-res').textContent = (root.I18N ? root.I18N.t('{n} replacements made.', { n }) : n + ' replacements made.');
        } }] : []),
        { label: 'Find All', run: (d, err) => { const hits = this.findHits(d); if (!hits) return; d.querySelector('.xl-res').textContent = (root.I18N ? root.I18N.t('{n} cells found.', { n: hits.length }) : hits.length + ' cells found.'); if (hits.length) this.selectRefs(hits); } },
        { label: 'Find Next', run: (d) => { const hits = this.findHits(d); if (!hits) return; const here = this.cur.r * 1000 + this.cur.c, nx = hits.find((h) => { const p = parseRef(h); return p.r * 1000 + p.c > here; }) || hits[0]; d.querySelector('.xl-res').textContent = hits.length ? '' : (root.I18N ? root.I18N.t('Nothing found.') : 'Nothing found.'); if (nx) { const p = parseRef(nx); this.cur = { ...p }; this.ext = { ...p }; this.sel(); } } },
      ] });
  }
  findHits(d) { const f = d.querySelector('[data-f]').value; if (f === '') return null; return this.sheet.findAll(f, { matchCase: d.querySelector('[data-case]').checked, whole: d.querySelector('[data-whole]').checked, range: this.scopeRange() }); }
  openGoTo() {
    this.dialog({ title: 'Go To', body: '<div class="rb-drow"><span class="xl-lbl">Reference:</span><input type="text" data-ref size="18" aria-label="Reference"></div>', okLabel: 'OK',
      extraButtons: [{ label: 'Special…', run: (d, err, close) => { close(); this.openSpecial(); } }],
      ok: (d) => { const v = d.querySelector('[data-ref]').value.trim(); if (!v) return 'Type a cell or range, such as B2:D9.'; return this.select(v) ? true : 'That is not a valid reference.'; } });
  }
  openSpecial() {
    const opt = (v, t, chk) => `<label class="rb-rad"><input type="radio" name="sp" value="${v}"${chk ? ' checked' : ''}> ${t}</label>`;
    this.dialog({ title: 'Go To Special', body: '<p>Select:</p>' + opt('constants', 'Constants', true) + opt('formulas', 'Formulas') + opt('blanks', 'Blanks') + opt('errors', 'Errors'),
      ok: (d) => { const kind = d.querySelector('[name=sp]:checked').value, refs = this.sheet.selectSpecial(kind, this.scopeRange()); if (!refs.length) return 'No cells were found.'; this.selectRefs(refs); return true; } });
  }
  openRemoveDuplicates() {
    if (this.locked) return this.o.onBlocked && this.o.onBlocked('locked');
    let b = this.bbox(); if (b.r1 === b.r2 && b.c1 === b.c2) { const reg = this.region(); if (!reg) return this.toast('Select the table first.'); b = reg; }
    const range = addr(b.c1, b.r1) + ':' + addr(b.c2, b.r2), cols = []; for (let c = b.c1; c <= b.c2; c++) cols.push(c);
    const label = (c, hdr) => (hdr ? esc(this.sheet.display(addr(c, b.r1)) || 'Column ' + numToCol(c)) : 'Column ' + numToCol(c));
    const draw = (d) => { const hdr = d.querySelector('[data-hdr]').checked; d.querySelector('.xl-cols').innerHTML = cols.map((c) => `<label class="rb-rad"><input type="checkbox" data-col="${c}" checked> ${label(c, hdr)}</label>`).join(''); };
    const ov = this.dialog({ title: 'Remove Duplicates', okLabel: 'OK',
      body: '<p>To delete duplicate values, select one or more columns that contain duplicates.</p><label class="rb-rad"><input type="checkbox" data-hdr checked> My data has headers</label><div class="xl-cols"></div>',
      ok: (d, err) => {
        const use = [...d.querySelectorAll('[data-col]:checked')].map((x) => +x.dataset.col); if (!use.length) return 'Select at least one column.';
        this.pushUndo(); const res = this.sheet.removeDuplicates(range, use.length === cols.length ? null : use, d.querySelector('[data-hdr]').checked); this.afterChange('dedupe');
        this.toast(root.I18N ? root.I18N.t('{a} duplicate values found and removed; {b} unique values remain.', { a: res.removed, b: res.unique }) : res.removed + ' duplicates removed; ' + res.unique + ' unique values remain.'); return true;
      } });
    draw(ov); ov.querySelector('[data-hdr]').onchange = () => draw(ov);
  }
  openTextToColumns() {
    if (this.locked) return this.o.onBlocked && this.o.onBlocked('locked');
    const b = this.bbox(); if (b.c1 !== b.c2) return this.toast('Select a single column of text first.');
    const opt = (v, t, chk) => `<label class="rb-rad"><input type="radio" name="dl" value="${v}"${chk ? ' checked' : ''}> ${t}</label>`;
    this.dialog({ title: 'Convert Text to Columns', okLabel: 'Finish',
      body: '<p>Choose the character that separates your data:</p>' + opt(',', 'Comma', true) + opt(';', 'Semicolon') + opt(' ', 'Space') + opt('\t', 'Tab') + '<label class="rb-rad"><input type="radio" name="dl" value="other"> Other: <input type="text" data-oth size="3" aria-label="Other"></label>',
      ok: (d) => {
        let dl = d.querySelector('[name=dl]:checked').value; if (dl === 'other') dl = d.querySelector('[data-oth]').value; if (!dl) return 'Type the separator character.';
        this.pushUndo(); const n = this.sheet.textToColumns(b.c1, b.r1, b.r2, dl); if (!n) { this.undo.pop(); return 'The separator was not found in the selected cells.'; } this.afterChange('split'); return true;
      } });
  }
  /* ── drop-down lists (Data Validation) ── */
  updateDropdown(ra, w) {
    const b = this.bbox(), rule = b.r1 === b.r2 && b.c1 === b.c2 && !this.extra ? this.sheet.validationFor(this.curRef()) : null;
    if (!rule || rule.type !== 'list' || this.locked || this.editing) { this.ddBtn.hidden = true; if (this.ddList) this.closeList(); return; }
    this.ddBtn.hidden = false; Object.assign(this.ddBtn.style, { left: ra.right - w.left - 22 + 'px', top: ra.top - w.top + 4 + 'px' });
  }
  openList() {
    const rule = this.sheet.validationFor(this.curRef()); if (!rule || rule.type !== 'list' || this.locked) return;
    this.closeList(); const items = this.sheet.listItems(rule); if (!items.length) return;
    const r = this.td[this.cur.r][this.cur.c].getBoundingClientRect(), m = document.createElement('div'); m.className = 'xl-ddl'; m.setAttribute('role', 'listbox');
    m.innerHTML = items.map((x) => '<button type="button" role="option">' + esc(x) + '</button>').join('');
    Object.assign(m.style, { left: r.left + 'px', top: r.bottom + 1 + 'px', minWidth: r.width + 'px' }); document.body.appendChild(m); this.ddList = m;
    m.onmousedown = (e) => e.preventDefault();
    m.onclick = (e) => { const bt = e.target.closest('button'); if (!bt) return; const val = items[[...m.children].indexOf(bt)]; this.closeList(); this.pushUndo(); this.sheet.commit(this.curRef(), val); this.afterChange('commit'); this.focus(); };
    this.ddOff = (e) => { if (!m.contains(e.target) && e.target !== this.ddBtn) this.closeList(); }; setTimeout(() => document.addEventListener('mousedown', this.ddOff), 0);
  }
  closeList() { if (this.ddList) { this.ddList.remove(); this.ddList = null; } if (this.ddOff) { document.removeEventListener('mousedown', this.ddOff); this.ddOff = null; } }
  rejectEntry(text) {
    this.bar.value = this.sheet.raw(this.curRef());
    this.dialog({ title: 'Stop', okLabel: 'Retry', body: "<p>This value doesn't match the data validation restrictions defined for this cell.</p>", ok: () => { setTimeout(() => this.beginEdit(text, 'edit'), 0); return true; } });
  }
  openDataValidation() {
    if (this.locked) return this.o.onBlocked && this.o.onBlocked('locked');
    const range = this.rangeStr(), cur = this.sheet.validationFor(this.curRef()) || { type: 'any' }, opt = (v, t) => '<option value="' + v + '"' + (cur.type === v ? ' selected' : '') + '>' + t + '</option>';
    const d = this.dialog({ title: 'Data Validation', okLabel: 'OK',
      body: '<div class="rb-drow"><span class="xl-lbl">Allow:</span><select data-type>' + opt('any', 'Any value') + opt('whole', 'Whole number') + opt('decimal', 'Decimal') + opt('list', 'List') + '</select></div>'
        + '<div data-box="list" class="rb-drow"><span class="xl-lbl">Source:</span><input type="text" data-src size="26" value="' + esc(cur.source || '') + '" aria-label="Source"></div>'
        + '<div data-box="num" class="rb-drow"><span class="xl-lbl">Minimum:</span><input type="text" data-min size="7" value="' + (cur.min !== undefined ? cur.min : '') + '" aria-label="Minimum"> <span>Maximum:</span> <input type="text" data-max size="7" value="' + (cur.max !== undefined ? cur.max : '') + '" aria-label="Maximum"></div>'
        + '<p class="xl-dlg-note">Applies to ' + range + '</p>',
      ok: (dd) => {
        const type = dd.querySelector('[data-type]').value; this.pushUndo();
        if (type === 'any') { this.sheet.clearValidation(range); this.afterChange('validation'); return true; }
        if (type === 'list') { const src = dd.querySelector('[data-src]').value.trim(); if (!src) { this.undo.pop(); return 'Type at least one item or a range.'; } this.sheet.setValidation(range, { type, source: src }); }
        else { const mn = dd.querySelector('[data-min]').value.trim(), mx = dd.querySelector('[data-max]').value.trim(), a = Number(mn), z = Number(mx);
          if (mn === '' || mx === '' || isNaN(a) || isNaN(z)) { this.undo.pop(); return 'Type numbers for the minimum and maximum.'; } if (a > z) { this.undo.pop(); return 'The minimum cannot be larger than the maximum.'; }
          this.sheet.setValidation(range, { type, min: a, max: z }); }
        this.afterChange('validation'); return true;
      } });
    const upd = () => { const t = d.querySelector('[data-type]').value; d.querySelector('[data-box=list]').hidden = t !== 'list'; d.querySelector('[data-box=num]').hidden = t !== 'whole' && t !== 'decimal'; };
    d.querySelector('[data-type]').onchange = upd; upd();
  }
  openDefineName() {
    if (this.locked) return this.o.onBlocked && this.o.onBlocked('locked');
    const abs = (ref) => { const p = parseRef(ref); return '$' + numToCol(p.c) + '$' + p.r; }, rs = this.rangeStr(), refAbs = rs.split(':').map(abs).join(':');
    this.dialog({ title: 'New Name', okLabel: 'OK',
      body: '<div class="rb-drow"><span class="xl-lbl">Name:</span><input type="text" data-nm size="22" aria-label="Name"></div><div class="rb-drow"><span class="xl-lbl">Refers to:</span><input type="text" data-rf size="22" value="=' + refAbs + '" aria-label="Refers to"></div>',
      ok: (dd) => {
        const name = dd.querySelector('[data-nm]').value.trim(); if (!name) return 'Type a name.';
        const ref = dd.querySelector('[data-rf]').value.trim().replace(/^=/, '').replace(/^[^!]*!/, '').replace(/\$/g, '').toUpperCase();
        if (!/^[A-Z]+\d+(:[A-Z]+\d+)?$/.test(ref)) return 'Enter a valid reference, such as B2 or D2:D9.';
        this.pushUndo(); const bad = this.sheet.defineName(name, ref);
        if (bad) { this.undo.pop(); return bad === 'cell' ? 'A name cannot look like a cell address such as A1.' : 'Names must start with a letter or underscore and contain no spaces.'; }
        this.afterChange('name'); return true;
      } });
  }
  openNameManager() {
    const render = (d) => {
      const list = this.sheet.listNames();
      d.querySelector('.xl-names').innerHTML = list.length ? list.map((n) => '<div class="xl-nm"><b>' + esc(n.name) + '</b><span>' + esc(n.ref) + '</span><button type="button" class="btn btn-secondary btn-sm" data-del="' + esc(n.name) + '">Delete</button></div>').join('') : '<p class="xl-dlg-note">No names defined yet.</p>';
      d.querySelectorAll('[data-del]').forEach((bt) => { bt.onclick = () => { this.pushUndo(); this.sheet.deleteName(bt.dataset.del); this.afterChange('name'); render(d); }; });
    };
    render(this.dialog({ title: 'Name Manager', cancelLabel: 'Close', body: '<div class="xl-names"></div>' }));
  }
  openPasteSpecial() {
    if (this.locked) return this.o.onBlocked && this.o.onBlocked('locked');
    if (!this.clip) return this.toast('Copy something first (Ctrl + C), then open Paste Special.');
    const r = (g, v, t, chk) => `<label class="rb-rad"><input type="radio" name="${g}" value="${v}"${chk ? ' checked' : ''}> ${t}</label>`;
    this.dialog({ title: 'Paste Special', okLabel: 'OK',
      body: '<p><b>Paste</b></p>' + r('m', 'all', 'All', true) + r('m', 'formulas', 'Formulas') + r('m', 'values', 'Values') + r('m', 'formats', 'Formats') + '<p><b>Operation</b></p>' + r('o', '', 'None', true) + r('o', 'add', 'Add') + r('o', 'subtract', 'Subtract') + r('o', 'multiply', 'Multiply') + r('o', 'divide', 'Divide') + '<label class="rb-rad"><input type="checkbox" data-tr> Transpose</label>',
      ok: (d) => {
        this.pushUndo(); const mode = d.querySelector('[name=m]:checked').value, op = d.querySelector('[name=o]:checked').value || null, b = this.bbox(), one = this.clip.rows.length === 1 && this.clip.rows[0].length === 1;
        if (one && (b.r1 !== b.r2 || b.c1 !== b.c2) && !d.querySelector('[data-tr]').checked) { for (let rr = b.r1; rr <= b.r2; rr++) for (let cc = b.c1; cc <= b.c2; cc++) this.sheet.pasteBlock(this.clip.block, rr, cc, { mode, op }); }
        else this.sheet.pasteBlock(this.clip.block, this.cur.r, this.cur.c, { mode, op, transpose: d.querySelector('[data-tr]').checked });
        this.afterChange('paste'); return true;
      } });
  }
  switchTo(name) {
    if (!this.sheets || !this.sheets[name]) return;
    this.ac.abort(); this.ac = new AbortController(); this.activeName = name; this.sheet = this.sheets[name];
    this.cur = { r: 1, c: 1 }; this.ext = { r: 1, c: 1 }; this.editing = false; this.build(); this.refresh(); this.paint(); this.focus(); this.emitSelect();
    this.o.onChange && this.o.onChange({ kind: 'sheet' });
  }
}
root.Grid = Grid;
})(window);
