/* Spreadsheet Academy app shell: router, progress store, home dashboard and the lesson player. */
(function () {
'use strict';
const { TRACKS, LESSONS } = window.COURSE;
const { Sheet } = window.XL;
const I18N = window.I18N, tt = I18N.t;
const setTitle = (s) => { document.title = I18N.tx(s); };
const app = document.getElementById('app');
const E = window.Entitlement;
const payHref = () => '#/lesson/' + (LESSONS.find((l) => !E.has(l.id)) || LESSONS[0]).id;
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/* ───────── icons ───────── */
const I = (d, extra = '') => `<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" ${extra}>${d}</svg>`;
const ico = { skip: I('<path d="m5 5 7 7-7 7M13 5l7 7-7 7"/>'),
  bolt: I('<path d="M13 2 4 14h7l-1 8 9-12h-7l1-8Z" fill="currentColor" stroke="none"/>'),
  flame: I('<path d="M12 2c1 4 5 5.5 5 10a5 5 0 0 1-10 0c0-2 1-3 2-4 0 2 1 3 2 3 0-3-1-5 1-9Z" fill="currentColor" stroke="none"/>'),
  check: I('<path d="m5 12.5 4.5 4.5L19 7"/>'),
  chev: I('<path d="m6 9 6 6 6-6"/>'),
  arrow: I('<path d="M5 12h14M13 6l6 6-6 6"/>'),
  back: I('<path d="M19 12H5M11 6l-6 6 6 6"/>'),
  bulb: I('<path d="M9 18h6M10 21h4M12 3a6 6 0 0 0-3.5 10.9c.6.5 1 1.2 1 2.1h5c0-.9.4-1.6 1-2.1A6 6 0 0 0 12 3Z"/>'),
  eye: I('<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z"/><circle cx="12" cy="12" r="3"/>'),
  reset: I('<path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5"/>'),
  clock: I('<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>'),
  x: I('<path d="M18 6 6 18M6 6l12 12"/>'),
  info: I('<circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/>'),
  play: I('<path d="M7 4.5v15l12-7.5-12-7.5Z" fill="currentColor" stroke="none"/>'),
  kb: I('<rect x="2" y="6" width="20" height="12" rx="2.5"/><path d="M6 10h.01M10 10h.01M14 10h.01M18 10h.01M7 14h10"/>'),
  globe: I('<circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3a14 14 0 0 1 0 18M12 3a14 14 0 0 0 0 18"/>'),
  sound: I('<path d="M11 5 6 9H3v6h3l5 4V5Z"/><path d="M15.5 8.5a5 5 0 0 1 0 7M18.5 5.5a9 9 0 0 1 0 13"/>'),
  mute: I('<path d="M11 5 6 9H3v6h3l5 4V5Z"/><path d="m16 9 6 6M22 9l-6 6"/>'),
  lock: I('<rect x="4" y="11" width="16" height="10" rx="2.5"/><path d="M8 11V8a4 4 0 0 1 8 0v3M12 15v2"/>'),
  cog: I('<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1Z"/>'),
  moon: I('<path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8Z"/>'),
  grid: I('<rect x="3" y="3"  width="18" height="18" rx="3"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/>'),
};
const logo = `<img class="logo" src="assets/logo-128.png" width="36" height="36" alt="">`;

/* ───────── progress store ───────── */
const KEY = 'gridwise.progress.v1';
const today = () => new Date().toLocaleDateString('en-CA');
const yesterday = () => { const d = new Date(); d.setDate(d.getDate() - 1); return d.toLocaleDateString('en-CA'); };
const store = {
  d: { xp: 0, done: {}, skipped: {}, completed: {}, streak: { count: 0, last: null } },
  load() { try { const s = JSON.parse(localStorage.getItem(KEY)); if (s) this.d = Object.assign(this.d, s); } catch (e) { /* storage unavailable */ } },
  save() { try { localStorage.setItem(KEY, JSON.stringify(this.d)); } catch (e) { /* ignore */ } },
  reset() { this.d = { xp: 0, done: {}, skipped: {}, completed: {}, streak: { count: 0, last: null } }; this.save(); },
  doneSet(id) { return new Set(this.d.done[id] || []); },
  skippedSet(id) { return new Set(((this.d.skipped || {})[id] || []).filter((i) => !(this.d.done[id] || []).includes(i))); },
  /** the learner chose to move on without solving this step: no XP, and it can be solved later */
  skip(id, idx) { const s = this.skippedSet(id); s.add(idx); (this.d.skipped = this.d.skipped || {})[id] = [...s]; this.touch(); this.save(); },
  mark(id, idx, xp) {
    const s = this.doneSet(id); s.add(idx); this.d.done[id] = [...s]; if (this.d.skipped && this.d.skipped[id]) this.d.skipped[id] = this.d.skipped[id].filter((i) => i !== idx); this.d.xp += xp; this.touch(); this.save();
  },
  complete(id, bonus) { this.d.completed[id] = true; this.d.xp += bonus; this.touch(); this.save(); },
  touch() { const s = this.d.streak, t = today(); if (s.last === t) return; s.count = s.last === yesterday() ? s.count + 1 : 1; s.last = t; },
  streakNow() { const s = this.d.streak; return s.last === today() || s.last === yesterday() ? s.count : 0; },
  streakToday() { return this.d.streak.last === today(); },
};
const THEME_KEY = 'sa.theme.v1';
const theme = {
  get dark() { return document.documentElement.getAttribute('data-theme') === 'dark'; },
  set(d) { document.documentElement.setAttribute('data-theme', d ? 'dark' : 'light'); try { localStorage.setItem(THEME_KEY, d ? 'dark' : 'light'); } catch (e) { /* ignore */ } },
};
I18N.ui([['Settings', 'Configurações', 'Configuración'], ['Dark mode', 'Modo escuro', 'Modo oscuro']]);
const settingsMenu = () => `<span class="set-wrap"><button type="button" class="nav-gear" id="gear" aria-haspopup="true" aria-expanded="false" aria-label="Settings" title="Settings">${ico.cog}</button>
  <div class="set-panel" id="set-panel" role="menu" hidden>
    <div class="set-row"><span>${ico.moon}Dark mode</span><button type="button" class="switch" role="switch" aria-checked="${theme.dark}" data-theme-toggle aria-label="Dark mode"><i></i></button></div>
    ${E.full || E.web ? '' : `<button type="button" class="set-item" id="restore-purchase" role="menuitem">${ico.lock}Restore purchase</button>`}
    <div class="set-sep"></div>
    <button type="button" class="set-item danger" id="reset-progress" role="menuitem">${ico.reset}Reset my progress</button>
  </div></span>`;
const LEVELS = [['Rookie', 0], ['Apprentice', 100], ['Analyst', 250], ['Specialist', 450], ['Strategist', 700], ['Wizard', 1000], ['Grandmaster', 1400]];
function levelInfo(xp) {
  let i = 0; while (i + 1 < LEVELS.length && xp >= LEVELS[i + 1][1]) i++;
  const cur = LEVELS[i][1], next = LEVELS[i + 1] ? LEVELS[i + 1][1] : null;
  return { n: i + 1, name: LEVELS[i][0], pct: next ? Math.round(((xp - cur) / (next - cur)) * 100) : 100, toNext: next ? next - xp : 0 };
}
const lessonProgress = (l) => { const d = store.doneSet(l.id).size; return { done: d, total: l.steps.length, pct: Math.round((d / l.steps.length) * 100) }; };
const nextLesson = () => LESSONS.find((l) => E.has(l.id) && lessonProgress(l).done < l.steps.length) || null;

/* ───────── shared chrome ───────── */
function topbar(active, bare) {
  const li = levelInfo(store.d.xp), st = store.streakNow();
  const link = (h, k, t) => `<a href="${h}" class="${active === k ? 'on' : ''}">${t}</a>`;
  return `<header class="nav"><div class="wrap nav-in">
    <a class="brand" href="${bare ? '#/welcome' : '#/'}" aria-label="Spreadsheet Academy home">${logo}<span>Spreadsheet Academy</span></a>
    <nav class="nav-links" aria-label="Main">${link('#/', 'dash', hasProgress() ? 'My learning' : 'Home')}${link('#/courses', 'courses', 'Courses')}${settingsMenu()}</nav>
    <div class="nav-r">${langSel()}${soundBtn()}${bare ? `<a class="btn btn-light" href="#/courses">Browse courses</a>` : `
      <span class="pill streak ${st ? 'on' : ''}" title="Days in a row you've practised">${ico.flame}<b>${st}</b><span class="hide-s">day streak</span></span>
      <span class="pill xp" title="Experience points">${ico.bolt}<b>${store.d.xp}</b><span class="hide-s">XP</span></span>
      <span class="pill lvl" title="${li.toNext ? li.toNext + ' XP to next level' : 'Max level'}">Lv ${li.n} · ${li.name}</span>`}
    </div></div></header>`;
}
const soundBtn = () => `<button class="icon-btn snd" data-sound-toggle aria-label="${window.Sound.on ? 'Mute sounds' : 'Unmute sounds'}" title="${window.Sound.on ? 'Sound on' : 'Sound off'}">${window.Sound.on ? ico.sound : ico.mute}</button>`;
const langSel = () => `<label class="lang" title="Language / Idioma">${ico.globe}<select class="lang-sel" data-lang aria-label="Language / Idioma">${I18N.LANGS.map(([c, n]) => `<option value="${c}"${c === I18N.lang ? ' selected' : ''}>${n}</option>`).join('')}</select></label>`;
const hasProgress = () => LESSONS.some((l) => lessonProgress(l).done > 0);
const ring = (pct, tone) => `<div class="ring ${tone}" style="--p:${pct}"><span>${pct}%</span></div>`;
const AREAS = [
  ['Cells & formulas', ['grid-basics', 'first-formulas', 'sum-basics', 'true-false']],
  ['Shortcuts & speed', ['keyboard-nav', 'selection-2', 'fill-refs', 'edit-shortcuts']],
  ['Logic & decisions', ['logic-if', 'if-combos', 'advanced-logic']],
  ['Counting & summing', ['conditional-aggs', 'count-more', 'sum-practice', 'multi-criteria']],
  ['Text & dates', ['text-tools', 'text-more', 'text-combine', 'dates']],
  ['Lookups', ['lookups', 'lookups-2', 'lookup-mastery']],
  ['Stats & rounding', ['stats', 'rounding']],
  ['Arrays & dynamic', ['array-thinking', 'dynamic-arrays']],
  ['Tables & reporting', ['worksheets', 'cond-format', 'tables-filters', 'pivot-tables']],
  ['Charts & dashboards', ['charts-first', 'charts-choose', 'charts-polish', 'test-dashboard']],
  ['Real-world tests', ['test-foundations', 'test-retail', 'capstone-sales', 'test-realestate', 'test-hospital']],
];
const remainingMin = (l) => Math.max(1, Math.round((l.minutes * (l.steps.length - lessonProgress(l).done)) / l.steps.length));
const lessonRow = (l) => {
  const p = lessonProgress(l), tr = TRACKS.find((t) => t.id === l.track), lk = !E.has(l.id), state = lk ? 'locked' : p.done === 0 ? 'todo' : p.done === p.total ? 'done' : 'progress';
  return `<div class="row tone-${tr.tone}" data-state="${state}">
    <span class="row-ico">${state === 'done' ? ico.check : lk ? ico.lock : l.index}</span>
    <div class="row-main"><span class="row-k">${tr.level} · ${tr.name}</span><a class="row-t" href="#/lesson/${l.id}">${l.title}</a>${l.kind === 'test' ? '<span class="badge-test">Real-world test</span>' : ''}${E.isFree(l.id) && !E.full ? '<span class="badge-free">Free</span>' : ''}
      <div class="row-bar"><div class="bar bar-sm"><i style="width:${p.pct}%"></i></div><span>${p.pct}%</span></div></div>
    <p class="row-s">${l.summary}</p>
    <a class="btn ${state === 'done' || lk ? 'btn-secondary' : 'btn-primary'} btn-sm" href="#/lesson/${l.id}">${lk ? ico.lock + 'Unlock' : state === 'todo' ? 'Start' : state === 'done' ? 'Review' : 'Continue'}</a></div>`;
};
const footer = () => '';
/** In-app message box (replaces the browser's alert/confirm): ask({title, text, ok, cancel:true, danger}) -> Promise<boolean> */
function ask({ title, text, ok = 'OK', cancel = false, danger = false }) {
  return new Promise((res) => {
    const ov = document.createElement('div'); ov.className = 'overlay dlg'; ov.setAttribute('role', 'alertdialog'); ov.setAttribute('aria-modal', 'true');
    ov.innerHTML = `<div class="modal dlg-box">${title ? `<h2>${esc(title)}</h2>` : ''}<p class="muted">${esc(text)}</p>
      <div class="dlg-actions">${cancel ? `<button class="btn btn-secondary" data-r="0">${esc(tt('Cancel'))}</button>` : ''}<button class="btn ${danger ? 'btn-danger' : 'btn-primary'}" data-r="1">${esc(ok)}</button></div></div>`;
    const done = (v) => { document.removeEventListener('keydown', key, true); ov.remove(); res(v); };
    const key = (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); done(false); } else if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); done(true); } };
    ov.onclick = (e) => { const b = e.target.closest('[data-r]'); if (b) done(b.dataset.r === '1'); else if (e.target === ov) done(false); };
    document.addEventListener('keydown', key, true); document.body.appendChild(ov); ov.querySelector('[data-r="1"]').focus();
  });
}
async function restorePurchase() {
  const r = await E.restore();
  await ask({ text: tt(r === 'unavailable' ? 'Purchases are available in the Microsoft Store version of the app.' : r === 'none' ? 'No purchase found for this account.' : 'Full course unlocked. Enjoy!') });
}
function wireCommon() {
  const f = document.getElementById('find-level'); if (f) f.onclick = levelModal;
  document.querySelectorAll('[data-find-level]').forEach((b) => { b.onclick = levelModal; });
  const rp = document.getElementById('restore-purchase');
  if (rp) rp.onclick = restorePurchase;
  const r = document.getElementById('reset-progress');
  if (r) r.onclick = async () => { if (await ask({ title: tt('Reset my progress'), text: tt('Erase all progress, XP and streak on this device?'), ok: tt('Reset my progress'), cancel: true, danger: true })) { store.reset(); location.hash = '#/welcome'; route(); } };
}
const mockSheet = `<div class="mock">
  <div class="mock-bar"><span class="mock-name">G2</span><i>fx</i><span class="mock-f"><span class="typing">=SUMIFS(D2:D9,B2:B9,"East")</span></span></div>
  <div class="mock-grid">
    ${['', 'A', 'B', 'C', 'D'].map((x) => `<b class="mh">${x}</b>`).join('')}
    ${[['1', 'Rep', 'Region', 'Product', 'Sales'], ['2', 'Ana', 'West', 'Laptop', '1,200'], ['3', 'Ben', 'East', 'Phone', '800'], ['4', 'Cy', 'West', 'Phone', '650'], ['5', 'Eli', 'East', 'Laptop', '1,100']].map((r, i) => r.map((v, j) => `<span class="${j === 0 ? 'mh' : ''} ${i === 0 && j ? 'th' : ''} ${(i === 2 || i === 4) && j === 4 ? 'hit' : ''} ${j === 4 && i ? 'num' : ''}">${v}</span>`).join('')).join('')}
  </div>
  <div class="mock-chart" aria-hidden="true"><b class="mc-t">Sales by rep</b><svg viewBox="0 0 160 92"><path class="mc-grid" d="M6 16H154M6 36H154M6 56H154"/><path class="mc-base" d="M6 76H154"/><rect class="mc-bar" style="--k:0" x="12" y="16" width="26" height="60" rx="2.5"/><text class="mc-x" x="25" y="88">Ana</text><rect class="mc-bar east" style="--k:1" x="48" y="36" width="26" height="40" rx="2.5"/><text class="mc-x" x="61" y="88">Ben</text><rect class="mc-bar" style="--k:2" x="84" y="43.5" width="26" height="32.5" rx="2.5"/><text class="mc-x" x="97" y="88">Cy</text><rect class="mc-bar east" style="--k:3" x="120" y="21" width="26" height="55" rx="2.5"/><text class="mc-x" x="133" y="88">Eli</text><path class="mc-line" d="M25 16 L61 36 L97 43.5 L133 21"/></svg><span class="mc-sum">East = 1,900</span></div>
  <div class="mock-toast">${ico.check}<div><b>Correct!</b><span>+10 XP</span></div></div></div>`;

/* ───────── LANDING (marketing only) ───────── */
function landingView() {
  setTitle('Spreadsheet Academy · Learn Excel by doing');
  const first = LESSONS[0];
  app.innerHTML = `${topbar('dash', true)}<main>
  <section class="hero wrap">
    <div class="hero-copy">
      <span class="eyebrow">Hands-on Excel training</span>
      <h1>Learn Excel by <em>actually using it.</em></h1>
      <p class="lead">Short lessons with a live spreadsheet in every step. Type formulas, press real shortcuts and get instant feedback, from your very first cell to lookups and array formulas.</p>
      <div class="hero-cta"><a class="btn btn-primary btn-lg" href="#/lesson/${first.id}">Start learning free ${ico.arrow}</a><button class="btn btn-ghost btn-lg" id="find-level">Find my level</button></div>
      <ul class="points"><li>${ico.check}No install, no signup</li><li>${ico.kb}Keyboard-first practice</li><li>${ico.bolt}Checked as you type</li></ul>
    </div>
    <div class="hero-art" aria-hidden="true">${mockSheet}</div>
  </section>

  <section class="band"><div class="wrap numbers">
    <div><b>${LESSONS.length}</b><span>hands-on lessons</span></div>
    <div><b>${LESSONS.reduce((a, l) => a + l.steps.length, 0)}</b><span>checked exercises</span></div>
    <div><b>110+</b><span>Excel functions in the live sheet</span></div>
    <div><b>${TRACKS.length}</b><span>learning paths</span></div></div></section>

  <section class="wrap sect">
    <div class="sec-head center"><span class="eyebrow">How it works</span><h2>Learn the way you'll actually work</h2></div>
    <div class="how">
      <div class="how-c"><span class="how-n">1</span><h3>Read a short idea</h3><p>One concept per step, explained in a few sentences. No videos to sit through.</p></div>
      <div class="how-c"><span class="how-n">2</span><h3>Do it in a live sheet</h3><p>Type real formulas and press real shortcuts in a spreadsheet that calculates as you go.</p></div>
      <div class="how-c"><span class="how-n">3</span><h3>Get instant feedback</h3><p>Know straight away what's right, what's off and why. Earn XP and build a streak.</p></div>
    </div></section>

  <section class="wrap sect">
    <div class="sec-head center"><span class="eyebrow">Learning paths</span><h2>Start where you are</h2><p>${E.full ? 'Pick a path. Every lesson is open from day one.' : 'Pick a path. The first two lessons are free.'}</p></div>
    <div class="paths">${TRACKS.map((t) => `<article class="path tone-${t.tone}">
      <span class="lvl-badge">${t.level}</span><h3>${t.name}</h3><p>${t.tagline}</p>
      <div class="chips">${[...new Set(t.lessons.flatMap((l) => l.skills))].slice(0, 6).map((s) => `<span>${esc(s)}</span>`).join('')}</div>
      <div class="path-meta"><span>${t.lessons.length} lessons</span><span>~${t.lessons.reduce((a, l) => a + l.minutes, 0)} min</span></div>
      <a class="btn btn-secondary" href="#/lesson/${t.lessons[0].id}">Start path ${ico.arrow}</a></article>`).join('')}</div></section>

  <section class="wrap sect"><div class="cta-band"><div><h2>Ready to stop guessing at formulas?</h2><p>Your first lesson takes five minutes.</p></div>
    <a class="btn btn-light btn-lg" href="#/lesson/${first.id}">Start the first lesson ${ico.arrow}</a></div></section>
  ${footer()}</main>`;
  wireCommon();
}

/* ───────── DASHBOARD (personal progress only) ───────── */
function dashboardView() {
  setTitle('My learning · Spreadsheet Academy');
  const next = nextLesson(), li = levelInfo(store.d.xp);
  const total = LESSONS.reduce((a, l) => a + l.steps.length, 0), solved = LESSONS.reduce((a, l) => a + lessonProgress(l).done, 0);
  const lessonsDone = LESSONS.filter((l) => lessonProgress(l).done === l.steps.length).length;
  const tr = next && TRACKS.find((t) => t.id === next.track), np = next && lessonProgress(next);
  const active = LESSONS.filter((l) => { const p = lessonProgress(l); return p.done > 0 && p.done < p.total && l !== next; });
  const upNext = LESSONS.filter((l) => lessonProgress(l).done === 0 && l !== next).slice(0, Math.max(2, 4 - active.length));
  const area = AREAS.map(([name, ids]) => [name, ids.reduce((a, id) => a + store.doneSet(id).size, 0) * 10]);
  const max = Math.max(100, ...area.map((a) => a[1]));
  const statsBlock = `<div class="hc-stats"><div class="hs"><span>Steps solved</span><b>${solved}<small>/${total}</small></b></div><div class="hs"><span>Lessons done</span><b>${lessonsDone}<small>/${LESSONS.length}</small></b></div><div class="hs"><span>Streak</span><b>${store.streakNow()}<small> d</small></b></div><div class="hs-bar"><div class="bar bar-sm"><i style="width:${li.pct}%"></i></div><small>${li.toNext ? li.toNext + ' XP to next level' : 'Top level'}</small></div></div>`;
  app.innerHTML = `${topbar('dash')}<main class="wrap dash">
    <div class="dash-head">
      <div class="avatar" aria-hidden="true">${ico.grid}</div>
      <div><h1>Welcome back!</h1><p class="muted">${store.d.xp.toLocaleString({ en: 'en-US', pt: 'pt-BR', es: 'es-ES' }[I18N.lang])} XP · Level ${li.n} ${li.name}${li.toNext ? ` · ${li.toNext} XP to level ${li.n + 1}` : ''}</p></div>
    </div>
    ${next ? `<section class="hero-card tone-${tr.tone}">
      <span class="hc-ico">${ico.bulb}</span>
      <div class="hc-main"><span class="eyebrow-s">Continue learning</span><h2><a href="#/lesson/${next.id}">${next.title}</a></h2>
        <p class="muted">${tr.level} · ${tr.name} · Lesson ${next.index}</p>
        <div class="hc-prog"><div class="bar"><i style="width:${np.pct}%"></i></div><b>${np.pct}%</b><span class="muted">${ico.clock}${remainingMin(next)} min to go</span></div>
        <div class="hc-cta"><a class="btn btn-primary btn-lg" href="#/lesson/${next.id}">${np.done ? 'Keep making progress' : 'Start lesson'}</a></div></div>
      ${statsBlock}
    </section>` : !E.full ? `<section class="hero-card"><span class="hc-ico">${ico.lock}</span><div class="hc-main"><span class="eyebrow-s">Free lessons complete</span><h2>Unlock the full course</h2><p class="muted">You have finished the free lessons. Unlock the full course to keep going.</p><div class="hc-cta"><a class="btn btn-primary btn-lg" href="${payHref()}">Unlock full course</a></div></div>${statsBlock}</section>` : `<section class="hero-card"><span class="hc-ico">${ico.check}</span><div class="hc-main"><span class="eyebrow-s">Course complete</span><h2>You finished every lesson!</h2><p class="muted">Review any lesson to keep your skills sharp.</p><div class="hc-cta"><a class="btn btn-primary btn-lg" href="#/courses">Browse courses</a></div></div>${statsBlock}</section>`}
    <section class="tiles">
      <button class="tile" data-find-level><span class="tile-i t1">${ico.eye}</span><div><span class="tile-k">Assess</span><b>Find my level</b></div>${ico.arrow}</button>
      <a class="tile" href="#/lesson/keyboard-nav"><span class="tile-i t2">${ico.kb}</span><div><span class="tile-k">Practice</span><b>Keyboard shortcut drills</b></div>${ico.arrow}</a>
      <a class="tile" href="#/lesson/capstone-sales"><span class="tile-i t3">${ico.grid}</span><div><span class="tile-k">Apply</span><b>Capstone: sales report</b></div>${ico.arrow}</a>
    </section>
    <div class="dash-cols">
      <div class="dash-main">
        ${active.length ? `<section class="panel"><div class="panel-h"><h2>In progress</h2></div>${active.map(lessonRow).join('')}</section>` : ''}
        <section class="panel"><div class="panel-h"><h2>Up next</h2><a class="more" href="#/courses">All courses ${ico.arrow}</a></div>${upNext.length ? upNext.map(lessonRow).join('') : '<p class="muted pad">Nothing left. Great work!</p>'}</section>
      </div>
      <aside class="dash-side">
        <section class="panel">
          <div class="panel-h"><h2>Skills and XP</h2></div>
          <div class="skills">${area.map(([n, v]) => `<div class="sk"><span>${n}</span><div class="sk-bar"><i style="width:${Math.round((v / max) * 100)}%"></i></div><b>${v} XP</b></div>`).join('')}</div>
        </section>
      </aside>
    </div></main>${footer()}`;
  wireCommon();
}

/* ───────── CATALOG ───────── */
/* which path banners are expanded (remembered per device) */
const OPEN_KEY = 'sa.catalog.open.v1';
const loadOpen = () => { try { return JSON.parse(localStorage.getItem(OPEN_KEY)) || null; } catch (e) { return null; } };
const saveOpen = (o) => { try { localStorage.setItem(OPEN_KEY, JSON.stringify(o)); } catch (e) { /* ignore */ } };

function catalogView(filter) {
  setTitle('Courses · Spreadsheet Academy');
  const f = filter || 'all', nxt = nextLesson();
  let open = loadOpen(); if (!open) { open = {}; TRACKS.forEach((t) => { open[t.id] = !!nxt && nxt.track === t.id || (!nxt && t === TRACKS[0]); }); }
  const shown = TRACKS.filter((t) => f === 'all' || f === t.id), single = f !== 'all';
  const isOpen = (t) => single || !!open[t.id];
  app.innerHTML = `${topbar('courses')}<main class="wrap catalog">
    <div class="cat-head"><div><h1>Courses</h1><p class="muted">${LESSONS.length} lessons in ${TRACKS.length} learning paths. ${E.full ? 'Everything is unlocked, so start anywhere.' : 'The first two lessons are free.'}</p></div>
      <div class="cat-actions">${single ? '' : '<button class="btn btn-ghost" id="toggle-all"></button>'}<button class="btn btn-secondary" data-find-level>${ico.eye}Find my level</button></div></div>
    <div class="tabs" role="tablist">${[['all', 'All paths'], ...TRACKS.map((t) => [t.id, t.name])].map(([k, n]) => `<button class="tab ${f === k ? 'on' : ''}" data-f="${k}" role="tab">${n}</button>`).join('')}</div>
    ${shown.map((t) => {
      const steps = t.lessons.reduce((a, l) => a + l.steps.length, 0), done = t.lessons.reduce((a, l) => a + lessonProgress(l).done, 0), pct = Math.round((done / steps) * 100);
      const o = isOpen(t), tests = t.lessons.filter((l) => l.kind === 'test').length;
      return `<section class="track banner tone-${t.tone}" data-track="${t.id}" data-open="${o}">
        <div class="track-head" role="button" tabindex="0" aria-expanded="${o}" aria-controls="rows-${t.id}">
          <div class="tb-main"><span class="lvl-badge">${t.level}</span><h2>${t.name}</h2><p>${t.tagline}</p>
            <div class="track-meta"><span>${ico.grid}${t.lessons.length} lessons</span><span>${ico.clock}~${t.lessons.reduce((a, l) => a + l.minutes, 0)} min</span><span>${ico.check}${done}/${steps} steps solved</span>${tests ? `<span>${ico.bolt}${tests} ${tests === 1 ? 'real-world test' : 'real-world tests'}</span>` : ''}</div></div>
          <div class="tb-side">${ring(pct, t.tone)}</div>${single ? '' : `<span class="chev" aria-hidden="true">${ico.chev}</span>`}
        </div>
        <div class="panel rows" id="rows-${t.id}"${o ? '' : ' hidden'}>${t.lessons.map(lessonRow).join('')}</div></section>`;
    }).join('')}</main>${footer()}`;
  app.querySelectorAll('.tab').forEach((b) => { b.onclick = () => catalogView(b.dataset.f); });
  const setAll = () => { const all = shown.every((t) => open[t.id]); const btn = document.getElementById('toggle-all'); if (btn) btn.textContent = all ? 'Collapse all' : 'Expand all'; };
  const toggle = (sec) => {
    if (single) return; const id = sec.dataset.track, now = sec.dataset.open !== 'true';
    sec.dataset.open = now; sec.querySelector('.track-head').setAttribute('aria-expanded', now); sec.querySelector('.rows').hidden = !now; open[id] = now; saveOpen(open); setAll();
  };
  app.querySelectorAll('.track-head').forEach((h) => {
    h.onclick = () => toggle(h.closest('.track'));
    h.onkeydown = (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(h.closest('.track')); } };
  });
  const ta = document.getElementById('toggle-all');
  if (ta) { setAll(); ta.onclick = () => { const all = shown.every((t) => open[t.id]); shown.forEach((t) => { open[t.id] = !all; const sec = app.querySelector('[data-track="' + t.id + '"]'); sec.dataset.open = !all; sec.querySelector('.track-head').setAttribute('aria-expanded', !all); sec.querySelector('.rows').hidden = all; }); saveOpen(open); setAll(); }; }
  wireCommon();
}

function levelModal() {
  const ov = document.createElement('div'); ov.className = 'overlay';
  ov.innerHTML = `<div class="modal" role="dialog" aria-modal="true" aria-label="Find my level">
    <button class="icon-btn modal-x" aria-label="Close">${ico.x}</button>
    <h2>Where should you start?</h2><p class="muted">Pick the sentence that sounds most like you.</p>
    <a class="choice" href="#/lesson/${LESSONS[0].id}"><b>I'm new to Excel</b><span>Start with cells, formulas and keyboard basics.</span>${ico.arrow}</a>
    <a class="choice" href="#/lesson/${TRACKS[1].lessons[0].id}"><b>I can use SUM and simple formulas</b><span>Jump to IF, conditional totals, text and lookups.</span>${ico.arrow}</a>
    <a class="choice" href="#/lesson/${TRACKS[2].lessons[0].id}"><b>I use IF and VLOOKUP, but want mastery</b><span>Go straight to nested logic, SUMIFS and arrays.</span>${ico.arrow}</a></div>`;
  document.body.appendChild(ov);
  const close = () => ov.remove();
  ov.addEventListener('click', (e) => { if (e.target === ov || e.target.closest('.modal-x') || e.target.closest('.choice')) close(); });
  document.addEventListener('keydown', function esc1(e) { if (e.key === 'Escape') { close(); document.removeEventListener('keydown', esc1); } });
  ov.querySelector('.choice').focus();
}

/* ───────── LESSON PLAYER ───────── */
const PRAISE = ['Correct!', 'Nailed it!', 'Exactly right.', 'Clean work.', "That's it!", 'Spot on.'];
const KEYNAMES = { Control: 'Ctrl', Shift: 'Shift', Alt: 'Alt', Meta: 'Ctrl', ' ': 'Space', PageDown: 'PgDn', PageUp: 'PgUp', Delete: 'Del', ArrowUp: '↑', ArrowDown: '↓', ArrowLeft: '←', ArrowRight: '→', End: 'End', Home: 'Home', '=': '=' };
const keyLabel = (e) => {
  if (e.code && /^Digit\d$/.test(e.code)) return e.code.slice(5);
  if (e.code && /^Key[A-Z]$/.test(e.code)) return e.code.slice(3);
  if (e.code === 'Equal') return '=';
  return KEYNAMES[e.key] || (e.key && e.key.length === 1 ? e.key.toUpperCase() : e.key);
};
let teardownLesson = null;

function lessonView(id, startAt) {
  const L = LESSONS.find((l) => l.id === id);
  if (!L) { location.hash = '#/'; return; }
  const track = TRACKS.find((t) => t.id === L.track), n = L.steps.length;
  setTitle(`${L.title} · Spreadsheet Academy`);
  let done = store.doneSet(id), skipped = store.skippedSet(id), idx = 0, step, sheet, grid, passed = false, assisted = false, fails = 0, gained = 0, flashT = null, book = null, toolApi = null;
  const cleared = (i) => done.has(i) || skipped.has(i);   /* solved, or skipped on purpose */
  const firstUndone = () => { for (let i = 0; i < n; i++) if (!cleared(i)) return i; return 0; };

  app.innerHTML = `<div class="lesson tone-${track.tone}">
    <header class="ltop">
      <a class="icon-btn" href="#/" aria-label="Back to all lessons" title="All lessons">${ico.back}</a>
      <div class="ltitle"><span class="crumb">${track.name} · Lesson ${L.index}</span><h1>${L.title}</h1></div>
      <ol class="segs" aria-label="Steps">${L.steps.map((s, i) => `<li><button class="seg" data-i="${i}" aria-label="Step ${i + 1}: ${esc(s.title)}"></button></li>`).join('')}</ol>
      ${langSel()}${soundBtn()}<span class="pill xp">${ico.bolt}<b id="xp-n">${store.d.xp}</b><span class="hide-s">XP</span></span>
    </header>
    <main class="lmain">
      <section class="card ins" id="ins" aria-live="polite"></section>
      <section class="card wk" id="wk"></section>
    </main>
    <footer class="foot" id="foot" data-state="idle">
      <div class="foot-in">
        <div class="foot-l">
          <button class="btn btn-quiet" id="b-hint">${ico.bulb}<span>Hint</span></button>
          <button class="btn btn-quiet" id="b-sol">${ico.eye}<span>Show answer</span></button>
          <button class="btn btn-quiet" id="b-reset" title="Restart this step">${ico.reset}<span class="hide-s">Reset</span></button>
          <button class="btn btn-quiet" id="b-skip" title="Skip this step">${ico.skip}<span class="hide-s">Skip step</span></button>
        </div>
        <div class="foot-msg"><span class="foot-ico"></span><span class="foot-txt" id="foot-txt"></span></div>
        <div class="foot-r">
          <button class="btn btn-secondary" id="b-check">Check my work</button>
          <button class="btn btn-primary" id="b-next" hidden>Continue ${ico.arrow}</button>
        </div>
      </div></footer></div>`;

  const $ = (s) => app.querySelector(s), ins = $('#ins'), wk = $('#wk'), foot = $('#foot'), txt = $('#foot-txt');
  const bHint = $('#b-hint'), bSol = $('#b-sol'), bCheck = $('#b-check'), bNext = $('#b-next'), bSkip = $('#b-skip');

  function paintSegs() {
    app.querySelectorAll('.seg').forEach((b, i) => {
      b.dataset.state = i === idx ? 'current' : done.has(i) ? 'done' : skipped.has(i) ? 'skipped' : 'todo';
      b.disabled = !(cleared(i) || i === idx || i === firstUndone());
    });
    const xp = $('#xp-n'); if (xp) xp.textContent = store.d.xp;
  }
  function setFoot(state, html, icon) {
    foot.dataset.state = state; txt.innerHTML = html;
    $('.foot-ico').innerHTML = icon || (state === 'ok' ? ico.check : state === 'bad' ? ico.x : ico.info);
  }
  function idleMsg() { setFoot('idle', step.quiz ? 'Choose an answer.' : step.check && step.locked ? 'Use the keyboard to complete the task.' : 'Complete the task, then press Check or Enter.'); }
  function flash(html) {
    if (passed) return; Sound.play('blocked'); setFoot('info', html); clearTimeout(flashT); flashT = setTimeout(() => { if (foot.dataset.state === 'info') idleMsg(); }, 3200);
  }

  function keycaps() {
    if (!step.keys) return '';
    const cap = (k) => `<kbd class="cap" data-k="${esc(k)}">${esc(k)}</kbd>`;
    if (Array.isArray(step.keys[0])) return `<div class="keys" aria-label="Shortcut">${step.keys.map((g) => `<span class="combo">${g.map(cap).join('<i class="plus">+</i>')}</span>`).join(`<i class="then">${tt('then')}</i>`)}</div>`;
    const sep = step.keysOr ? '<i class="or">or</i>' : '<i class="plus">+</i>';
    return `<div class="keys" aria-label="Shortcut">${step.keys.map(cap).join(sep)}</div>`;
  }

  function renderStep(i) {
    clearTimeout(flashT);
    if (grid) grid.destroy();
    idx = i; step = L.steps[i]; passed = false; assisted = false; fails = 0;
    if (step.grid.sheets) {
      book = { sheets: {} };
      Object.keys(step.grid.sheets).forEach((nm) => { const sh = new Sheet(step.grid.rows, step.grid.cols).load(step.grid.sheets[nm]); sh.book = book; sh.name = nm; book.sheets[nm] = sh; });
      sheet = book.sheets[step.grid.active || Object.keys(book.sheets)[0]];
    } else { book = null; sheet = new Sheet(step.grid.rows, step.grid.cols).load(step.grid.data); sheet.book = { sheets: { Sheet1: sheet } }; }
    if (step.grid.charts) sheet.charts = JSON.parse(JSON.stringify(step.grid.charts));
    sheet.setup(step.grid);
    toolApi = null;
    ins.innerHTML = `<div class="kicker"><span>${step.challenge ? '<b class="chal">Challenge</b> · ' : ''}Step ${i + 1} of ${n}</span>${done.has(i) ? `<span class="tick">${ico.check}Solved</span>` : skipped.has(i) ? `<span class="tick skip">${ico.skip}Skipped</span>` : ''}</div>
      <h2 data-notranslate>${step.title}</h2><div class="prose" data-notranslate>${step.body}</div>
      <div class="task"><span class="task-tag">Your task</span><p data-notranslate>${step.task}</p>${step.fill ? `<p class="fill-tip">${step.fill.range ? tt('To fill down: select <b>{r}</b> and press <b>Ctrl + D</b>. With only one cell selected, Ctrl + D copies the cell above it.', { r: step.fill.range }) : tt('To fill down: select the cell with the formula and the cells below it, then press <b>Ctrl + D</b>. With only one cell selected, Ctrl + D copies the cell above it.')}</p>` : ''}</div>
      <div class="callout hint" id="c-hint" hidden></div><div class="callout sol" id="c-sol" hidden></div><div class="callout why" id="c-why" hidden></div>`;
    wk.innerHTML = `${keycaps()}${step.tool ? '<div id="tool-host"></div>' : ''}<div class="xl-host" id="xl-host"></div>${step.quiz ? `<div class="quiz" role="group" data-notranslate>${step.quiz.options.map((o, k) => `<button class="opt" data-k="${k}"><span class="opt-l">${'ABCD'[k]}</span><span>${o}</span></button>`).join('')}</div>` : ''}`;
    grid = new window.Grid($('#xl-host'), {
      sheet, sheets: book ? book.sheets : null, activeName: book ? sheet.name : null, cur: step.grid.cur, goto: !!step.grid.goto, locked: !!step.locked || !!step.quiz, mouse: step.mouse !== false,
      onSelect: () => { if (toolApi && toolApi.onSelect) toolApi.onSelect(); evaluate('select'); }, onChange: () => { if (toolApi && toolApi.onChange) toolApi.onChange(); evaluate('commit'); },
      onBlocked: (why, msg) => flash(why === 'info' ? msg : why === 'locked' ? 'Editing is switched off for this exercise. Just navigate.' : 'Keyboard only here. Use the keys shown above the sheet.'),
    });
    wk.classList.toggle('has-side', !!step.tool && step.tool.kind === 'chart');
    if (step.tool) toolApi = window.Tools.mount(step.tool, $('#tool-host'), { grid });
    bNext.hidden = true; bCheck.hidden = !!step.quiz; bSol.hidden = !!step.quiz; bHint.hidden = !!step.quiz && false;
    bHint.disabled = false; bSol.disabled = false; bSkip.hidden = false;
    bNext.innerHTML = `${i === n - 1 ? 'Finish lesson' : 'Continue'} ${ico.arrow}`;
    idleMsg(); paintSegs();
    if (!step.quiz) grid.focus();
    ins.scrollTop = 0;
  }

  const norm = (r) => (r === true ? { ok: true } : r === false || !r ? { ok: false, quiet: true } : r);
  function evaluate(trigger) {
    if (passed || !step.check) return;
    const r = norm(step.check({ sheet: grid.sheet, grid, book }));
    if (r.ok) return pass();
    if (trigger === 'manual' || (trigger === 'commit' && r.msg && !r.quiet)) fail(r.msg || 'Nothing to check yet. Finish the task, then try again.');
  }
  function fail(msg) {
    fails++; setFoot('bad', I18N.msg(msg)); Sound.play('wrong');
    if (fails >= 2) showHint(false);
  }
  function pass(extra) {
    if (passed) return; passed = true; grid.setLocked(true);
    const first = !done.has(idx), xp = first && !assisted ? (step.quiz ? 5 : step.challenge ? 20 : 10) : 0;
    if (first) { store.mark(id, idx, xp); done.add(idx); skipped.delete(idx); gained += xp; }
    const praise = assisted ? 'Answer filled in. No XP this time, but you can keep going.' : PRAISE[Math.floor(Math.random() * PRAISE.length)];
    Sound.play('correct');
    setFoot('ok', `<b>${praise}</b>${xp ? `<span class="xp-pop">+${xp} XP</span>` : ''}`);
    bCheck.hidden = true; bSkip.hidden = true; bNext.hidden = false; bNext.focus();
    app.querySelectorAll('.opt').forEach((o) => { o.disabled = true; });
    if (extra) { const w = $('#c-why'); w.innerHTML = extra; w.hidden = false; }
    paintSegs();
  }
  function showHint(focus) {
    Sound.play('hint');
    const c = $('#c-hint'); c.innerHTML = `<span class="callout-t">${ico.bulb}Hint</span><span data-notranslate>${step.hint || tt(step.keys ? 'Look at the keys above the sheet and try them in order.' : 'Re-read the task, then try it step by step.')}</span>`;
    c.hidden = false; if (focus && grid && !step.quiz) grid.focus();
  }
  function solutionHtml() {
    const s = step.sol || {}, parts = [];
    if (s.text) parts.push(`<p>${s.text}</p>`);
    if (s.sel) parts.push(`<p>${tt('Target selection:')} <b>${s.sel}</b>${step.keys ? ` ${tt('using')} ${step.keys.map((k) => (Array.isArray(k) ? k.map((x) => `<kbd>${esc(x)}</kbd>`).join('+') : `<kbd>${esc(k)}</kbd>`)).join(Array.isArray(step.keys[0]) ? ` ${tt('then')} ` : '+')}` : ''}</p>`);
    if (s.cells) {
      const k = Object.keys(s.cells), shown = k.slice(0, 3);
      parts.push(`<ul class="sol-list">${shown.map((r) => `<li><b>${r}</b><code>${esc(I18N.formula ? I18N.formula(String(s.cells[r])) : s.cells[r])}</code></li>`).join('')}${k.length > 3 ? `<li class="more">${tt(k.length - 3 > 1 ? '…then fill down for the remaining {n} cells' : '…then fill down for the remaining {n} cell', { n: k.length - 3 })}</li>` : ''}</ul>`);
    }
    if (s.fmt) parts.push(`<p>${tt('Select <b>{a}</b> and apply the <b>{f}</b> format.', { a: s.fmt.refs[0] + (s.fmt.refs.length > 1 ? ':' + s.fmt.refs[s.fmt.refs.length - 1] : ''), f: tt(s.fmt.fmt) })}</p>`);
    return parts.join('');
  }
  function applySolution() {
    const s = step.sol || {}; assisted = true;
    if (s.cells) Object.keys(s.cells).forEach((r) => sheet.commit(r, String(s.cells[r])));
    if (s.fmt) s.fmt.refs.forEach((r) => sheet.setFmt(r, s.fmt.fmt));
    if (s.bold) s.bold.forEach((r) => sheet.setBold(r, true));
    if (s.run) s.run({ sheet: grid.sheet, book, grid });
    if (s.sel) grid.select(s.sel);
    grid.refresh(); grid.paint(); evaluate('manual');
  }

  /* quiz */
  wk.addEventListener('click', (e) => {
    const o = e.target.closest('.opt'); if (!o || passed) return;
    const k = +o.dataset.k;
    if (k === step.quiz.answer) { o.classList.add('right'); pass(`<span class="callout-t">${ico.check}Why</span><span data-notranslate>${step.quiz.explain}</span>`); }
    else { o.classList.add('wrong'); o.disabled = true; fail('Not quite. Think it through and try another answer.'); }
  });

  /* controls */
  bHint.onclick = () => showHint(true);
  bCheck.onclick = () => evaluate('manual');
  bNext.onclick = () => { if (idx < n - 1) renderStep(idx + 1); else finish(); };
  $('#b-reset').onclick = () => renderStep(idx);
  bSkip.onclick = async () => {
    if (passed) return;
    if (!(await ask({ title: tt('Skip this step'), text: tt('Skip this step? You will not earn XP for it, and you can come back to it at any time.'), ok: tt('Skip step'), cancel: true }))) return;
    store.skip(id, idx); skipped.add(idx); Sound.play('blocked');
    if (idx < n - 1) renderStep(idx + 1); else finish();
  };
  bSol.onclick = () => {
    const c = $('#c-sol');
    if (!c.hidden) { c.hidden = true; return; }
    c.innerHTML = `<span class="callout-t">${ico.eye}Answer</span><div data-notranslate>${solutionHtml()}</div><button class="btn btn-secondary btn-sm" id="b-fill">Fill it in for me</button><span class="muted small"> No XP is awarded for assisted steps.</span>`;
    c.hidden = false; c.querySelector('#b-fill').onclick = applySolution;
  };
  app.querySelectorAll('.seg').forEach((b) => { b.onclick = () => { const i = +b.dataset.i; if (!b.disabled && i !== idx) renderStep(i); }; });

  /* global keys: keycap highlighting + Enter to continue */
  const kd = (e) => {
    if (step && step.keys) { const lab = keyLabel(e); wk.querySelectorAll('.cap').forEach((c) => { if (c.dataset.k === lab || (lab === 'Ctrl' && c.dataset.k === 'Ctrl')) c.classList.add('down'); }); }
    if (e.key === 'Enter' && passed && !e.target.closest('input,textarea,button,a')) { e.preventDefault(); bNext.click(); }
  };
  const ku = (e) => { const lab = keyLabel(e); wk.querySelectorAll('.cap').forEach((c) => { if (c.dataset.k === lab) c.classList.remove('down'); }); if (e.key === 'Control' || e.key === 'Meta') wk.querySelectorAll('.cap').forEach((c) => c.classList.remove('down')); };
  const clearCaps = () => wk.querySelectorAll('.cap').forEach((c) => c.classList.remove('down'));
  document.addEventListener('keydown', kd, true); document.addEventListener('keyup', ku, true); window.addEventListener('blur', clearCaps);
  teardownLesson = () => { document.removeEventListener('keydown', kd, true); document.removeEventListener('keyup', ku, true); window.removeEventListener('blur', clearCaps); clearTimeout(flashT); if (grid) grid.destroy(); };

  /* completion */
  function finish() {
    const first = !store.d.completed[id], bonus = first && skipped.size === 0 ? (L.kind === 'test' ? 60 : 30) : 0;
    if (first) { store.complete(id, bonus); gained += bonus; }
    const nxt = LESSONS[LESSONS.indexOf(L) + 1], li = levelInfo(store.d.xp), solved = store.doneSet(id).size;
    const ov = document.createElement('div'); ov.className = 'overlay finish';
    ov.innerHTML = `<div class="modal celebrate" role="dialog" aria-modal="true">
      <div class="burst" aria-hidden="true">${Array.from({ length: 14 }, (_, i) => `<i style="--a:${i * 25.7}deg;--d:${i % 3}"></i>`).join('')}<span class="big-check">${ico.check}</span></div>
      <h2>Lesson complete!</h2><p class="muted">${L.title} · ${solved}/${n} steps solved</p>${skipped.size ? `<p class="muted small">${skipped.size === 1 ? 'You skipped 1 step. You can come back to it at any time.' : `You skipped ${skipped.size} steps. You can come back to them at any time.`}</p>` : ''}
      <div class="rewards"><div><b>+${gained}</b><span>XP earned</span></div><div><b>${store.d.xp}</b><span>Total XP</span></div><div><b>Lv ${li.n}</b><span>${li.name}</span></div></div>
      ${first && bonus ? `<p class="bonus">Includes a +${bonus} XP ${L.kind === 'test' ? 'test' : 'lesson'} bonus.</p>` : ''}
      <div class="modal-actions">${nxt && !E.has(nxt.id) ? `<a class="btn btn-primary btn-lg" href="#/lesson/${nxt.id}">${ico.lock}Unlock to continue</a>` : nxt ? `<a class="btn btn-primary btn-lg" href="#/lesson/${nxt.id}">Next: ${esc(nxt.title)} ${ico.arrow}</a>` : '<a class="btn btn-primary btn-lg" href="#/">You finished the course! Back to start</a>'}
      <a class="btn btn-ghost" href="#/">All lessons</a></div></div>`;
    Sound.play('complete');
    document.body.appendChild(ov); ov.querySelector('.btn-primary').focus();
    ov.addEventListener('click', (e) => { if (e.target.closest('a')) ov.remove(); });
  }

  renderStep(Math.min(n - 1, startAt != null ? startAt : firstUndone()));
}

/* ───────── PAYWALL ───────── */
const paGrid = () => {
  const rows = [['Q1', 120, 80], ['Q2', 135, 95], ['Q3', 160, 110], ['Q4', 150, 140], ['Q5', 180, 120], ['Q6', 210, 165]];
  const head = ['', 'A', 'B', 'C'].map((h, k) => `<b class="pa-h" style="--i:${k}">${h}</b>`).join('');
  return head + rows.map((r, i) => `<b class="pa-h" style="--i:${4 + i * 4}">${i + 1}</b>` + r.map((v, k) => `<span class="pa-c${k ? ' n' : ''}" style="--i:${5 + i * 4 + k}">${v}</span>`).join('')).join('');
};
function paywallView(L) {
  setTitle('Unlock the full course · Spreadsheet Academy');
  const steps = LESSONS.reduce((a, l) => a + l.steps.length, 0);
  app.innerHTML = `${topbar('courses')}<main class="wrap paywall"><section class="pay-card"><div class="pay-copy">
    <span class="pay-ico">${ico.lock}</span>
    <div class="pay-lesson"><span>Locked lesson</span><b data-notranslate>${esc(L.title)}</b></div>
    <h1>Unlock the full course</h1>
    <p class="lead">The first two lessons are free. Unlock the rest, including every path and real-world test, with a single purchase.</p>
    <ul class="points pay-points"><li>${ico.check}<span><b>${LESSONS.length}</b> hands-on lessons in all</span></li><li>${ico.check}<span><b>${steps}</b> checked exercises</span></li><li>${ico.check}<span>Real-world tests and a capstone project</span></li><li>${ico.check}<span>One-time purchase, yours to keep</span></li><li>${ico.check}<span>Works offline once unlocked</span></li></ul>
    <div class="pay-cta">${E.web ? `<a class="btn btn-primary btn-lg" id="buy-web" href="${E.storeHref() || '#'}" rel="noopener">${ico.lock}Get the full course on the Microsoft Store</a>` : `<button class="btn btn-primary btn-lg" id="buy">${ico.lock}Unlock full course</button>`}<a class="btn btn-ghost btn-lg" href="#/courses">Back to courses</a></div>
    ${E.web ? '' : '<button class="link" id="restore">Restore purchase</button>'}<p class="pay-msg muted" id="pay-msg" role="status"></p>
  </div><div class="pay-art" aria-hidden="true"><div class="pa-scene">
      <div class="pa-sheet">
        <div class="pa-fbar"><i>fx</i><span class="pa-f">=SUM(B2:B7)</span></div>
        <div class="pa-grid">${paGrid()}<span class="pa-sel"></span></div>
      </div>
      <div class="pa-chart"><svg viewBox="0 0 150 96"><g class="pa-bars">${[34, 52, 44, 68, 58, 80].map((h, k) => `<rect x="${8 + k * 23}" y="${88 - h}" width="15" height="${h}" rx="2" style="--k:${k}"/>`).join('')}</g><path class="pa-line" d="M15 62 L38 46 L61 54 L84 30 L107 38 L130 14"/><g class="pa-dots">${[[15, 62], [38, 46], [61, 54], [84, 30], [107, 38], [130, 14]].map(([x, y], k) => `<circle cx="${x}" cy="${y}" r="3" style="--k:${k}"/>`).join('')}</g></svg></div>
      <div class="pa-pie"><svg viewBox="0 0 64 64"><circle class="p0" cx="32" cy="32" r="22"/><circle class="p1" cx="32" cy="32" r="22"/><circle class="p2" cx="32" cy="32" r="22"/></svg></div>
      <span class="pa-chip c1">+10 XP</span><span class="pa-chip c2">${ico.check} Correct!</span>
    </div></div></section></main>${footer()}`;
  const msg = (m) => { document.getElementById('pay-msg').textContent = tt(m); };
  const bw = document.getElementById('buy-web');
  if (bw) bw.onclick = (e) => { if (window.gtag) window.gtag('event', 'store_click', { transport_type: 'beacon', from: 'app_demo' }); if (!E.storeHref()) { e.preventDefault(); msg('The full course is a Windows app. Open this page on a Windows PC to get it.'); } };
  const bb = document.getElementById('buy');
  if (bb) bb.onclick = async () => {
    const r = await E.buy();
    if (r === 'unavailable') msg('Purchases are available in the Microsoft Store version of the app.'); else if (r === 'cancelled') msg('The purchase was not completed.');
  };
  const br = document.getElementById('restore');
  if (br) br.onclick = async () => {
    const r = await E.restore();
    if (r === 'unavailable') msg('Purchases are available in the Microsoft Store version of the app.'); else if (r === 'none') msg('No purchase found for this account.');
  };
  wireCommon();
}

/* ───────── router ───────── */
function route() {
  if (teardownLesson) { teardownLesson(); teardownLesson = null; }
  document.querySelectorAll('.overlay').forEach((o) => o.remove());
  const m = /^#\/lesson\/([\w-]+)(?:\/(\d+))?/.exec(location.hash);
  window.scrollTo(0, 0);
  const target = m && LESSONS.find((l) => l.id === m[1]);
  if (target && !E.has(target.id)) paywallView(target);
  else if (m) lessonView(m[1], m[2] ? +m[2] - 1 : null);
  else if (location.hash === '#/courses') catalogView();
  else if (location.hash === '#/welcome' || !hasProgress()) landingView();
  else dashboardView();
}
store.load();
I18N.applyLessons();
document.addEventListener('change', (e) => { const s = e.target.closest && e.target.closest('[data-lang]'); if (!s) return; I18N.set(s.value); route(); });
document.addEventListener('click', (e) => {
  const b = e.target.closest('[data-sound-toggle]'); if (!b) return;
  const on = window.Sound.toggle();
  document.querySelectorAll('[data-sound-toggle]').forEach((x) => { x.innerHTML = on ? ico.sound : ico.mute; x.title = on ? 'Sound on' : 'Sound off'; x.setAttribute('aria-label', on ? 'Mute sounds' : 'Unmute sounds'); });
});
const closeSettings = () => { const p = document.getElementById('set-panel'), g = document.getElementById('gear'); if (p) p.hidden = true; if (g) g.setAttribute('aria-expanded', 'false'); };
document.addEventListener('click', (e) => {
  const g = e.target.closest('#gear');
  if (g) { const p = document.getElementById('set-panel'), open = p.hidden; p.hidden = !open; g.setAttribute('aria-expanded', String(open)); return; }
  const sw = e.target.closest('[data-theme-toggle]');
  if (sw) { theme.set(!theme.dark); sw.setAttribute('aria-checked', String(theme.dark)); return; }
  if (!e.target.closest('#set-panel') || e.target.closest('.set-item')) closeSettings();
});
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeSettings(); });
window.addEventListener('hashchange', route);
E.onChange(route);
route();
E.refresh();
})();
