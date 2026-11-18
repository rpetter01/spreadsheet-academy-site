/* Spreadsheet Academy i18n: English (source) + Português (Brasil) + Español.
   - UI text is matched by its English source string (exact or pattern) and replaced in the DOM.
   - Lessons are translated in place from js/l10n-*.js (falls back to English per step).
   - Feedback messages from exercise checkers go through I18N.msg().
   window.I18N */
(function (root) {
'use strict';
const LANGS = [['en', 'English'], ['pt', 'Português (Brasil)'], ['es', 'Español']];
const HTML_LANG = { en: 'en', pt: 'pt-BR', es: 'es' };
const KEY = 'sa.lang.v1';
let lang = 'en';
try {
  const s = localStorage.getItem(KEY);
  if (s && LANGS.some((l) => l[0] === s)) lang = s;
  else { const n = (navigator.language || 'en').toLowerCase(); lang = n.startsWith('pt') ? 'pt' : n.startsWith('es') ? 'es' : 'en'; }
} catch (e) { /* storage unavailable */ }

const D = { pt: new Map(), es: new Map() };   // exact English -> translation
const P = { pt: [], es: [] };                  // patterns [RegExp, template]
const LT = { pt: {}, es: {} };                 // lesson translations
const ui = (rows) => rows.forEach(([en, pt, es]) => { D.pt.set(en, pt); D.es.set(en, es); });
const pat = (re, pt, es) => { P.pt.push([re, pt]); P.es.push([re, es]); };

function fill(tpl, m) { return tpl.replace(/\$(t?)(\d)/g, (_, t, i) => (t ? tx(m[i] || '') : m[i] || '')); }
/** Translate one plain string (exact match, then patterns). */
function tx(s) {
  if (lang === 'en' || !s) return s;
  const d = D[lang]; if (d.has(s)) return d.get(s);
  for (const [re, tpl] of P[lang]) { const m = re.exec(s); if (m) return fill(tpl, m); }
  return s;
}
/** Translate with {placeholders}. */
function t(key, params) {
  let out = tx(key);
  if (params) out = out.replace(/\{(\w+)\}/g, (_, k) => (k in params ? params[k] : `{${k}}`));
  return out;
}
const msg = (s) => tx(s);

/* ───────── DOM translation ───────── */
const SKIP = '.xl,code,kbd,script,style,.keys,.lang-sel,[data-notranslate]';
const inSkip = (el) => !el || !!el.closest(SKIP);
function txtNode(n) {
  const v = n.nodeValue, s = v.trim(); if (!s) return;
  const r = tx(s); if (r !== s) n.nodeValue = v.replace(s, r);
}
const ATTRS = ['title', 'aria-label', 'placeholder'];
function attrs(el) { if (inSkip(el)) return; ATTRS.forEach((a) => { const v = el.getAttribute && el.getAttribute(a); if (v) { const r = tx(v); if (r !== v) el.setAttribute(a, r); } }); }
function walk(node) {
  if (lang === 'en') return;
  if (node.nodeType === 3) { if (!inSkip(node.parentElement)) txtNode(node); return; }
  if (node.nodeType !== 1 || inSkip(node)) return;
  const w = document.createTreeWalker(node, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (inSkip(n.parentElement) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) });
  const list = []; while (w.nextNode()) list.push(w.currentNode); list.forEach(txtNode);
  attrs(node); node.querySelectorAll('[title],[aria-label],[placeholder]').forEach(attrs);
}
function observe() {
  if (!root.MutationObserver) return;
  new MutationObserver((muts) => {
    if (lang === 'en') return;
    for (const m of muts) {
      if (m.type === 'childList') m.addedNodes.forEach(walk);
      else if (m.type === 'characterData') { if (!inSkip(m.target.parentElement)) txtNode(m.target); }
      else if (m.type === 'attributes') attrs(m.target);
    }
  }).observe(document.body, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ATTRS });
}

/* ───────── sheet glossary ─────────
   The text inside the practice spreadsheets (headers, months, categories) follows the app language, like the lesson text does.
   gloss([[english, português, español], ...]) registers whole-cell words. word() returns the local one (or the English when none is known). */
const GL = { pt: new Map(), es: new Map() };
const gloss = (rows) => rows.forEach(([en, pt, es]) => { GL.pt.set(en, pt); GL.es.set(en, es); });
const word = (s) => (lang === 'en' || typeof s !== 'string' ? s : (GL[lang].has(s) ? GL[lang].get(s) : s));
/** a cell value: plain text is translated whole; in a formula only the "string literals" are */
const wordRaw = (v) => (typeof v !== 'string' || lang === 'en' ? v : v[0] === '=' ? v.replace(/"((?:[^"]|"")*)"/g, (m, x) => '"' + word(x) + '"') : word(v));
/** in a lesson text, the "string literals" inside <code> formulas follow the glossary too (=SUMIF(B:B,"West",...) -> "Oeste") */
const wordHtml = (h) => (lang === 'en' || !h ? h : h.replace(/<code>([\s\S]*?)<\/code>/g, (m, c) => '<code>' + c.replace(/"((?:[^"<]|"")*)"/g, (q, x) => '"' + word(x) + '"') + '</code>'));
/** A formula as it is typed in Excel in this language (Português, Español): local function names, ";" between arguments and "," as the decimal mark,
    e.g. =IF(B2>10000,B2*0.1,0) -> =SE(B2>10000;B2*0,1;0). Text between quotes is left alone. The engine accepts both forms. */
function localFormula(f) {
  const XL = root.XL; if (lang === 'en' || typeof f !== 'string' || f[0] !== '=' || !XL || !XL.localName) return f;
  const TF = lang === 'pt' ? { TRUE: 'VERDADEIRO', FALSE: 'FALSO' } : { TRUE: 'VERDADERO', FALSE: 'FALSO' };
  return f.split(/("(?:[^"]|"")*")/).map((p, k) => (k % 2 ? p : p
    .replace(/\b([A-Z][A-Z0-9.]*)(?=\()/g, (m) => XL.localName(m, lang))
    .replace(/\b(TRUE|FALSE)\b/g, (m) => TF[m])
    .replace(/,/g, ';')
    .replace(/(\d)\.(\d)/g, '$1,$2'))).join('');
}
/** the formulas in <code> spans of a lesson text (the ones starting with =) in the typing style of the language */
const formulaHtml = (h) => (lang === 'en' || !h ? h : h.replace(/<code>(=[\s\S]*?)<\/code>/g, (m, c) => '<code>' + localFormula(c) + '</code>'));
const mapCells = (o) => { if (!o || typeof o !== 'object') return o; const r = {}; Object.keys(o).forEach((k) => { r[k] = wordRaw(o[k]); }); return r; };
const cloneGrid = (g) => (g ? { data: g.data ? Object.assign({}, g.data) : undefined, sheets: g.sheets ? JSON.parse(JSON.stringify(g.sheets)) : undefined } : null);

/* ───────── lessons ───────── */
const EN = new WeakMap();
/** a task that says "fill down": step.fill = { range: 'C2:C6' } when the first and last cell are named, else { range: null } */
function fillInfo(task) {
  if (!task || !/fill (?:it )?down/i.test(task)) return null;
  const a = /\bIn <b>([A-Z]+)(\d+)<\/b>/.exec(task), z = /fill (?:it )?down to <b>([A-Z]+)(\d+)<\/b>/i.exec(task) || /fill (?:it )?down to ([A-Z]+)(\d+)/i.exec(task);
  return { range: a && z && a[1] === z[1] && +z[2] > +a[2] ? a[1] + a[2] + ':' + z[1] + z[2] : null };
}
function captureEnglish() {
  const C = root.COURSE; if (!C) return;
  C.LESSONS.forEach((l) => {
    if (EN.has(l)) return;
    l.steps.forEach((s) => { const f = fillInfo(s.task); if (f) s.fill = f; });
    EN.set(l, { title: l.title, summary: l.summary, steps: l.steps.map((s) => ({ title: s.title, body: s.body, task: s.task, hint: s.hint, explain: s.quiz && s.quiz.explain, soltext: s.sol && s.sol.text, options: s.quiz && s.quiz.options.slice(), grid: cloneGrid(s.grid), cells: s.sol && s.sol.cells ? Object.assign({}, s.sol.cells) : null })) });
  });
}
const mk = (s) => (s == null ? s : s.replace(/\{\{([\s\S]*?)\}\}/g, '<code>$1</code>'));
/** Português / Español: the first time a function appears in a lesson's text (course order), add its local name after it, e.g. COUNT (em português, CONT.NÚM) */
function nameFirstMentions() {
  const C = root.COURSE, XL = root.XL; if (lang === 'en' || !C || !XL || !XL.localName) return;
  const seen = new Set(), tpl = lang === 'pt' ? 'em português' : 'en español', and = lang === 'pt' ? ' e ' : ' y ';
  const strip = (s) => s.replace(/<[^>]+>/g, ' ');
  C.TRACKS.forEach((tr) => (tr.lessons || []).forEach((l) => l.steps.forEach((s) => ['body', 'task'].forEach((f) => {
    if (!s[f]) return; const plain = strip(s[f]);
    s[f] = s[f].replace(/<code>([\s\S]*?)<\/code>/g, (m, inner) => {
      const found = [];
      const add = (n) => { if (!n || seen.has(n) || found.includes(n)) return; const loc = XL.localName(n, lang); if (loc === n) { seen.add(n); return; } seen.add(n); if (!new RegExp('(^|[^A-Za-zÀ-ÿ0-9.])' + loc.replace(/\./g, '\\.') + '([^A-Za-zÀ-ÿ0-9.]|$)').test(plain)) found.push(n); };
      const t0 = inner.trim(); if (/^[A-Z][A-Z0-9.]*$/.test(t0)) add(t0);
      (inner.match(/\b[A-Z][A-Z0-9.]*(?=\()/g) || []).forEach(add);
      return found.length && inner.trim()[0] !== '=' ? m + ' (' + tpl + ', ' + found.map((n) => '<code>' + XL.localName(n, lang) + '</code>').join(and) + ')' : m;
    });
  }))));
}
function applyLessons() {
  const C = root.COURSE; if (!C) return; captureEnglish();
  C.LESSONS.forEach((l) => {
    const en = EN.get(l), tr = lang === 'en' ? null : LT[lang][l.id];
    l.title = tr && tr.title ? tr.title : en.title; l.summary = tr && tr.summary ? tr.summary : en.summary;
    l.steps.forEach((s, i) => {
      const e = en.steps[i], x = tr && tr.steps && tr.steps[i] && tr.steps[i][0] === e.title ? tr.steps[i] : null;
      s.title = x && x[1] ? x[1] : e.title; s.body = x && x[2] ? mk(x[2]) : e.body; s.task = x && x[3] ? mk(x[3]) : e.task;
      if (e.hint !== undefined) s.hint = x && x[4] ? mk(x[4]) : e.hint;
      if (s.quiz) { s.quiz.explain = x && x[5] ? mk(x[5]) : e.explain; s.quiz.options = x && x[7] ? x[7].map(mk) : e.options.slice(); }
      if (s.sol && e.soltext !== undefined) s.sol.text = x && x[6] ? mk(x[6]) : e.soltext;
      /* the text inside the practice sheet, and the cell values "Show answer" types, in the app language */
      if (e.grid) { if (e.grid.data) s.grid.data = mapCells(e.grid.data); if (e.grid.sheets) { const sh = {}; Object.keys(e.grid.sheets).forEach((nm) => { sh[nm] = mapCells(e.grid.sheets[nm]); }); s.grid.sheets = sh; } }
      if (e.cells && s.sol) s.sol.cells = mapCells(e.cells);
      if (lang !== 'en') { s.body = wordHtml(s.body); s.task = wordHtml(s.task); if (s.hint !== undefined) s.hint = wordHtml(s.hint); if (s.quiz) { s.quiz.explain = wordHtml(s.quiz.explain); s.quiz.options = s.quiz.options.map(wordHtml); } if (s.sol && s.sol.text) s.sol.text = wordHtml(s.sol.text); }
    });
  });
  nameFirstMentions();
  if (lang !== 'en') C.LESSONS.forEach((l) => l.steps.forEach((s) => {
    s.body = formulaHtml(s.body); s.task = formulaHtml(s.task); if (s.hint !== undefined) s.hint = formulaHtml(s.hint);
    if (s.quiz) { s.quiz.explain = formulaHtml(s.quiz.explain); s.quiz.options = s.quiz.options.map(formulaHtml); }
    if (s.sol && s.sol.text) s.sol.text = formulaHtml(s.sol.text);
  }));
  (C.TRACKS || []).forEach((t) => {
    if (!t._en) t._en = { name: t.name, level: t.level, tagline: t.tagline };
    t.name = lang === 'en' ? t._en.name : tx(t._en.name); t.level = lang === 'en' ? t._en.level : tx(t._en.level); t.tagline = lang === 'en' ? t._en.tagline : tx(t._en.tagline);
  });
}
function addLessons(code, map) { Object.assign(LT[code], map); }
/** Coverage report used by tests: [translatedSteps, totalSteps]. */
function coverage(code) {
  const C = root.COURSE; captureEnglish(); let ok = 0, all = 0, bad = [];
  C.LESSONS.forEach((l) => { const en = EN.get(l), tr = LT[code][l.id]; l.steps.forEach((s, i) => { all++; const x = tr && tr.steps && tr.steps[i]; if (x && x[0] === en.steps[i].title && x[1] && x[2] && x[3]) ok++; else bad.push(l.id + '#' + (i + 1)); }); });
  return { ok, all, bad };
}

function set(code) {
  if (!LANGS.some((l) => l[0] === code)) return;
  lang = code; try { localStorage.setItem(KEY, code); } catch (e) { /* ignore */ }
  document.documentElement.lang = HTML_LANG[code]; applyLessons();
}

/* ═════════════ UI DICTIONARY ═════════════ */
ui([
  // levels & tracks
  ['Beginner', 'Iniciante', 'Principiante'], ['Intermediate', 'Intermediário', 'Intermedio'], ['Advanced', 'Avançado', 'Avanzado'],
  ['Rookie', 'Novato', 'Novato'], ['Apprentice', 'Aprendiz', 'Aprendiz'], ['Analyst', 'Analista', 'Analista'], ['Specialist', 'Especialista', 'Especialista'], ['Strategist', 'Estrategista', 'Estratega'], ['Wizard', 'Mago', 'Mago'], ['Grandmaster', 'Grão-mestre', 'Gran maestro'],
  ['Foundations', 'Fundamentos', 'Fundamentos'], ['Data Skills', 'Habilidades com Dados', 'Habilidades con Datos'], ['Tables & Reporting', 'Tabelas e Relatórios', 'Tablas e Informes'], ['Power User', 'Usuário Avançado', 'Usuario Avanzado'],
  ['Never opened Excel? Start here. Cells, formulas, shortcuts and your first functions, finishing with a real-world test.', 'Nunca abriu o Excel? Comece aqui. Células, fórmulas, atalhos e suas primeiras funções, terminando com um teste do mundo real.', '¿Nunca abriste Excel? Empieza aquí. Celdas, fórmulas, atajos y tus primeras funciones, y al final una prueba del mundo real.'],
  ['Make spreadsheets think: logic, counting, text, dates, lookups and statistics, finishing with a retail audit.', 'Faça as planilhas pensarem: lógica, contagem, texto, datas, procuras e estatística, terminando com uma auditoria de varejo.', 'Haz que las hojas piensen: lógica, conteo, texto, fechas, búsquedas y estadística, y al final una auditoría de comercio minorista.'],
  ['Turn data into reports: worksheets, conditional formatting, filters, sorting and PivotTables.', 'Transforme dados em relatórios: planilhas, formatação condicional, filtros, classificação e Tabelas Dinâmicas.', 'Convierte datos en informes: hojas, formato condicional, filtros, ordenación y Tablas Dinámicas.'],
  ['Nested logic, multi-criteria analysis, arrays, dynamic arrays and expert lookups, with a sales capstone and two real-world case studies.', 'Lógica aninhada, análise com vários critérios, matrizes, matrizes dinâmicas e procuras avançadas, com um projeto final de vendas e dois estudos de caso reais.', 'Lógica anidada, análisis con varios criterios, matrices, matrices dinámicas y búsquedas avanzadas, con un proyecto final de ventas y dos casos de estudio reales.'],
  // skill areas
  ['Cells & formulas', 'Células e fórmulas', 'Celdas y fórmulas'], ['Shortcuts & speed', 'Atalhos e velocidade', 'Atajos y velocidad'], ['Logic & decisions', 'Lógica e decisões', 'Lógica y decisiones'], ['Counting & summing', 'Contagem e soma', 'Conteo y suma'],
  ['Text & dates', 'Texto e datas', 'Texto y fechas'], ['Lookups', 'Procuras', 'Búsquedas'], ['Stats & rounding', 'Estatística e arredondamento', 'Estadística y redondeo'], ['Arrays & dynamic', 'Matrizes e dinâmicas', 'Matrices y dinámicas'],
  ['Tables & reporting', 'Tabelas e relatórios', 'Tablas e informes'], ['Real-world tests', 'Testes do mundo real', 'Pruebas del mundo real'],
  // navigation & chrome
  ['Home', 'Início', 'Inicio'], ['My learning', 'Meu aprendizado', 'Mi aprendizaje'], ['Courses', 'Cursos', 'Cursos'], ['Browse courses', 'Ver cursos', 'Ver cursos'],
  ['day streak', 'dias seguidos', 'días seguidos'], ['XP', 'XP', 'XP'], ["Days in a row you've practised", 'Dias seguidos em que você praticou', 'Días seguidos en los que practicaste'], ['Experience points', 'Pontos de experiência', 'Puntos de experiencia'], ['Max level', 'Nível máximo', 'Nivel máximo'],
  ['Mute sounds', 'Silenciar sons', 'Silenciar sonidos'], ['Unmute sounds', 'Ativar sons', 'Activar sonidos'], ['Sound on', 'Som ligado', 'Sonido activado'], ['Sound off', 'Som desligado', 'Sonido desactivado'],
  ['Spreadsheet Academy home', 'Página inicial do Spreadsheet Academy', 'Inicio de Spreadsheet Academy'], ['Main', 'Principal', 'Principal'],
  // landing
  ['Spreadsheet Academy · Learn Excel by doing', 'Spreadsheet Academy · Aprenda Excel na prática', 'Spreadsheet Academy · Aprende Excel haciendo'],
  ['My learning · Spreadsheet Academy', 'Meu aprendizado · Spreadsheet Academy', 'Mi aprendizaje · Spreadsheet Academy'], ['Courses · Spreadsheet Academy', 'Cursos · Spreadsheet Academy', 'Cursos · Spreadsheet Academy'],
  ['Hands-on Excel training', 'Treinamento prático de Excel', 'Entrenamiento práctico de Excel'], ['Learn Excel by', 'Aprenda Excel', 'Aprende Excel'], ['actually using it.', 'usando de verdade.', 'usándolo de verdad.'],
  ['Short lessons with a live spreadsheet in every step. Type formulas, press real shortcuts and get instant feedback, from your very first cell to lookups and array formulas.', 'Aulas curtas com uma planilha de verdade em cada etapa. Digite fórmulas, use atalhos reais e receba feedback na hora, da sua primeira célula até procuras e fórmulas de matriz.', 'Lecciones cortas con una hoja de cálculo real en cada paso. Escribe fórmulas, usa atajos reales y recibe respuesta al instante, desde tu primera celda hasta búsquedas y fórmulas matriciales.'],
  ['Start learning free', 'Começar grátis', 'Empezar gratis'], ['Find my level', 'Descobrir meu nível', 'Encontrar mi nivel'], ['No install, no signup', 'Sem instalar, sem cadastro', 'Sin instalar, sin registro'], ['Keyboard-first practice', 'Prática focada no teclado', 'Práctica centrada en el teclado'], ['Checked as you type', 'Verificado enquanto você digita', 'Verificado mientras escribes'],
  ['hands-on lessons', 'aulas práticas', 'lecciones prácticas'], ['checked exercises', 'exercícios verificados', 'ejercicios verificados'], ['Excel functions in the live sheet', 'funções do Excel na planilha ao vivo', 'funciones de Excel en la hoja en vivo'], ['learning paths', 'trilhas de aprendizado', 'rutas de aprendizaje'],
  ['How it works', 'Como funciona', 'Cómo funciona'], ["Learn the way you'll actually work", 'Aprenda do jeito que você vai trabalhar de verdade', 'Aprende como vas a trabajar de verdad'],
  ['Read a short idea', 'Leia uma ideia curta', 'Lee una idea corta'], ['One concept per step, explained in a few sentences. No videos to sit through.', 'Um conceito por etapa, explicado em poucas frases. Sem vídeos longos para assistir.', 'Un concepto por paso, explicado en pocas frases. Sin videos largos que ver.'],
  ['Do it in a live sheet', 'Faça em uma planilha ao vivo', 'Hazlo en una hoja en vivo'], ['Type real formulas and press real shortcuts in a spreadsheet that calculates as you go.', 'Digite fórmulas reais e use atalhos reais em uma planilha que calcula enquanto você trabalha.', 'Escribe fórmulas reales y usa atajos reales en una hoja que calcula sobre la marcha.'],
  ['Get instant feedback', 'Receba feedback na hora', 'Recibe respuesta al instante'], ["Know straight away what's right, what's off and why. Earn XP and build a streak.", 'Saiba na hora o que está certo, o que está errado e por quê. Ganhe XP e crie uma sequência.', 'Sabe al instante qué está bien, qué falla y por qué. Gana XP y crea una racha.'],
  ['Learning paths', 'Trilhas de aprendizado', 'Rutas de aprendizaje'], ['Start where you are', 'Comece de onde você está', 'Empieza desde donde estás'], ['Pick a path. Every lesson is open from day one.', 'Escolha uma trilha. Todas as aulas estão liberadas desde o primeiro dia.', 'Elige una ruta. Todas las lecciones están abiertas desde el primer día.'],
  ['Start path', 'Começar trilha', 'Empezar ruta'], ['Ready to stop guessing at formulas?', 'Pronto para parar de chutar fórmulas?', '¿Listo para dejar de adivinar fórmulas?'], ['Your first lesson takes five minutes.', 'Sua primeira aula leva cinco minutos.', 'Tu primera lección dura cinco minutos.'], ['Start the first lesson', 'Começar a primeira aula', 'Empezar la primera lección'],
  ['is an independent practice simulator and is not affiliated with Microsoft. Excel is a trademark of Microsoft Corporation.', 'é um simulador de prática independente e não é afiliado à Microsoft. Excel é uma marca registrada da Microsoft Corporation.', 'es un simulador de práctica independiente y no está afiliado a Microsoft. Excel es una marca registrada de Microsoft Corporation.'],
  ['Reset my progress', 'Zerar meu progresso', 'Reiniciar mi progreso'], ['Erase all progress, XP and streak on this device?', 'Apagar todo o progresso, XP e sequência neste dispositivo?', '¿Borrar todo el progreso, XP y racha en este dispositivo?'],
  // dashboard
  ['Welcome back!', 'Bem-vindo de volta!', '¡Bienvenido de nuevo!'], ['Continue learning', 'Continuar aprendendo', 'Seguir aprendiendo'], ['Keep making progress', 'Continuar progredindo', 'Seguir avanzando'], ['Start lesson', 'Começar aula', 'Empezar lección'],
  ['Course complete', 'Curso concluído', 'Curso completado'], ['You finished every lesson!', 'Você concluiu todas as aulas!', '¡Terminaste todas las lecciones!'], ['Review any lesson to keep your skills sharp.', 'Revise qualquer aula para manter suas habilidades afiadas.', 'Repasa cualquier lección para mantener tus habilidades al día.'],
  ['Assess', 'Avaliar', 'Evaluar'], ['Practice', 'Praticar', 'Practicar'], ['Apply', 'Aplicar', 'Aplicar'], ['Keyboard shortcut drills', 'Treino de atalhos de teclado', 'Práctica de atajos de teclado'],
  ['In progress', 'Em andamento', 'En curso'], ['Up next', 'A seguir', 'A continuación'], ['All courses', 'Todos os cursos', 'Todos los cursos'], ['Nothing left. Great work!', 'Nada mais por fazer. Ótimo trabalho!', 'No queda nada. ¡Buen trabajo!'],
  ['Skills and XP', 'Habilidades e XP', 'Habilidades y XP'], ['Steps solved', 'Etapas resolvidas', 'Pasos resueltos'], ['Lessons done', 'Aulas concluídas', 'Lecciones completadas'], ['Streak', 'Sequência', 'Racha'], ['d', 'd', 'd'], ['Top level', 'Nível máximo', 'Nivel máximo'],
  // catalog & rows
  ['All paths', 'Todas as trilhas', 'Todas las rutas'], ['Start', 'Começar', 'Empezar'], ['Review', 'Revisar', 'Repasar'], ['Continue', 'Continuar', 'Continuar'], ['Real-world test', 'Teste do mundo real', 'Prueba del mundo real'],
  ['Where should you start?', 'Por onde você deve começar?', '¿Por dónde deberías empezar?'], ['Pick the sentence that sounds most like you.', 'Escolha a frase que mais se parece com você.', 'Elige la frase que más se parezca a ti.'], ['Close', 'Fechar', 'Cerrar'],
  ["I'm new to Excel", 'Sou novo no Excel', 'Soy nuevo en Excel'], ['Start with cells, formulas and keyboard basics.', 'Comece com células, fórmulas e o básico do teclado.', 'Empieza con celdas, fórmulas y lo básico del teclado.'],
  ['I can use SUM and simple formulas', 'Sei usar SOMA e fórmulas simples', 'Sé usar SUMA y fórmulas sencillas'], ['Jump to IF, conditional totals, text and lookups.', 'Vá direto para SE, totais condicionais, texto e procuras.', 'Salta a SI, totales condicionales, texto y búsquedas.'],
  ['I use IF and VLOOKUP, but want mastery', 'Uso SE e PROCV, mas quero dominar', 'Uso SI y BUSCARV, pero quiero dominarlo'], ['Go straight to nested logic, SUMIFS and arrays.', 'Vá direto para lógica aninhada, SOMASES e matrizes.', 'Ve directo a lógica anidada, SUMAR.SI.CONJUNTO y matrices.'],
  // lesson player
  ['Hint', 'Dica', 'Pista'], ['Show answer', 'Mostrar resposta', 'Mostrar respuesta'], ['Reset', 'Reiniciar', 'Reiniciar'], ['Restart this step', 'Reiniciar esta etapa', 'Reiniciar este paso'], ['Check my work', 'Verificar meu trabalho', 'Comprobar mi trabajo'], ['Finish lesson', 'Concluir aula', 'Terminar lección'],
  ['Back to all lessons', 'Voltar para todas as aulas', 'Volver a todas las lecciones'], ['All lessons', 'Todas as aulas', 'Todas las lecciones'], ['Steps', 'Etapas', 'Pasos'],
  ['Your task', 'Sua tarefa', 'Tu tarea'], ['Solved', 'Resolvido', 'Resuelto'], ['Challenge', 'Desafio', 'Desafío'], ['Answer', 'Resposta', 'Respuesta'], ['Why', 'Por quê', 'Por qué'],
  ['Choose an answer.', 'Escolha uma resposta.', 'Elige una respuesta.'], ['Use the keyboard to complete the task.', 'Use o teclado para concluir a tarefa.', 'Usa el teclado para completar la tarea.'], ['Complete the task, then press Check or Enter.', 'Conclua a tarefa e depois pressione Verificar ou Enter.', 'Completa la tarea y luego pulsa Comprobar o Enter.'],
  ['Editing is switched off for this exercise. Just navigate.', 'A edição está desativada neste exercício. Apenas navegue.', 'La edición está desactivada en este ejercicio. Solo navega.'], ['Keyboard only here. Use the keys shown above the sheet.', 'Aqui é só teclado. Use as teclas mostradas acima da planilha.', 'Aquí solo teclado. Usa las teclas mostradas sobre la hoja.'],
  ['Not quite. Think it through and try another answer.', 'Quase. Pense com calma e tente outra resposta.', 'Casi. Piénsalo bien e inténtalo con otra respuesta.'], ['Nothing to check yet. Finish the task, then try again.', 'Ainda não há nada para verificar. Conclua a tarefa e tente de novo.', 'Aún no hay nada que comprobar. Termina la tarea y vuelve a intentarlo.'],
  ['Look at the keys above the sheet and try them in order.', 'Olhe as teclas acima da planilha e tente-as em ordem.', 'Mira las teclas sobre la hoja y pruébalas en orden.'], ['Re-read the task, then try it step by step.', 'Releia a tarefa e tente passo a passo.', 'Vuelve a leer la tarea e inténtalo paso a paso.'],
  ['Correct!', 'Correto!', '¡Correcto!'], ['Nailed it!', 'Mandou bem!', '¡Lo clavaste!'], ['Exactly right.', 'Exatamente isso.', 'Exactamente.'], ['Clean work.', 'Trabalho limpo.', 'Trabajo limpio.'], ["That's it!", 'É isso aí!', '¡Eso es!'], ['Spot on.', 'Na mosca.', 'Justo en el clavo.'],
  ['Answer filled in. No XP this time, but you can keep going.', 'Resposta preenchida. Sem XP desta vez, mas você pode continuar.', 'Respuesta completada. Sin XP esta vez, pero puedes continuar.'], ['Fill it in for me', 'Preencher para mim', 'Rellenar por mí'], [' No XP is awarded for assisted steps.', ' Etapas com ajuda não dão XP.', ' Los pasos con ayuda no dan XP.'], ['No XP is awarded for assisted steps.', 'Etapas com ajuda não dão XP.', 'Los pasos con ayuda no dan XP.'],
  ['Lesson complete!', 'Aula concluída!', '¡Lección completada!'], ['XP earned', 'XP ganho', 'XP ganado'], ['Total XP', 'XP total', 'XP total'], ['You finished the course! Back to start', 'Você concluiu o curso! Voltar ao início', '¡Terminaste el curso! Volver al inicio'],
  ['Language', 'Idioma', 'Idioma'],
  // tools
  ['Conditional Formatting', 'Formatação Condicional', 'Formato Condicional'], ['Home › Conditional Formatting › Highlight Cells Rules', 'Página Inicial › Formatação Condicional › Regras de Realce das Células', 'Inicio › Formato Condicional › Reglas para resaltar celdas'], ['Applies to', 'Aplica-se a', 'Se aplica a'], ['(your current selection)', '(sua seleção atual)', '(tu selección actual)'],
  ['Greater Than…', 'É Maior Que…', 'Mayor que…'], ['Less Than…', 'É Menor Que…', 'Menor que…'], ['Between…', 'Entre…', 'Entre…'], ['Equal To…', 'É Igual A…', 'Igual a…'], ['Text that Contains…', 'Texto que Contém…', 'Texto que contiene…'], ['Duplicate Values…', 'Valores Duplicados…', 'Valores duplicados…'],
  ['value', 'valor', 'valor'], ['and', 'e', 'y'], ['text', 'texto', 'texto'], ['with', 'com', 'con'], ['Light red fill with dark red text', 'Preenchimento vermelho claro com texto vermelho escuro', 'Relleno rojo claro con texto rojo oscuro'], ['Yellow fill with dark yellow text', 'Preenchimento amarelo com texto amarelo escuro', 'Relleno amarillo con texto amarillo oscuro'], ['Green fill with dark green text', 'Preenchimento verde com texto verde escuro', 'Relleno verde con texto verde oscuro'],
  ['Apply rule', 'Aplicar regra', 'Aplicar regla'], ['Clear rules', 'Limpar regras', 'Borrar reglas'], ['Enter a value first.', 'Digite um valor primeiro.', 'Escribe un valor primero.'], ['Enter both values.', 'Digite os dois valores.', 'Escribe ambos valores.'], ['Rules cleared.', 'Regras limpas.', 'Reglas borradas.'],
  ['Sort & Filter', 'Classificar e Filtrar', 'Ordenar y Filtrar'], ['Filter', 'Filtrar', 'Filtrar'], ['equals', 'é igual a', 'es igual a'], ['does not equal', 'é diferente de', 'es distinto de'], ['is greater than', 'é maior que', 'es mayor que'], ['is less than', 'é menor que', 'es menor que'], ['is at least', 'é pelo menos', 'es al menos'], ['is at most', 'é no máximo', 'es como máximo'], ['contains', 'contém', 'contiene'],
  ['Apply filter', 'Aplicar filtro', 'Aplicar filtro'], ['Clear', 'Limpar', 'Borrar'], ['Sort', 'Classificar', 'Ordenar'], ['A → Z / Smallest first', 'A → Z / Menor primeiro', 'A → Z / Menor primero'], ['Z → A / Largest first', 'Z → A / Maior primeiro', 'Z → A / Mayor primero'],
  ['PivotTable Fields', 'Campos da Tabela Dinâmica', 'Campos de la Tabla Dinámica'], ['Rows', 'Linhas', 'Filas'], ['Columns', 'Colunas', 'Columnas'], ['Values', 'Valores', 'Valores'], ['(none)', '(nenhum)', '(ninguno)'], ['Sum of', 'Soma de', 'Suma de'], ['Count of', 'Contagem de', 'Recuento de'], ['Average of', 'Média de', 'Promedio de'],
  ['Create PivotTable', 'Criar Tabela Dinâmica', 'Crear Tabla Dinámica'], ['Pick a numeric field for Values.', 'Escolha um campo numérico para Valores.', 'Elige un campo numérico para Valores.'], ['Cleared.', 'Limpo.', 'Borrado.'],
  ['percent', 'porcentagem', 'porcentaje'], ['currency', 'moeda', 'moneda'], ['date', 'data', 'fecha'], ['comma', 'milhar', 'millares'],
  // static checker messages
  ['Use parentheses so the addition happens first.', 'Use parênteses para que a soma seja feita primeiro.', 'Usa paréntesis para que la suma se haga primero.'],
  ['Use the cell references B1 and B2 instead of typing the numbers.', 'Use as referências de célula B1 e B2 em vez de digitar os números.', 'Usa las referencias de celda B1 y B2 en lugar de escribir los números.'],
  ['Lock the tax rate with <b>$B$1</b>, otherwise it will drift when you fill down.', 'Trave a alíquota com <b>$B$1</b>, senão ela vai se deslocar ao preencher para baixo.', 'Fija la tasa de impuesto con <b>$B$1</b>; si no, se desplazará al rellenar hacia abajo.'],
  ['Wrap it in VALUE() so the result is a number, not text.', 'Envolva em VALOR() para que o resultado seja um número, não texto.', 'Envuélvelo en VALOR() para que el resultado sea un número, no texto.'],
  ['Use INDEX with MATCH inside.', 'Use ÍNDICE com CORRESP dentro.', 'Usa INDICE con COINCIDIR dentro.'], ['Nest one IF inside another.', 'Aninhe um SE dentro de outro.', 'Anida un SI dentro de otro.'],
  ['Use the criteria in <b>H1</b> instead of typing "East".', 'Use o critério da célula <b>H1</b> em vez de digitar "East".', 'Usa el criterio de la celda <b>H1</b> en lugar de escribir "East".'],
  ['The result is right, but this exercise wants a specific function. Re-read the task.', 'O resultado está certo, mas este exercício pede uma função específica. Releia a tarefa.', 'El resultado es correcto, pero este ejercicio pide una función concreta. Vuelve a leer la tarea.'],
  ['The result is longer than expected. Check your condition.', 'O resultado é maior que o esperado. Verifique sua condição.', 'El resultado es más largo de lo esperado. Revisa tu condición.'],
  ['Clear B2:B5 only and leave the names in column A.', 'Limpe apenas B2:B5 e deixe os nomes na coluna A.', 'Borra solo B2:B5 y deja los nombres en la columna A.'],
  ['Extend the range so it ends at row 9.', 'Estenda o intervalo para que termine na linha 9.', 'Amplía el rango para que termine en la fila 9.'], ['Use a wildcard like "Back*" in an XLOOKUP.', 'Use um curinga como "Back*" em um PROCX.', 'Usa un comodín como "Back*" en un BUSCARX.'],
  ['Nest a MATCH inside the VLOOKUP.', 'Aninhe um CORRESP dentro do PROCV.', 'Anida un COINCIDIR dentro del BUSCARV.'], ['Use INDEX with MATCH(1, (cond1)*(cond2), 0).', 'Use ÍNDICE com CORRESP(1, (cond1)*(cond2), 0).', 'Usa INDICE con COINCIDIR(1, (cond1)*(cond2), 0).'],
  ['Return a real number, not text: wrap the MID in VALUE().', 'Retorne um número de verdade, não texto: envolva o EXT.TEXTO em VALOR().', 'Devuelve un número de verdad, no texto: envuelve el EXTRAE en VALOR().'], ['Put FILTER inside SUM.', 'Coloque FILTRO dentro de SOMA.', 'Pon FILTRAR dentro de SUMA.'],
  ['Reference Jan!B3, Feb!B3 and Mar!B3 inside AVERAGE.', 'Referencie Jan!B3, Feb!B3 e Mar!B3 dentro de MÉDIA.', 'Referencia Jan!B3, Feb!B3 y Mar!B3 dentro de PROMEDIO.'],
  ['The filter isn\'t showing the right records yet. On the Data tab turn on <b>Filter</b>, then use the drop-down arrow on the column header.', 'O filtro ainda não mostra os registros certos. Na guia Dados, ative <b>Filtro</b> e use a seta suspensa no cabeçalho da coluna.', 'El filtro aún no muestra los registros correctos. En la pestaña Datos, activa <b>Filtro</b> y usa la flecha desplegable del encabezado de la columna.'],
  ["Amounts aren't in descending order yet.", 'Os valores ainda não estão em ordem decrescente.', 'Los importes aún no están en orden descendente.'], ['Apply a filter first.', 'Aplique um filtro primeiro.', 'Aplica un filtro primero.'], ['Show only West, with Amount in largest-first order.', 'Mostre apenas West, com Amount em ordem do maior para o menor.', 'Muestra solo West, con Amount ordenado de mayor a menor.'],
  ["Return the item's name with a lookup, not by typing it.", 'Retorne o nome do item com uma procura, não digitando.', 'Devuelve el nombre del artículo con una búsqueda, no escribiéndolo.'], ['Use SUMIFS and point at the store name in J3.', 'Use SOMASES e aponte para o nome da loja em J3.', 'Usa SUMAR.SI.CONJUNTO y apunta al nombre de la tienda en J3.'],
  ['Combine MEDIAN with FILTER.', 'Combine MED com FILTRO.', 'Combina MEDIANA con FILTRAR.'], ['Use COUNTIFS for the patients and a lookup for the beds.', 'Use CONT.SES para os pacientes e uma procura para os leitos.', 'Usa CONTAR.SI.CONJUNTO para los pacientes y una búsqueda para las camas.'],
  ['Capstone: sales report', 'Projeto final: relatório de vendas', 'Proyecto final: informe de ventas'],
  ['Charts & dashboards', 'Gráficos e painéis', 'Gráficos y paneles'],
  ['Turn data into reports: worksheets, conditional formatting, filters, sorting, PivotTables and charts, finishing with a sales dashboard.', 'Transforme dados em relatórios: planilhas, formatação condicional, filtros, classificação, Tabelas Dinâmicas e gráficos, terminando com um painel de vendas.', 'Convierte datos en informes: hojas, formato condicional, filtros, ordenación, tablas dinámicas y gráficos, y al final un panel de ventas.'],
  ['Insert Chart', 'Inserir Gráfico', 'Insertar Gráfico'], ['Insert › Charts', 'Inserir › Gráficos', 'Insertar › Gráficos'], ['Data', 'Dados', 'Datos'], ['Chart type', 'Tipo de gráfico', 'Tipo de gráfico'],
  ['Column', 'Colunas', 'Columnas'], ['Bar', 'Barras', 'Barras'], ['Line', 'Linhas', 'Líneas'], ['Pie', 'Pizza', 'Circular'], ['column', 'colunas', 'columnas'], ['bar', 'barras', 'barras'], ['line', 'linhas', 'líneas'], ['pie', 'pizza', 'circular'],
  ['Elements', 'Elementos', 'Elementos'], ['Chart title', 'Título do gráfico', 'Título del gráfico'], ['Category axis title', 'Título do eixo de categorias', 'Título del eje de categorías'], ['Value axis title', 'Título do eixo de valores', 'Título del eje de valores'],
  ['Legend', 'Legenda', 'Leyenda'], ['Data labels', 'Rótulos de dados', 'Etiquetas de datos'], ['Switch row/column', 'Alternar linha/coluna', 'Cambiar fila/columna'],
  ['Insert chart', 'Inserir gráfico', 'Insertar gráfico'], ['Apply to chart', 'Aplicar ao gráfico', 'Aplicar al gráfico'], ['Remove chart', 'Remover gráfico', 'Quitar gráfico'], ['Your chart will appear here.', 'Seu gráfico vai aparecer aqui.', 'Tu gráfico aparecerá aquí.'],
  ['Select the data first.', 'Selecione os dados primeiro.', 'Selecciona primero los datos.'], ['Insert a chart first.', 'Insira um gráfico primeiro.', 'Inserta primero un gráfico.'], ['Chart updated.', 'Gráfico atualizado.', 'Gráfico actualizado.'], ['Chart removed.', 'Gráfico removido.', 'Gráfico quitado.'],
  ['No chart yet. Select the data, choose a chart type and press <b>Insert chart</b>.', 'Ainda não há gráfico. Selecione os dados, escolha um tipo de gráfico e pressione <b>Inserir gráfico</b>.', 'Todavía no hay gráfico. Selecciona los datos, elige un tipo de gráfico y pulsa <b>Insertar gráfico</b>.'],
  ['Turn the <b>legend</b> on.', 'Ative a <b>legenda</b>.', 'Activa la <b>leyenda</b>.'], ['Turn the <b>legend</b> off.', 'Desative a <b>legenda</b>.', 'Desactiva la <b>leyenda</b>.'], ['Turn on <b>data labels</b>.', 'Ative os <b>rótulos de dados</b>.', 'Activa las <b>etiquetas de datos</b>.'],
  ['Switch rows and columns: each <b>row</b> should become a series.', 'Alterne linhas e colunas: cada <b>linha</b> deve virar uma série.', 'Cambia filas y columnas: cada <b>fila</b> debe convertirse en una serie.'],
  ["Look the month up with a formula, don't type it.", 'Procure o mês com uma fórmula, não o digite.', 'Busca el mes con una fórmula, no lo escribas.'],
  ['Expand all', 'Expandir tudo', 'Expandir todo'], ['Collapse all', 'Recolher tudo', 'Contraer todo'], ['real-world test', 'teste do mundo real', 'prueba del mundo real'], ['real-world tests', 'testes do mundo real', 'pruebas del mundo real'],
  ['using', 'usando', 'usando'], ['then', 'depois', 'luego'], ['Select <b>{a}</b> and apply the <b>{f}</b> format.', 'Selecione <b>{a}</b> e aplique o formato <b>{f}</b>.', 'Selecciona <b>{a}</b> y aplica el formato <b>{f}</b>.'],
  ['…then fill down for the remaining {n} cells', '…depois preencha para baixo as {n} células restantes', '…luego rellena hacia abajo las {n} celdas restantes'], ['…then fill down for the remaining {n} cell', '…depois preencha para baixo a {n} célula restante', '…luego rellena hacia abajo la {n} celda restante'],
]);

/* patterns (regex on the full text node / message) */
const L_ = '(Beginner|Intermediate|Advanced)';
pat(/^(\d+) XP to next level$/, '$1 XP para o próximo nível', '$1 XP para el siguiente nivel');
pat(/^Lv (\d+) · (.+)$/, 'Nv $1 · $t2', 'Nv $1 · $t2');
pat(/^Lv (\d+)$/, 'Nv $1', 'Nv $1');
pat(/^Step (\d+) of (\d+)$/, 'Etapa $1 de $2', 'Paso $1 de $2');
pat(/^· Step (\d+) of (\d+)$/, '· Etapa $1 de $2', '· Paso $1 de $2');
pat(/^Step (\d+): (.+)$/, 'Etapa $1: $2', 'Paso $1: $2');
pat(/^\+(\d+) XP$/, '+$1 XP', '+$1 XP');
pat(/^(.+) · Lesson (\d+)$/, '$t1 · Aula $2', '$t1 · Lección $2');
pat(new RegExp('^' + L_ + ' · (.+)$'), '$t1 · $t2', '$t1 · $t2');
pat(/^(\d+) lessons?$/, '$1 aulas', '$1 lecciones');
pat(/^(\d+)\/(\d+) steps solved$/, '$1/$2 etapas resolvidas', '$1/$2 pasos resueltos');
pat(/^1 real-world test$/, '1 teste do mundo real', '1 prueba del mundo real');
pat(/^(\d+) real-world tests$/, '$1 testes do mundo real', '$1 pruebas del mundo real');
pat(/^(\d+) steps?$/, '$1 etapas', '$1 pasos');
pat(/^~(\d+) min$/, '~$1 min', '~$1 min');
pat(/^(\d+) min to go$/, 'faltam $1 min', 'faltan $1 min');
pat(/^(\d+) lessons in (\d+) learning paths\. Everything is unlocked, so start anywhere\.$/, '$1 aulas em $2 trilhas de aprendizado. Tudo está liberado, então comece por onde quiser.', '$1 lecciones en $2 rutas de aprendizaje. Todo está desbloqueado, así que empieza donde quieras.');
pat(/^Next: (.+)$/, 'Próxima: $1', 'Siguiente: $1');
pat(/^(.+) · (\d+)\/(\d+) steps solved$/, '$1 · $2/$3 etapas resolvidas', '$1 · $2/$3 pasos resueltos');
pat(/^Includes a \+(\d+) XP lesson bonus\.$/, 'Inclui um bônus de +$1 XP pela aula.', 'Incluye un bono de +$1 XP por la lección.');
pat(/^Includes a \+(\d+) XP test bonus\.$/, 'Inclui um bônus de +$1 XP pelo teste.', 'Incluye un bono de +$1 XP por la prueba.');
pat(/^…then fill down for the remaining (\d+) cells?$/, '…depois preencha para baixo as $1 células restantes', '…luego rellena hacia abajo las $1 celdas restantes');
pat(/^Target selection:$/, 'Seleção alvo:', 'Selección objetivo:');
pat(/^Chart inserted from (.+)\.$/, 'Gráfico inserido a partir de $1.', 'Gráfico insertado desde $1.');
pat(/^The chart uses <b>(.*?)<\/b>, but this task needs <b>(.*?)<\/b>\.$/, 'O gráfico usa <b>$1</b>, mas esta tarefa precisa de <b>$2</b>.', 'El gráfico usa <b>$1</b>, pero esta tarea necesita <b>$2</b>.');
pat(/^That is a <b>(.*?)<\/b> chart\. This task needs a <b>(.*?)<\/b> chart\.$/, 'Esse é um gráfico de <b>$t1</b>. Esta tarefa precisa de um gráfico de <b>$t2</b>.', 'Ese es un gráfico de <b>$t1</b>. Esta tarea necesita un gráfico de <b>$t2</b>.');
pat(/^Give the chart the title <b>(.*?)<\/b>\.$/, 'Dê ao gráfico o título <b>$1</b>.', 'Ponle al gráfico el título <b>$1</b>.');
pat(/^Set the category axis title to <b>(.*?)<\/b>\.$/, 'Defina o título do eixo de categorias como <b>$1</b>.', 'Define el título del eje de categorías como <b>$1</b>.');
pat(/^Set the value axis title to <b>(.*?)<\/b>\.$/, 'Defina o título do eixo de valores como <b>$1</b>.', 'Define el título del eje de valores como <b>$1</b>.');
pat(/^Rule applied to (.+)\.$/, 'Regra aplicada a $1.', 'Regla aplicada a $1.');
pat(/^PivotTable created at (.+)\.$/, 'Tabela Dinâmica criada em $1.', 'Tabla Dinámica creada en $1.');
pat(/^(\d+) of (\d+) records found\.$/, '$1 de $2 registros encontrados.', '$1 de $2 registros encontrados.');
pat(/^(\d+) records\.$/, '$1 registros.', '$1 registros.');
pat(/^Data › Filter \/ Sort · table (.+)$/, 'Dados › Filtrar / Classificar · tabela $1', 'Datos › Filtrar / Ordenar · tabla $1');
pat(/^Insert › PivotTable · source (.+) · output at (.+)$/, 'Inserir › Tabela Dinâmica · origem $1 · saída em $2', 'Insertar › Tabla Dinámica · origen $1 · salida en $2');

/* messages from the exercise checkers (HTML strings) */
const NUM = '(?:,)?';
pat(/^(\S+) is still empty\.$/, '$1 ainda está vazia.', '$1 todavía está vacía.');
pat(/^(\S+) returned <b>(.*?)<\/b>\. Check the spelling, commas and brackets in your formula\.$/, '$1 retornou <b>$2</b>. Verifique a grafia, as vírgulas e os parênteses da sua fórmula.', '$1 devolvió <b>$2</b>. Revisa la ortografía, las comas y los paréntesis de tu fórmula.');
pat(/^(\S+) shows <b>(.*?)<\/b>,? but (?:it )?should show <b>(.*?)<\/b>\.$/, '$1 mostra <b>$2</b>, mas deveria mostrar <b>$3</b>.', '$1 muestra <b>$2</b>, pero debería mostrar <b>$3</b>.');
pat(/^You've selected <b>(.*?)<\/b>; the target is <b>(.*?)<\/b>\.$/, 'Você selecionou <b>$1</b>; o alvo é <b>$2</b>.', 'Has seleccionado <b>$1</b>; el objetivo es <b>$2</b>.');
pat(/^Not yet: every highlighted cell needs the <b>(.*?)<\/b> format\.$/, 'Ainda não: todas as células destacadas precisam do formato <b>$t1</b>.', 'Todavía no: todas las celdas resaltadas necesitan el formato <b>$t1</b>.');
pat(/^Right result, but (\S+) should be a <b>formula<\/b> \(start with <b>=<\/b>\) so it updates by itself\.$/, 'Resultado certo, mas $1 deve ser uma <b>fórmula</b> (comece com <b>=</b>) para se atualizar sozinha.', 'Resultado correcto, pero $1 debe ser una <b>fórmula</b> (empieza con <b>=</b>) para actualizarse sola.');
pat(/^(\S+) should contain a <b>formula<\/b> starting with =\.$/, '$1 deve conter uma <b>fórmula</b> começando com =.', '$1 debe contener una <b>fórmula</b> que empiece con =.');
pat(/^Not bold yet\. Select (.+) and press Ctrl \+ B\.$/, 'Ainda não está em negrito. Selecione $1 e pressione Ctrl + B.', 'Aún no está en negrita. Selecciona $1 y pulsa Ctrl + B.');
pat(/^(\S+) should be empty after a cut\.$/, '$1 deve ficar vazia depois de recortar.', '$1 debe quedar vacía después de cortar.');
pat(/^(\S+) should hold a formula\.$/, '$1 deve conter uma fórmula.', '$1 debe contener una fórmula.');
pat(/^(\S+) returned <b>(.*?)<\/b>\. Check the formula( and make sure the cells below are empty)?\.$/, '$1 retornou <b>$2</b>. Verifique a fórmula$3.', '$1 devolvió <b>$2</b>. Revisa la fórmula$3.');
pat(/^You're on <b>(.*?)<\/b>\. Open <b>Feb<\/b>\.$/, 'Você está em <b>$1</b>. Abra <b>Feb</b>.', 'Estás en <b>$1</b>. Abre <b>Feb</b>.');
pat(/^No matching rule on <b>(.*?)<\/b> yet\. Select the range, then use Home › Conditional Formatting › Highlight Cells Rules\.$/, 'Ainda não há uma regra correspondente em <b>$1</b>. Selecione o intervalo e use Página Inicial › Formatação Condicional › Regras de Realce das Células.', 'Todavía no hay una regla que coincida en <b>$1</b>. Selecciona el rango y usa Inicio › Formato condicional › Reglas para resaltar celdas.');
pat(/^No matching rule on <b>(.*?)<\/b> yet\.$/, 'Ainda não há uma regra correspondente em <b>$1</b>.', 'Todavía no hay una regla que coincida en <b>$1</b>.');
pat(/^(\S+) shows <b>(.*?)<\/b> but should show <b>(.*?)<\/b>\.$/, '$1 mostra <b>$2</b> mas deveria mostrar <b>$3</b>.', '$1 muestra <b>$2</b> pero debería mostrar <b>$3</b>.');
// the dashboard "Welcome" line has two shapes; handled with functions below
P.pt = P.pt.filter((p) => p[1] !== null); P.es = P.es.filter((p) => p[1] !== null);
pat(/^([\d.,]+) XP · Level (\d+) (.+?) · (\d+) XP to level (\d+)$/, '$1 XP · Nível $2 $t3 · $4 XP para o nível $5', '$1 XP · Nivel $2 $t3 · $4 XP para el nivel $5');
pat(/^([\d.,]+) XP · Level (\d+) (.+)$/, '$1 XP · Nível $2 $t3', '$1 XP · Nivel $2 $t3');

const I18N = { LANGS, get lang() { return lang; }, set, t, tx, msg, addLessons, applyLessons, coverage, walk, observe, ui, pat, gloss, word, formula: localFormula, glossWords: () => [...GL.pt.keys()] };
root.I18N = I18N;
document.documentElement.lang = HTML_LANG[lang];
if (document.body) observe(); else document.addEventListener('DOMContentLoaded', observe);
})(window);
