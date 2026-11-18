/* Spreadsheet Academy entitlement: the first two Foundations lessons are free, the rest needs the "fullcourse"
   durable add-on from the Microsoft Store. window.Entitlement

   In the Tauri / MSIX build the licence comes from two Rust commands (see README, "Store licence commands"):
     store_is_unlocked({ productId, storeId }) -> bool      (is the add-on owned?)
     store_buy_unlock({ productId, storeId })  -> bool      (shows the Store purchase dialog; true if owned afterwards)
   Outside Tauri (browser, development) nothing can be bought. Open the page with ?unlock=1 to unlock everything
   for testing (remembered on this device) and ?unlock=0 to switch it off again. */
(function (root) {
'use strict';
const CFG = { productId: 'fullcourse', storeId: '9N4ZJVS7LM8Z', freeLessons: 2 };
const CACHE = 'sa.entitled.v1', DEV = 'sa.dev.unlock';
/* the free browser demo (learnexcel.online/app): only the free lessons are shipped, nothing can be unlocked, the purchase button opens the Store */
const WEB = !!root.__WEB__;
const get = (k) => { try { return localStorage.getItem(k); } catch (e) { return null; } };
const put = (k, v) => { try { if (v == null) localStorage.removeItem(k); else localStorage.setItem(k, v); } catch (e) { /* storage unavailable */ } };

if (!WEB) try { const q = new URLSearchParams(location.search).get('unlock'); if (q === '1') put(DEV, '1'); else if (q === '0') put(DEV, null); } catch (e) { /* ignore */ }

const invoke = () => {
  const T = root.__TAURI__, I = root.__TAURI_INTERNALS__;
  return (T && T.core && T.core.invoke) || (T && T.tauri && T.tauri.invoke) || (I && I.invoke) || null;
};
const free = new Set(((root.COURSE && root.COURSE.TRACKS[0] && root.COURSE.TRACKS[0].lessons) || []).slice(0, CFG.freeLessons).map((l) => l.id));
let owned = get(CACHE) === '1'; const listeners = [];
const full = () => !WEB && (owned || get(DEV) === '1');
const setOwned = (v) => { const was = full(); owned = !!v; put(CACHE, owned ? '1' : null); if (full() !== was) listeners.forEach((f) => f()); return owned; };

const E = {
  CFG,
  get full() { return full(); },
  web: WEB,
  inStore: () => !!invoke(),
  /* Windows: opens the Store listing of the app; other systems: '' (the Store is a Windows app) */
  storeHref: () => (/Windows NT 1[01]/.test(navigator.userAgent) ? 'ms-windows-store://pdp/?PFN=HerenLabs.SpreadsheetAcademyLearnXLSXSkills_4rr9weehh6kfm' : ''),
  isFree: (id) => free.has(id),
  has: (id) => full() || free.has(id),
  onChange: (f) => { listeners.push(f); },
  /* ask the Store whether the add-on is owned (keeps the last known answer when the Store can't be reached) */
  async refresh() {
    const call = invoke(); if (!call) return full();
    try { setOwned(await call('store_is_unlocked', { productId: CFG.productId, storeId: CFG.storeId })); } catch (e) { /* offline or Store unavailable: keep the cached answer */ }
    return full();
  },
  /* returns 'owned' | 'cancelled' | 'unavailable' */
  async buy() {
    const call = invoke(); if (!call) return 'unavailable';
    try { return setOwned(await call('store_buy_unlock', { productId: CFG.productId, storeId: CFG.storeId })) ? 'owned' : 'cancelled'; } catch (e) { return String(e).includes('not-packaged') ? 'unavailable' : 'cancelled'; }
  },
  /* returns 'owned' | 'none' | 'unavailable' */
  async restore() {
    if (!invoke()) return 'unavailable';
    await E.refresh(); return full() ? 'owned' : 'none';
  },
};
root.Entitlement = E;

if (root.I18N) {
  root.I18N.ui([
    ['Unlock full course', 'Desbloquear curso completo', 'Desbloquear curso completo'], ['Unlock the full course', 'Desbloqueie o curso completo', 'Desbloquea el curso completo'],
    ['Unlock', 'Desbloquear', 'Desbloquear'], ['Locked lesson', 'Aula bloqueada', 'Lección bloqueada'], ['Free', 'Grátis', 'Gratis'],
    ['The first two lessons are free. Unlock the rest, including every path and real-world test, with a single purchase.', 'As duas primeiras aulas são gratuitas. Desbloqueie o restante, com todas as trilhas e testes reais, em uma única compra.', 'Las dos primeras lecciones son gratuitas. Desbloquea el resto, con todas las rutas y pruebas reales, en una sola compra.'],
    ['hands-on lessons in all', 'aulas práticas no total', 'lecciones prácticas en total'], ['Real-world tests and a capstone project', 'Testes reais e um projeto final', 'Pruebas reales y un proyecto final'],
    ['One-time purchase, yours to keep', 'Compra única, sua para sempre', 'Compra única, tuya para siempre'], ['Works offline once unlocked', 'Funciona offline depois de desbloqueado', 'Funciona sin conexión una vez desbloqueado'],
    ['Restore purchase', 'Restaurar compra', 'Restaurar compra'], ['Get the full course on the Microsoft Store', 'Baixar o curso completo na Microsoft Store', 'Obtener el curso completo en Microsoft Store'],
    ['The full course is a Windows app. Open this page on a Windows PC to get it.', 'O curso completo é um app para Windows. Abra esta página em um PC com Windows para baixá-lo.', 'El curso completo es una aplicación para Windows. Abre esta página en un PC con Windows para obtenerlo.'], ['Back to courses', 'Voltar aos cursos', 'Volver a los cursos'],
    ['Full course unlocked. Enjoy!', 'Curso completo desbloqueado. Bom estudo!', '¡Curso completo desbloqueado! ¡A disfrutar!'],
    ['No purchase found for this account.', 'Nenhuma compra encontrada para esta conta.', 'No se encontró ninguna compra en esta cuenta.'],
    ['Purchases are available in the Microsoft Store version of the app.', 'As compras estão disponíveis na versão do app na Microsoft Store.', 'Las compras están disponibles en la versión de la aplicación en Microsoft Store.'],
    ['The purchase was not completed.', 'A compra não foi concluída.', 'La compra no se completó.'],
    ['Free lessons complete', 'Aulas gratuitas concluídas', 'Lecciones gratuitas completadas'],
    ['You have finished the free lessons. Unlock the full course to keep going.', 'Você concluiu as aulas gratuitas. Desbloqueie o curso completo para continuar.', 'Has terminado las lecciones gratuitas. Desbloquea el curso completo para continuar.'],
    ['Unlock to continue', 'Desbloqueie para continuar', 'Desbloquea para continuar'],
    ['Pick a path. The first two lessons are free.', 'Escolha uma trilha. As duas primeiras aulas são gratuitas.', 'Elige una ruta. Las dos primeras lecciones son gratuitas.'],
  ]);
  root.I18N.pat(/^(\d+) lessons in (\d+) learning paths\. The first two lessons are free\.$/, '$1 aulas em $2 trilhas de aprendizado. As duas primeiras aulas são gratuitas.', '$1 lecciones en $2 rutas de aprendizaje. Las dos primeras lecciones son gratuitas.');
}
})(window);
