/* Store button of the blog pages. Keep PFN and STORE_ID in sync with the landing pages (index.html and pt/index.html). */
(function () {
  var PFN = 'HerenLabs.SpreadsheetAcademyLearnXLSXSkills_4rr9weehh6kfm';
  var STORE_ID = '__STORE_ID__';
  var pt = document.documentElement.lang.toLowerCase().indexOf('pt') === 0;
  var hasId = /^9[A-Z0-9]{11}$/i.test(STORE_ID), isWin = /Windows NT 1[01]/.test(navigator.userAgent);
  var href = isWin ? 'ms-windows-store://pdp/?PFN=' + PFN : hasId ? 'https://apps.microsoft.com/detail/' + STORE_ID + '?hl=' + (pt ? 'pt-br&gl=BR' : 'en-us&gl=US') + '&cid=website' : '';
  document.querySelectorAll('[data-store]').forEach(function (a) {
    if (href) a.href = href; else { a.href = (pt ? '/pt/' : '/'); }
    a.addEventListener('click', function () { if (window.gtag) gtag('event', 'store_click', { transport_type: 'beacon', from: 'blog' }); });
  });
})();
