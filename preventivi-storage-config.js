(() => {
  'use strict';

  const PV = window.HeraPreventivi;
  if (!PV) throw new Error('Preventivi core non caricato.');

  // Usa un percorso già protetto dalle regole della commessa. In questo modo
  // soltanto gli utenti autenticati e autorizzati possono leggere e scrivere,
  // senza allargare le regole Firestore ad altre collezioni.
  PV.collections.priceLists = 'commesse/__preventivi__/prezziari';
  PV.collections.quotes = 'commesse/__preventivi__/preventivi';

  // Carica in modo isolato la matrice Excel e l'aggiornamento completo dei
  // prezziari. Il modulo attende autonomamente che Prezziari e Preventivi siano
  // pronti, senza modificare le funzioni operative dell'app.
  if (!document.querySelector('link[data-preventivi-matrix-css]')) {
    const css = document.createElement('link');
    css.rel = 'stylesheet';
    css.href = './preventivi-price-list-matrix.css?v=20260731b';
    css.dataset.preventiviMatrixCss = '1';
    document.head.appendChild(css);
  }

  if (!document.querySelector('script[data-preventivi-matrix-js]')) {
    const script = document.createElement('script');
    script.src = './preventivi-price-list-matrix.js?v=20260731b';
    script.dataset.preventiviMatrixJs = '1';
    script.addEventListener('error', () => {
      console.warn('Preventivi: modulo matrice prezziario non caricato.');
    }, { once: true });
    document.head.appendChild(script);
  }
})();
