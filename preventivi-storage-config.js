(() => {
  'use strict';

  const PV = window.HeraPreventivi;
  if (!PV) throw new Error('Preventivi core non caricato.');

  // Usa un percorso già protetto dalle regole della commessa. In questo modo
  // soltanto gli utenti autenticati e autorizzati possono leggere e scrivere,
  // senza allargare le regole Firestore ad altre collezioni.
  PV.collections.priceLists = 'commesse/__preventivi__/prezziari';
  PV.collections.quotes = 'commesse/__preventivi__/preventivi';
})();
