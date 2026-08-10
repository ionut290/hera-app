(() => {
  'use strict';

  const scripts = [
    ['./commessa-listener-cleanup.js?v=20260806a', 'commessaListenerCleanup'],
    ['./header-menu-runtime-original.js?v=20260804-diagnostics-reset1', 'headerMenuRuntimeOriginal'],
    ['./control-center-loader.js?v=20260810-pinned2', 'controlCenterLoader']
  ];

  scripts.forEach(([src, datasetKey]) => {
    if (document.querySelector(`script[data-${datasetKey.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`)}]`)) return;
    const script = document.createElement('script');
    script.src = src;
    script.async = false;
    script.dataset[datasetKey] = '1';
    script.addEventListener('error', () => console.warn(`Script non caricato: ${src}`), { once: true });
    document.head.appendChild(script);
  });
})();
