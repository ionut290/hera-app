(() => {
  const href = 'visual-refresh.css?v=20260726a';
  const alreadyLoaded = Array.from(document.styleSheets || []).some((sheet) => {
    try {
      return sheet.href && sheet.href.includes('visual-refresh.css');
    } catch (_) {
      return false;
    }
  });

  if (alreadyLoaded || document.querySelector('link[data-visual-refresh="true"]')) return;

  const link = document.createElement('link');
  link.rel = 'stylesheet';
  link.href = href;
  link.dataset.visualRefresh = 'true';
  document.head.appendChild(link);
})();
