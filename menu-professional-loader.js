(() => {
  const loadStylesheet = () => {
    if (document.querySelector('link[data-professional-menu="true"]')) return;
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "professional-menu.css?v=20260731a";
    link.dataset.professionalMenu = "true";
    document.head.appendChild(link);
  };

  const loadMenuScript = () => {
    if (document.querySelector('script[data-professional-menu="true"]')) return;
    const script = document.createElement("script");
    script.src = "professional-menu.js?v=20260731a";
    script.defer = true;
    script.dataset.professionalMenu = "true";
    document.head.appendChild(script);
  };

  loadStylesheet();
  loadMenuScript();
})();
