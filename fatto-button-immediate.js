(() => {
  "use strict";

  // Lo stato del pulsante FATTO è gestito dal flusso transazionale in app.js.
  // Non intercettare pointer/click qui: una modifica grafica anticipata potrebbe
  // mostrare un completamento non ancora salvato o generare un doppio invio.
})();

(() => {
  "use strict";
  if (document.querySelector('script[data-password-access-manager="true"]')) return;
  const script = document.createElement("script");
  script.src = "password-access-manager.js?v=20260727a";
  script.defer = true;
  script.dataset.passwordAccessManager = "true";
  document.head.appendChild(script);
})();
