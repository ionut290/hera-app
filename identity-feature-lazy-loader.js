(() => {
  "use strict";

  const FEATURE_SCRIPT = "./identity-card-feature.js?v=20260815-lazy1";
  const FEATURE_STYLE = "./identity-card-feature.css?v=20260727a";
  const TRIGGER_SELECTOR = "#identity-card-btn, #fuel-pin-btn";
  let loadPromise = null;

  const ensureStyle = () => {
    const existing = document.querySelector('link[data-hera-identity-style="true"]');
    if (existing) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const link = document.createElement("link");
      link.rel = "stylesheet";
      link.href = FEATURE_STYLE;
      link.dataset.heraIdentityStyle = "true";
      link.onload = () => resolve();
      link.onerror = () => reject(new Error("Stile Tesserino/PIN non disponibile"));
      document.head.appendChild(link);
    });
  };

  const waitForAuthBinding = async (trigger) => {
    const auth = window.firebase?.auth?.();
    if (!auth?.currentUser) return;
    const deadline = Date.now() + 3000;
    while (trigger.disabled && Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 25));
    }
  };

  const ensureFeature = () => {
    if (document.querySelector('script[data-hera-identity-feature="true"]')) return loadPromise || Promise.resolve();
    if (loadPromise) return loadPromise;
    loadPromise = Promise.all([
      ensureStyle(),
      new Promise((resolve, reject) => {
        const script = document.createElement("script");
        script.src = FEATURE_SCRIPT;
        script.dataset.heraIdentityFeature = "true";
        script.onload = () => resolve();
        script.onerror = () => reject(new Error("Modulo Tesserino/PIN non disponibile"));
        document.head.appendChild(script);
      })
    ]).catch((error) => {
      loadPromise = null;
      throw error;
    });
    return loadPromise;
  };

  document.addEventListener("click", async (event) => {
    const trigger = event.target.closest?.(TRIGGER_SELECTOR);
    if (!trigger || trigger.dataset.heraIdentityReplay === "true") return;
    event.preventDefault();
    event.stopImmediatePropagation();
    trigger.disabled = true;
    try {
      await ensureFeature();
      await waitForAuthBinding(trigger);
      trigger.dataset.heraIdentityReplay = "true";
      trigger.disabled = false;
      trigger.click();
    } catch (error) {
      console.error("Caricamento Tesserino/PIN non riuscito:", error);
      trigger.disabled = false;
      window.alert("Funzione Tesserino/PIN non disponibile. Controlla la connessione e riprova.");
    } finally {
      delete trigger.dataset.heraIdentityReplay;
    }
  }, true);
})();
