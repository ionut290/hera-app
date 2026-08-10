(function () {
  "use strict";

  const STYLE_ID = "control-center-accordion-style";
  const SCRIPT_ID = "control-center-accordion-script";
  const REGISTRY_CACHE_SCRIPT_ID = "registry-device-cache-script";

  function addStyle() {
    if (document.getElementById(STYLE_ID)) return;
    const link = document.createElement("link");
    link.id = STYLE_ID;
    link.rel = "stylesheet";
    link.href = "./control-center-accordion.css?v=20260810-pinned2";
    document.head.appendChild(link);
  }

  function addScript() {
    if (document.getElementById(SCRIPT_ID)) return;
    const script = document.createElement("script");
    script.id = SCRIPT_ID;
    script.src = "./control-center-accordion.js?v=20260810-pinned2";
    script.defer = true;
    document.head.appendChild(script);
  }

  function addRegistryDeviceCache() {
    if (document.getElementById(REGISTRY_CACHE_SCRIPT_ID)) return;
    const script = document.createElement("script");
    script.id = REGISTRY_CACHE_SCRIPT_ID;
    script.src = "./registry-device-cache.js?v=20260804a";
    script.defer = true;
    document.head.appendChild(script);
  }

  addStyle();
  addScript();
  addRegistryDeviceCache();
})();
