(function installHeraOracleReadAdapter(global) {
  "use strict";

  if (global.HeraOracleReadAdapter?.installed) return;

  const config = Object.freeze({
    enabled: global.__HERA_ORACLE_PILOT__?.enabled === true,
    baseUrl: String(global.__HERA_ORACLE_PILOT__?.baseUrl || "/oracle-api").replace(/\/$/, ""),
    allowedEmail: String(global.__HERA_ORACLE_PILOT__?.allowedEmail || "").trim().toLowerCase()
  });

  async function getVerifiedToken() {
    if (!config.enabled) throw new Error("oracle_pilot_disabled");
    const user = global.firebase?.auth?.().currentUser;
    const email = String(user?.email || "").trim().toLowerCase();
    if (!user || !email) throw new Error("authentication_required");
    if (config.allowedEmail && email !== config.allowedEmail) throw new Error("pilot_access_denied");
    return user.getIdToken(false);
  }

  async function request(path) {
    const token = await getVerifiedToken();
    const response = await fetch(`${config.baseUrl}${path}`, {
      method: "GET",
      cache: "no-store",
      credentials: "same-origin",
      headers: { Authorization: `Bearer ${token}` }
    });
    const body = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(String(body.error || `oracle_http_${response.status}`));
    return body;
  }

  function normalizeCommessa(item) {
    return Object.freeze({
      id: String(item.document_id || ""),
      nome: String(item.denominazione || item.codice || item.document_id || "Commessa"),
      codice: String(item.codice || ""),
      stato: String(item.stato || ""),
      oraclePilotReadOnly: true
    });
  }

  function normalizeImpianto(item) {
    return Object.freeze({
      id: String(item.document_id || ""),
      denominazione: String(item.denominazione || ""),
      idSap: String(item.codice || ""),
      comune: String(item.comune || ""),
      indirizzo: String(item.indirizzo || ""),
      distretto: String(item.distretto || ""),
      tipologiaIntervento: String(item.tipologia_intervento || ""),
      oraclePilotReadOnly: true
    });
  }

  async function loadCatalog() {
    const [commesse, impianti] = await Promise.all([
      request("/v1/catalogo?collection=commesse&limit=100"),
      request("/v1/catalogo?collection=impianti&limit=500")
    ]);
    return Object.freeze({
      commesse: Object.freeze((commesse.items || []).map(normalizeCommessa)),
      impianti: Object.freeze((impianti.items || []).map(normalizeImpianto))
    });
  }

  async function loadMap(collection = "impianti") {
    if (!new Set(["commesse", "impianti"]).has(collection)) throw new Error("invalid_collection");
    const result = await request(`/v1/mappa?collection=${encodeURIComponent(collection)}&limit=500`);
    return Object.freeze((result.items || []).map((item) => Object.freeze({
      id: String(item.document_id || ""),
      collectionName: String(item.collection_name || ""),
      latitudine: Number(item.latitudine),
      longitudine: Number(item.longitudine),
      oraclePilotReadOnly: true
    })));
  }

  global.HeraOracleReadAdapter = Object.freeze({
    installed: true,
    enabled: config.enabled,
    mode: "read-only",
    loadCatalog,
    loadMap
  });
})(window);
