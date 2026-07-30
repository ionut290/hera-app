(function () {
  "use strict";
  const TECHNICAL_COLUMNS = ["RECORD_ID", "UPDATED_AT", "UPDATED_BY", "SYNC_VERSION", "SYNC_SOURCE", "ROW_STATUS"];
  const SHEETS = ["PERSONALE", "MEZZI", "COMMESSE_PERSONALE", "COMMESSE_MEZZI", "LOG_SINCRONIZZAZIONE"];
  const configs = new Map();
  const clean = (value) => String(value == null ? "" : value).trim();
  const configRef = (type) => db.collection("registryGoogleSheetLinks").doc(type);
  const records = (type) => type === "personale" ? (window.personaleRecords || personaleRecords || []) : (window.mezziRecords || mezziRecords || []);
  const timestamp = (value) => value?.toDate instanceof Function ? value.toDate().toISOString() : clean(value) || new Date(0).toISOString();

  async function load(type) {
    if (configs.has(type)) return configs.get(type);
    const snap = await configRef(type).get();
    const config = snap.exists ? snap.data() : {};
    configs.set(type, config);
    return config;
  }
  async function token() {
    const user = auth.currentUser;
    if (!user) throw new Error("Sessione scaduta.");
    return user.getIdToken();
  }
  async function call(payload) {
    const response = await fetch("/.netlify/functions/google-sheet-sync", { method: "POST", headers: { "Content-Type": "application/json", Authorization: `Bearer ${await token()}` }, body: JSON.stringify(payload) });
    const result = await response.json().catch(() => ({}));
    if (!response.ok || !result.ok) throw new Error(result.error || "Sincronizzazione non riuscita.");
    return result;
  }
  function technical(record, source = "APP") {
    return { RECORD_ID: record.id, UPDATED_AT: timestamp(record.updatedAt || record.createdAt), UPDATED_BY: clean(record.updatedBy), SYNC_VERSION: Number(record.syncVersion) || 1, SYNC_SOURCE: source, ROW_STATUS: "ACTIVE" };
  }
  function makeSheets() {
    const api = window.HeraManagementV2;
    const personnel = api.workbookRows("personale", personaleRecords).map((row, index) => ({ ...row, ...technical(personaleRecords[index]) }));
    const vehicles = api.workbookRows("mezzi", mezziRecords).map((row, index) => ({ ...row, ...technical(mezziRecords[index]) }));
    const personnelLinks = personaleRecords.flatMap((person) => (HeraManagementCore.legacyEnabledIds(person, (value) => value) || []).map((commessaId) => ({ RECORD_ID: `${person.id}:${commessaId}`, ID_OPERATORE: person.id, ID_COMMESSA: commessaId, ROW_STATUS: "ACTIVE", UPDATED_AT: timestamp(person.updatedAt), UPDATED_BY: clean(person.updatedBy), SYNC_VERSION: Number(person.syncVersion) || 1, SYNC_SOURCE: "APP" })));
    const vehicleLinks = mezziRecords.flatMap((vehicle) => (Array.isArray(vehicle.commessaIds) ? vehicle.commessaIds : []).map((commessaId) => ({ RECORD_ID: `${vehicle.id}:${commessaId}`, ID_MEZZO: vehicle.id, ID_COMMESSA: commessaId, ROW_STATUS: "ACTIVE", UPDATED_AT: timestamp(vehicle.updatedAt), UPDATED_BY: clean(vehicle.updatedBy), SYNC_VERSION: Number(vehicle.syncVersion) || 1, SYNC_SOURCE: "APP" })));
    return { PERSONALE: personnel, MEZZI: vehicles, COMMESSE_PERSONALE: personnelLinks, COMMESSE_MEZZI: vehicleLinks, LOG_SINCRONIZZAZIONE: [] };
  }
  async function link(type) {
    const existing = await load(type);
    const supplied = window.prompt("Incolla il link di un Google Sheet esistente oppure lascia vuoto per crearne uno nuovo.", existing.sheetUrl || "");
    if (supplied === null) return;
    const result = await call({ action: "createRegistrySpreadsheet", registry: type, sheetUrl: clean(supplied), sheetNames: SHEETS });
    const patch = { sheetUrl: result.sheetUrl, spreadsheetId: result.spreadsheetId, linkedAt: firebase.firestore.FieldValue.serverTimestamp(), linkedBy: currentUser?.uid || "", autoSync: false, conflictPolicy: "LATEST_WINS", noAutomaticDeletion: true };
    await configRef(type).set(patch, { merge: true }); configs.set(type, { ...existing, ...patch, sheetUrl: result.sheetUrl, spreadsheetId: result.spreadsheetId });
    alert("Google Sheet collegato. Puoi avviare la prima sincronizzazione.");
  }
  async function sync(type) {
    const config = await load(type);
    if (!config.sheetUrl) return link(type);
    const result = await call({ action: "syncRegistrySpreadsheet", registry: type, sheetUrl: config.sheetUrl, spreadsheetId: config.spreadsheetId, sheets: makeSheets(), conflictPolicy: config.conflictPolicy || "LATEST_WINS", noAutomaticDeletion: true });
    await configRef(type).set({ lastSyncAt: firebase.firestore.FieldValue.serverTimestamp(), lastSyncBy: currentUser?.uid || "", lastConflictCount: result.conflicts?.length || 0 }, { merge: true });
    const primary = type === "personale" ? "PERSONALE" : "MEZZI", idColumn = type === "personale" ? "ID_OPERATORE" : "ID_MEZZO";
    const incoming = (result.incoming?.[primary] || []).filter((row) => clean(row.ROW_STATUS).toUpperCase() !== "DELETED").map((row) => ({ ...row, [idColumn]: row[idColumn] || row.RECORD_ID }));
    if (incoming.length) {
      const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(incoming), primary);
      const bytes = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      await window.HeraManagementV2.previewImport(type, { name: `Google_Sheets_${primary}.xlsx`, arrayBuffer: async () => bytes });
    } else alert(`Sincronizzazione completata. ${result.rowsWritten || 0} righe aggiornate; ${result.conflicts?.length || 0} conflitti; nessuna eliminazione automatica.`);
  }
  async function settings(type) {
    const config = await load(type); if (!config.sheetUrl) return link(type);
    const policy = window.prompt("Gestione conflitti: LATEST_WINS oppure APP_WINS.", config.conflictPolicy || "LATEST_WINS");
    if (policy === null) return;
    const normalized = clean(policy).toUpperCase(); if (!["LATEST_WINS", "APP_WINS"].includes(normalized)) return alert("Impostazione non valida.");
    const autoSync = window.confirm("Abilitare la sincronizzazione automatica incrementale all’apertura della sezione?");
    await configRef(type).set({ conflictPolicy: normalized, autoSync, noAutomaticDeletion: true, updatedAt: firebase.firestore.FieldValue.serverTimestamp() }, { merge: true });
    configs.set(type, { ...config, conflictPolicy: normalized, autoSync }); alert("Impostazioni sincronizzazione salvate.");
  }
  async function unlink(type) {
    const config = await load(type); if (!config.sheetUrl) return;
    if (!window.confirm("Scollegare il Google Sheet? Il foglio e i dati non saranno eliminati.")) return;
    await configRef(type).delete(); configs.delete(type); alert("Google Sheet scollegato senza eliminare dati.");
  }
  async function open(type) { const config = await load(type); if (!config.sheetUrl) return link(type); window.open(config.sheetUrl, "_blank", "noopener,noreferrer"); }
  function bind(root, type) {
    root.querySelector("[data-sheet-link]")?.addEventListener("click", () => void link(type).catch((error) => alert(error.message)));
    root.querySelector("[data-sheet-sync]")?.addEventListener("click", () => void sync(type).catch((error) => alert(error.message)));
    root.querySelector("[data-sheet-open]")?.addEventListener("click", () => void open(type).catch((error) => alert(error.message)));
    root.querySelector("[data-sheet-settings]")?.addEventListener("click", () => void settings(type).catch((error) => alert(error.message)));
    root.querySelector("[data-sheet-unlink]")?.addEventListener("click", () => void unlink(type).catch((error) => alert(error.message)));
    void load(type).then((config) => { if (config.autoSync && config.sheetUrl && navigator.onLine) sync(type).catch(() => {}); });
  }
  window.RegistryGoogleSheetSync = { bind, sync, SHEETS, TECHNICAL_COLUMNS };
})();
