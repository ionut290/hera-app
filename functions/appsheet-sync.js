"use strict";

const admin = require("firebase-admin");
const { onDocumentWritten } = require("firebase-functions/v2/firestore");
const { onCall, HttpsError } = require("firebase-functions/v2/https");
const { defineSecret } = require("firebase-functions/params");
const {
  buildMappedRow,
  validateConfig,
  shouldSyncWrite,
  buildActionBody,
  buildActionUrl
} = require("./appsheet-sync-core");

if (!admin.apps.length) admin.initializeApp();

const REGION = "europe-west1";
const DEFAULT_APP_ID = "a33fc9cd-0a18-4aa8-b70a-c067c0c6c278";
const ADMIN_EMAIL = "ionut29019@gmail.com";
const APPSHEET_ACCESS_KEY = defineSecret("APPSHEET_ACCESS_KEY");
const CONFIG_DOC_PATH = "appConfig/appsheetSync";
const UPSERT_BATCH_SIZE = 40;

let cachedConfig = null;
let cachedConfigAt = 0;
const CONFIG_CACHE_MS = 60 * 1000;

function chunk(items, size) {
  const result = [];
  for (let index = 0; index < items.length; index += size) result.push(items.slice(index, index + size));
  return result;
}

async function loadConfig({ force = false } = {}) {
  const now = Date.now();
  if (!force && cachedConfig && now - cachedConfigAt < CONFIG_CACHE_MS) return cachedConfig;

  const snapshot = await admin.firestore().doc(CONFIG_DOC_PATH).get();
  const raw = snapshot.exists ? snapshot.data() || {} : {};
  cachedConfig = {
    ...raw,
    appId: String(raw.appId || DEFAULT_APP_ID).trim(),
    region: String(raw.region || "www.appsheet.com").trim()
  };
  cachedConfigAt = now;
  return cachedConfig;
}

function getAccessKey() {
  return String(process.env.APPSHEET_ACCESS_KEY || "").trim();
}

async function invokeAppSheetAction(config, tableConfig, action, rows) {
  if (!rows.length) return { Rows: [] };
  const accessKey = getAccessKey();
  if (!accessKey) throw new Error("APPSHEET_ACCESS_KEY non configurata");

  const response = await fetch(buildActionUrl(config, tableConfig.tableName), {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      ApplicationAccessKey: accessKey
    },
    body: JSON.stringify(buildActionBody(action, rows, {
      locale: config.locale || "it-IT",
      timezone: config.timezone || "Europe/Rome"
    }))
  });

  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : {};
  } catch (_) {
    payload = { raw: text };
  }

  if (!response.ok) {
    const error = new Error(`AppSheet ${action} fallita (${response.status})`);
    error.status = response.status;
    error.payload = payload;
    throw error;
  }
  return payload || {};
}

async function syncSingleChange({ entity, before, after, id, parentId = "", path }) {
  try {
    const config = await loadConfig();
    const validation = validateConfig(config);
    if (!validation.valid) return null;

    const tableConfig = config.tables?.[entity];
    const beforeExists = Boolean(before?.exists);
    const afterExists = Boolean(after?.exists);
    const action = shouldSyncWrite(beforeExists, afterExists, tableConfig);
    if (action === "Ignore") return null;

    // Per sicurezza le cancellazioni restano disattivate: nessun dato AppSheet viene eliminato automaticamente.
    if (action === "Delete") return null;

    const source = after?.data() || {};
    const row = buildMappedRow(source, tableConfig, {
      id,
      parentId,
      path,
      deleted: false
    });

    const keyValue = row[tableConfig.keyColumn];
    if (keyValue === undefined || keyValue === null || String(keyValue).trim() === "") {
      console.warn("Sync AppSheet saltata: chiave riga mancante.", { entity, id, path });
      return null;
    }

    await invokeAppSheetAction(config, tableConfig, action, [row]);
    console.info("Sync AppSheet completata.", { entity, id, action, table: tableConfig.tableName });
  } catch (error) {
    // IMPORTANTE: il salvataggio Varga è già avvenuto. Un errore AppSheet non deve mai bloccare o annullare Varga Cantieri.
    console.error("Sync AppSheet non riuscita; Varga Cantieri resta invariata.", {
      entity,
      id,
      path,
      message: error?.message || String(error),
      status: error?.status || null,
      payload: error?.payload || null
    });
  }
  return null;
}

async function findExistingRows(config, tableConfig, rows) {
  if (!rows.length) return [];
  const keyRows = rows.map((row) => ({ [tableConfig.keyColumn]: row[tableConfig.keyColumn] }));
  const result = await invokeAppSheetAction(config, tableConfig, "Find", keyRows);
  return Array.isArray(result?.Rows) ? result.Rows : [];
}

async function upsertRows(config, tableConfig, rows) {
  let added = 0;
  let edited = 0;

  for (const batch of chunk(rows, UPSERT_BATCH_SIZE)) {
    const existing = await findExistingRows(config, tableConfig, batch);
    const existingKeys = new Set(
      existing
        .map((row) => row?.[tableConfig.keyColumn])
        .filter((value) => value !== undefined && value !== null)
        .map((value) => String(value))
    );

    const toAdd = batch.filter((row) => !existingKeys.has(String(row[tableConfig.keyColumn])));
    const toEdit = batch.filter((row) => existingKeys.has(String(row[tableConfig.keyColumn])));

    if (toAdd.length) {
      await invokeAppSheetAction(config, tableConfig, "Add", toAdd);
      added += toAdd.length;
    }
    if (toEdit.length) {
      await invokeAppSheetAction(config, tableConfig, "Edit", toEdit);
      edited += toEdit.length;
    }
  }

  return { added, edited };
}

async function assertAdmin(request) {
  const email = String(request.auth?.token?.email || "").trim().toLowerCase();
  if (!request.auth || !email) throw new HttpsError("unauthenticated", "Login richiesto.");
  if (email === ADMIN_EMAIL) return;

  const snapshot = await admin.firestore().collection("appConfig").doc("adminUsers").get();
  const emails = snapshot.exists && Array.isArray(snapshot.data()?.emails) ? snapshot.data().emails : [];
  const allowed = new Set(emails.map((value) => String(value || "").trim().toLowerCase()).filter(Boolean));
  if (!allowed.has(email)) throw new HttpsError("permission-denied", "Solo admin può avviare la sincronizzazione AppSheet.");
}

exports.syncCommessaToAppSheet = onDocumentWritten(
  {
    document: "commesse/{commessaId}",
    region: REGION,
    secrets: [APPSHEET_ACCESS_KEY]
  },
  async (event) => syncSingleChange({
    entity: "commesse",
    before: event.data?.before,
    after: event.data?.after,
    id: String(event.params?.commessaId || "").trim(),
    path: `commesse/${String(event.params?.commessaId || "").trim()}`
  })
);

exports.syncImpiantoToAppSheet = onDocumentWritten(
  {
    document: "commesse/{commessaId}/impianti/{impiantoId}",
    region: REGION,
    secrets: [APPSHEET_ACCESS_KEY]
  },
  async (event) => syncSingleChange({
    entity: "impianti",
    before: event.data?.before,
    after: event.data?.after,
    id: String(event.params?.impiantoId || "").trim(),
    parentId: String(event.params?.commessaId || "").trim(),
    path: `commesse/${String(event.params?.commessaId || "").trim()}/impianti/${String(event.params?.impiantoId || "").trim()}`
  })
);

exports.backfillAppSheetFromVarga = onCall(
  {
    region: REGION,
    secrets: [APPSHEET_ACCESS_KEY],
    timeoutSeconds: 540,
    memory: "512MiB"
  },
  async (request) => {
    await assertAdmin(request);
    const config = await loadConfig({ force: true });
    const validation = validateConfig(config);
    if (!validation.valid) {
      throw new HttpsError("failed-precondition", `Configurazione AppSheet non pronta: ${validation.reason}`);
    }
    if (!getAccessKey()) {
      throw new HttpsError("failed-precondition", "APPSHEET_ACCESS_KEY non configurata.");
    }

    const db = admin.firestore();
    const [commesseSnapshot, impiantiSnapshot] = await Promise.all([
      db.collection("commesse").get(),
      db.collectionGroup("impianti").get()
    ]);

    const commesseRows = commesseSnapshot.docs.map((doc) => buildMappedRow(
      doc.data() || {},
      config.tables.commesse,
      { id: doc.id, path: doc.ref.path, deleted: false }
    ));

    const impiantiRows = impiantiSnapshot.docs.map((doc) => {
      const commessaId = doc.ref.parent.parent?.id || "";
      return buildMappedRow(
        doc.data() || {},
        config.tables.impianti,
        { id: doc.id, parentId: commessaId, path: doc.ref.path, deleted: false }
      );
    });

    const validCommesseRows = commesseRows.filter((row) => {
      const value = row[config.tables.commesse.keyColumn];
      return value !== undefined && value !== null && String(value).trim() !== "";
    });
    const validImpiantiRows = impiantiRows.filter((row) => {
      const value = row[config.tables.impianti.keyColumn];
      return value !== undefined && value !== null && String(value).trim() !== "";
    });

    const commesseResult = await upsertRows(config, config.tables.commesse, validCommesseRows);
    const impiantiResult = await upsertRows(config, config.tables.impianti, validImpiantiRows);

    return {
      ok: true,
      appId: config.appId,
      commesse: {
        total: validCommesseRows.length,
        ...commesseResult
      },
      impianti: {
        total: validImpiantiRows.length,
        ...impiantiResult
      },
      note: "Sincronizzazione non distruttiva: nessuna riga AppSheet è stata eliminata."
    };
  }
);

exports.__test = {
  chunk,
  loadConfig,
  invokeAppSheetAction,
  syncSingleChange,
  findExistingRows,
  upsertRows
};
