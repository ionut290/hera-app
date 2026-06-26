const admin = require("firebase-admin");
const functions = require("firebase-functions");
const crypto = require("crypto");
const { Readable } = require("stream");
const { google } = require("googleapis");

admin.initializeApp();

const ADMIN_EMAIL = "ionut29019@gmail.com";
const CENTRAL_DRIVE_ROOT_FOLDER_ID = "1s6qmv2SsiTUbCjqFX4yIk4VoPQayFrU0";
const CENTRAL_DRIVE_ROOT_FOLDER_NAME = "Varga Cantieri";
const MAX_UPLOAD_BYTES = 20 * 1024 * 1024;

function normalizeFolderName(value, fallback = "Generale") {
  return String(value || fallback).trim().replace(/[\\/:*?"<>|]+/g, "-").slice(0, 120) || fallback;
}

function escapeDriveQueryValue(value) {
  return String(value || "").replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

async function getAdminEmails(db) {
  const snapshot = await db.collection("appConfig").doc("adminUsers").get();
  const configured = snapshot.exists && Array.isArray(snapshot.data().emails) ? snapshot.data().emails : [];
  return new Set([ADMIN_EMAIL, ...configured].map((email) => String(email || "").trim().toLowerCase()).filter(Boolean));
}

async function assertAdmin(context, db) {
  const email = String(context.auth?.token?.email || "").trim().toLowerCase();
  if (!context.auth || !email) {
    throw new functions.https.HttpsError("unauthenticated", "Login richiesto.");
  }
  const adminEmails = await getAdminEmails(db);
  if (!adminEmails.has(email)) {
    throw new functions.https.HttpsError("permission-denied", "Solo admin può configurare Drive.");
  }
}

async function buildDriveClient(db) {
  const secretSnapshot = await db.collection("appConfig").doc("driveAdminSecret").get();
  const secret = secretSnapshot.exists ? secretSnapshot.data() : null;
  if (!secret || (!secret.accessToken && !secret.refreshToken)) {
    throw new functions.https.HttpsError("failed-precondition", "Cloud amministratore non configurato");
  }

  const oauth2 = new google.auth.OAuth2(
    functions.config().google?.client_id,
    functions.config().google?.client_secret
  );
  oauth2.setCredentials({
    access_token: secret.accessToken || undefined,
    refresh_token: secret.refreshToken || undefined
  });

  return { drive: google.drive({ version: "v3", auth: oauth2 }), secret };
}

function getFolderRegistryId(parentId, name) {
  return crypto.createHash("sha1").update(`${parentId || "root"}/${name}`).digest("hex");
}

async function findExistingFolder(drive, name, parentId = "") {
  const clauses = [
    `name='${escapeDriveQueryValue(name)}'`,
    "mimeType='application/vnd.google-apps.folder'",
    "trashed=false"
  ];
  if (parentId) clauses.push(`'${parentId}' in parents`);
  const found = await drive.files.list({
    q: clauses.join(" and "),
    fields: "files(id,name,createdTime)",
    orderBy: "createdTime",
    pageSize: 1
  });
  return found.data.files && found.data.files[0] ? found.data.files[0] : null;
}

async function waitForRegisteredFolder(folderRef) {
  for (let attempt = 0; attempt < 10; attempt += 1) {
    await new Promise((resolve) => setTimeout(resolve, 300));
    const snapshot = await folderRef.get();
    const folderId = snapshot.exists ? String(snapshot.data().folderId || "") : "";
    if (folderId) return folderId;
  }
  return "";
}

async function getOrCreateFolder(db, drive, name, parentId = "") {
  const normalizedName = normalizeFolderName(name);
  const folderRef = db.collection("centralDriveFolders").doc(getFolderRegistryId(parentId, normalizedName));
  const registered = await folderRef.get();
  const registeredFolderId = registered.exists ? String(registered.data().folderId || "") : "";
  if (registeredFolderId) return registeredFolderId;

  try {
    await folderRef.create({
      parentId: parentId || "",
      name: normalizedName,
      status: "creating",
      createdAt: admin.firestore.FieldValue.serverTimestamp()
    });
  } catch (error) {
    if (error.code === 6 || error.code === "already-exists") {
      const waitedFolderId = await waitForRegisteredFolder(folderRef);
      if (waitedFolderId) return waitedFolderId;
    } else {
      throw error;
    }
  }

  const existing = await findExistingFolder(drive, normalizedName, parentId);
  const folderId = existing?.id || (await drive.files.create({
    requestBody: {
      name: normalizedName,
      mimeType: "application/vnd.google-apps.folder",
      ...(parentId ? { parents: [parentId] } : {})
    },
    fields: "id"
  })).data.id;

  await folderRef.set({
    parentId: parentId || "",
    name: normalizedName,
    folderId,
    status: "ready",
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  return folderId;
}

exports.uploadCentralDriveFile = functions.https.onCall(async (data, context) => {
  if (!context.auth) {
    throw new functions.https.HttpsError("unauthenticated", "Login richiesto.");
  }
  const base64 = String(data?.base64 || "");
  const fileName = normalizeFolderName(data?.fileName, "file");
  const mimeType = String(data?.mimeType || "application/octet-stream");
  const commessaName = normalizeFolderName(data?.commessaName, "Generale");
  const driveType = normalizeFolderName(data?.driveType, "EXPORT").toUpperCase();
  const buffer = Buffer.from(base64, "base64");
  if (!buffer.length) {
    throw new functions.https.HttpsError("invalid-argument", "File vuoto.");
  }
  if (buffer.length > MAX_UPLOAD_BYTES) {
    throw new functions.https.HttpsError("invalid-argument", "File troppo grande per upload backend.");
  }

  const db = admin.firestore();
  const { drive } = await buildDriveClient(db);
  const rootFolderId = CENTRAL_DRIVE_ROOT_FOLDER_ID;
  const commessaFolderId = await getOrCreateFolder(db, drive, commessaName, rootFolderId);
  const typeFolderId = await getOrCreateFolder(db, drive, driveType, commessaFolderId);

  const uploaded = await drive.files.create({
    requestBody: {
      name: fileName,
      mimeType,
      parents: [typeFolderId]
    },
    media: {
      mimeType,
      body: Readable.from(buffer)
    },
    fields: "id,name,webViewLink"
  });

  await db.collection("centralDriveUploads").add({
    fileId: uploaded.data.id || "",
    fileName,
    mimeType,
    commessaName,
    driveType,
    createdByUid: context.auth.uid,
    createdByEmail: context.auth.token.email || "",
    createdAt: admin.firestore.FieldValue.serverTimestamp()
  });

  return {
    fileId: uploaded.data.id || "",
    webViewLink: uploaded.data.webViewLink || ""
  };
});

exports.configureCentralDrive = functions.https.onCall(async (data, context) => {
  const db = admin.firestore();
  await assertAdmin(context, db);
  const ownerEmail = String(context.auth.token.email || ADMIN_EMAIL);
  const accessToken = String(data?.accessToken || "");
  const refreshToken = String(data?.refreshToken || "");
  if (!accessToken && !refreshToken) {
    throw new functions.https.HttpsError("invalid-argument", "Token Drive mancante.");
  }
  await db.collection("appConfig").doc("driveAdminSecret").set({
    ownerEmail,
    accessToken,
    refreshToken,
    rootFolderId: CENTRAL_DRIVE_ROOT_FOLDER_ID,
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  await db.collection("appConfig").doc("driveBridge").set({
    ownerEmail,
    configured: true,
    rootFolderId: CENTRAL_DRIVE_ROOT_FOLDER_ID,
    rootFolderName: CENTRAL_DRIVE_ROOT_FOLDER_NAME,
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  return { ok: true };
});

const WEATHER_ALERT_TYPES = ["caldo", "temporali", "vento forte", "pioggia intensa", "neve", "ghiaccio", "rischio idraulico", "rischio idrogeologico", "incendi"];
const ALERT_LEVELS = new Set(["giallo", "arancione", "rosso"]);

function normalizeAlertText(value) {
  return String(value || "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ");
}

function weatherAlertDocId(comune, data, tipoAllerta) {
  return `${normalizeAlertText(comune).replace(/[^a-z0-9]+/g, "_")}_${String(data || "").slice(0, 10)}_${normalizeAlertText(tipoAllerta).replace(/[^a-z0-9]+/g, "_")}`.replace(/^_+|_+$/g, "");
}

function dateKeyFromDate(date) {
  return date.toISOString().slice(0, 10);
}

function mapAlertLevel(value) {
  const normalized = normalizeAlertText(value);
  if (normalized.includes("rosso") || normalized.includes("elevat") || normalized.includes("alto")) return "rosso";
  if (normalized.includes("aranc") || normalized.includes("moderat")) return "arancione";
  if (normalized.includes("giall") || normalized.includes("ordin") || normalized.includes("basso")) return "giallo";
  return ALERT_LEVELS.has(normalized) ? normalized : "giallo";
}

function normalizeAlertType(value) {
  const normalized = normalizeAlertText(value);
  const match = WEATHER_ALERT_TYPES.find((type) => normalized.includes(normalizeAlertText(type)));
  if (match) return match;
  if (normalized.includes("idro")) return normalized.includes("geolog") ? "rischio idrogeologico" : "rischio idraulico";
  if (normalized.includes("piogg") || normalized.includes("precipit")) return "pioggia intensa";
  if (normalized.includes("vento")) return "vento forte";
  if (normalized.includes("tempor")) return "temporali";
  if (normalized.includes("incend")) return "incendi";
  return String(value || "altro evento").trim().toLowerCase() || "altro evento";
}

async function fetchJsonIfConfigured(url, label) {
  if (!url) return null;
  const response = await fetch(url, { headers: { "accept": "application/json" } });
  if (!response.ok) throw new Error(`${label} HTTP ${response.status}`);
  return response.json();
}

function normalizeExternalAlerts(payload, comune, source) {
  const items = Array.isArray(payload) ? payload : Array.isArray(payload?.alerts) ? payload.alerts : Array.isArray(payload?.data) ? payload.data : [];
  return items.map((item) => {
    const data = String(item.data || item.date || item.validDate || item.valid_from || item.validFrom || "").slice(0, 10) || dateKeyFromDate(new Date());
    const validFrom = item.validFrom || item.valid_from || `${data}T00:00:00.000Z`;
    const validTo = item.validTo || item.valid_to || `${data}T23:59:59.999Z`;
    return {
      comune,
      data,
      tipoAllerta: normalizeAlertType(item.tipoAllerta || item.event || item.type || item.risk || source),
      livello: mapAlertLevel(item.livello || item.level || item.color || item.riskLevel),
      descrizione: String(item.descrizione || item.description || item.message || `${source}: allerta disponibile`).trim(),
      validFrom: admin.firestore.Timestamp.fromDate(new Date(validFrom)),
      validTo: admin.firestore.Timestamp.fromDate(new Date(validTo)),
      fonte: source
    };
  }).filter((item) => item.comune && item.data && item.tipoAllerta);
}

async function collectComuniByCommessa(db) {
  const snapshot = await db.collectionGroup("impianti").get();
  const comuniByCommessa = new Map();
  snapshot.forEach((doc) => {
    const comune = String(doc.data()?.comune || doc.data()?.citta || doc.data()?.localita || "").trim();
    const commessaId = doc.ref.parent.parent?.id || "";
    if (!comune || !commessaId) return;
    if (!comuniByCommessa.has(commessaId)) comuniByCommessa.set(commessaId, new Set());
    comuniByCommessa.get(commessaId).add(comune);
  });
  return comuniByCommessa;
}

async function fetchAlertsForComune(comune) {
  const cfg = functions.config().weather || {};
  const encoded = encodeURIComponent(comune);
  const worklimateUrl = cfg.worklimate_url ? String(cfg.worklimate_url).replace("{comune}", encoded) : "";
  const civilUrl = cfg.civil_protection_url ? String(cfg.civil_protection_url).replace("{comune}", encoded) : "";
  const [worklimate, civil] = await Promise.allSettled([
    fetchJsonIfConfigured(worklimateUrl, "Worklimate"),
    fetchJsonIfConfigured(civilUrl, "Protezione Civile")
  ]);
  const alerts = [];
  if (worklimate.status === "fulfilled" && worklimate.value) {
    alerts.push(...normalizeExternalAlerts(worklimate.value, comune, "Worklimate").map((item) => ({
      ...item,
      tipoAllerta: "caldo",
      fasciaOraria: item.fasciaOraria || "12:00",
      scenario: "lavoratore esposto al sole, attività fisica intensa, previsione ore 12:00"
    })));
  }
  if (civil.status === "fulfilled" && civil.value) alerts.push(...normalizeExternalAlerts(civil.value, comune, "Protezione Civile / allerte regionali"));
  if (worklimate.status === "rejected") console.warn("Worklimate non aggiornato", comune, worklimate.reason?.message || worklimate.reason);
  if (civil.status === "rejected") console.warn("Allerte Protezione Civile non aggiornate", comune, civil.reason?.message || civil.reason);
  return alerts;
}

async function commitWeatherAlertWrites(db, writes) {
  for (let index = 0; index < writes.length; index += 450) {
    const batch = db.batch();
    writes.slice(index, index + 450).forEach((write) => batch.set(write.ref, write.data, { merge: true }));
    await batch.commit();
  }
}

exports.updateWeatherAlerts = functions.pubsub.schedule("every 2 hours").timeZone("Europe/Rome").onRun(async () => {
  const db = admin.firestore();
  const comuniByCommessa = await collectComuniByCommessa(db);
  const comuni = Array.from(new Set(Array.from(comuniByCommessa.values()).flatMap((set) => Array.from(set))));
  const now = admin.firestore.Timestamp.now();
  const allAlerts = [];
  for (const comune of comuni) allAlerts.push(...await fetchAlertsForComune(comune));

  const writes = [];
  allAlerts.forEach((alert) => {
    writes.push({
      ref: db.collection("weatherAlerts").doc(weatherAlertDocId(alert.comune, alert.data, alert.tipoAllerta)),
      data: { ...alert, active: true, updatedAt: now }
    });
  });
  Array.from(comuniByCommessa.entries()).forEach(([commessaId, comuniSet]) => {
    writes.push({
      ref: db.collection("commesse").doc(commessaId),
      data: { weatherAlertComuni: Array.from(comuniSet), weatherAlertComuniUpdatedAt: now }
    });
  });
  const expired = await db.collection("weatherAlerts").where("validTo", "<", now).where("active", "==", true).limit(450).get();
  expired.forEach((doc) => writes.push({ ref: doc.ref, data: { active: false, updatedAt: now } }));
  await commitWeatherAlertWrites(db, writes);
  return { comuni: comuni.length, alerts: allAlerts.length, expired: expired.size };
});

const WORKLIMATE_OPERATIONAL_ADVICE = [
  "Evitare le ore più calde.",
  "Aumentare le pause.",
  "Bere acqua.",
  "Lavorare all’ombra quando possibile.",
  "Modificare orario in caso di rischio alto."
];

function normalizeWorklimateRiskLevel(value) {
  const normalized = normalizeAlertText(value);
  if (normalized.includes("ross") || normalized.includes("emerg") || normalized.includes("molto") || normalized.includes("alto")) return "rosso";
  if (normalized.includes("aranc") || normalized.includes("medio") || normalized.includes("moderat")) return "arancione";
  if (normalized.includes("giall") || normalized.includes("atten") || normalized.includes("basso")) return "giallo";
  return "verde";
}

function buildWorklimateFallbackRisk(impianto) {
  return {
    riskLevel: "verde",
    source: "Worklimate - ultimo dato cloud/fallback",
    forecastAt: admin.firestore.Timestamp.fromDate(new Date()),
    operationalAdvice: WORKLIMATE_OPERATIONAL_ADVICE,
    raw: { fallback: true, reason: "Endpoint Worklimate non configurato" }
  };
}

async function fetchWorklimateRiskForImpianto(impianto) {
  const endpoint = functions.config().worklimate?.endpoint || process.env.WORKLIMATE_ENDPOINT || "";
  if (!endpoint) return buildWorklimateFallbackRisk(impianto);
  const url = new URL(endpoint);
  url.searchParams.set("lat", String(impianto.gpsY));
  url.searchParams.set("lon", String(impianto.gpsX));
  if (impianto.comune) url.searchParams.set("comune", String(impianto.comune));
  const payload = await fetchJsonIfConfigured(url.toString(), "WORKLIMATE");
  const riskLevel = normalizeWorklimateRiskLevel(payload?.riskLevel || payload?.risk || payload?.level || payload?.livelloRischio);
  const forecastRaw = payload?.forecastAt || payload?.forecast_at || payload?.validAt || payload?.dataOraPrevisione || new Date().toISOString();
  const forecastDate = new Date(forecastRaw);
  return {
    riskLevel,
    source: payload?.source || payload?.fonte || "Worklimate",
    forecastAt: Number.isNaN(forecastDate.getTime()) ? admin.firestore.Timestamp.fromDate(new Date()) : admin.firestore.Timestamp.fromDate(forecastDate),
    operationalAdvice: Array.isArray(payload?.operationalAdvice) && payload.operationalAdvice.length ? payload.operationalAdvice : WORKLIMATE_OPERATIONAL_ADVICE,
    raw: payload || null
  };
}

exports.updateWorklimateRisk = functions.pubsub.schedule("0 6,12,18 * * *").timeZone("Europe/Rome").onRun(async () => {
  const db = admin.firestore();
  const snapshot = await db.collectionGroup("impianti").get();
  let batch = db.batch();
  let count = 0;
  for (const doc of snapshot.docs) {
    const impianto = { id: doc.id, ...doc.data() };
    const lat = Number(impianto.gpsY);
    const lon = Number(impianto.gpsX);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) continue;
    const risk = await fetchWorklimateRiskForImpianto({ ...impianto, gpsY: lat, gpsX: lon });
    const worklimateDocId = crypto.createHash("sha1").update(doc.ref.path).digest("hex");
    batch.set(db.collection("worklimateRiskByImpianto").doc(worklimateDocId), {
      impiantoId: doc.id,
      impiantoPath: doc.ref.path,
      impiantoName: impianto.denominazione || "",
      comune: impianto.comune || "",
      coordinates: new admin.firestore.GeoPoint(lat, lon),
      lat,
      lon,
      riskLevel: risk.riskLevel,
      forecastAt: risk.forecastAt,
      source: risk.source || "Worklimate",
      operationalAdvice: risk.operationalAdvice || WORKLIMATE_OPERATIONAL_ADVICE,
      raw: risk.raw || null,
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    count += 1;
    if (count % 450 === 0) {
      await batch.commit();
      batch = db.batch();
    }
  }
  if (count % 450 !== 0) await batch.commit();
  await db.collection("appConfig").doc("worklimateUpdate").set({ lastRunAt: admin.firestore.FieldValue.serverTimestamp(), count }, { merge: true });
  return null;
});

const LAVAGNA_DEFAULT_URL = "https://coopavola.eggsnext.cloud/main/functions/app/eggs-lavagna/lavagna";

function getLavagnaConfig() {
  const cfg = functions.config().lavagna || {};
  return {
    url: process.env.LAVAGNA_URL || cfg.url || LAVAGNA_DEFAULT_URL,
    username: process.env.LAVAGNA_USERNAME || cfg.username || "",
    password: process.env.LAVAGNA_PASSWORD || cfg.password || ""
  };
}

function parseLavagnaDate(value) {
  const raw = String(value || "").trim();
  if (!raw) return dateKeyFromDate(new Date());
  const iso = raw.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  const ita = raw.match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);
  if (ita) {
    const year = ita[3].length === 2 ? `20${ita[3]}` : ita[3];
    return `${year}-${ita[2].padStart(2, "0")}-${ita[1].padStart(2, "0")}`;
  }
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? dateKeyFromDate(new Date()) : dateKeyFromDate(parsed);
}

function normalizeLavagnaRow(row) {
  const codiceCantiere = String(row.codiceCantiere || row.codice_cantiere || row.codice || row.cantiere || row.commessaCodice || row.commessa_codice || "").trim();
  const dateKey = parseLavagnaDate(row.data || row.date || row.giorno || row.riferimentoData || row.riferimento_data);
  const squadra = {
    caposquadra: String(row.caposquadra || row.responsabile || row.leader || "").trim(),
    personale: Array.isArray(row.personale) ? row.personale.join(", ") : String(row.personale || row.operatori || row.addetti || row.squadra || "").trim(),
    mezzi: Array.isArray(row.mezzi) ? row.mezzi.join(", ") : String(row.mezzi || row.veicoli || row.attrezzature || "").trim(),
    impianti: Array.isArray(row.impianti) ? row.impianti.join(", ") : String(row.impianti || row.zona || row.lavorazione || "").trim(),
    note: String(row.note || row.descrizione || "").trim(),
    orario: String(row.orario || row.oraInizio || row.ora_inizio || "").trim(),
    orarioFine: String(row.orarioFine || row.oraFine || row.ora_fine || "").trim()
  };
  return { codiceCantiere, dateKey, squadra };
}

function extractJsonFromLavagnaPayload(payload) {
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload?.data)) return payload.data;
  if (Array.isArray(payload?.rows)) return payload.rows;
  if (Array.isArray(payload?.items)) return payload.items;
  if (Array.isArray(payload?.records)) return payload.records;
  return [];
}

function extractLavagnaRowsFromHtml(html) {
  const tableRows = Array.from(String(html || "").matchAll(/<tr[\s\S]*?<\/tr>/gi));
  if (tableRows.length < 2) return [];
  const cellsFor = (tr) => Array.from(tr.matchAll(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi)).map((m) => m[1].replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").replace(/\s+/g, " ").trim());
  const headers = cellsFor(tableRows[0][0]).map((h) => h.toLowerCase().replace(/[^a-z0-9]+/g, "_"));
  return tableRows.slice(1).map((match) => {
    const cells = cellsFor(match[0]);
    return headers.reduce((row, header, index) => ({ ...row, [header]: cells[index] || "" }), {});
  });
}

async function fetchLavagnaSource() {
  const cfg = getLavagnaConfig();
  if (!cfg.username || !cfg.password) {
    throw new functions.https.HttpsError("failed-precondition", "Credenziali Lavagna mancanti: configura LAVAGNA_USERNAME e LAVAGNA_PASSWORD nelle variabili ambiente delle Cloud Functions.");
  }
  const authHeader = `Basic ${Buffer.from(`${cfg.username}:${cfg.password}`).toString("base64")}`;
  const response = await fetch(cfg.url, { headers: { "accept": "application/json,text/html;q=0.9,*/*;q=0.8", "authorization": authHeader } });
  if (!response.ok) throw new functions.https.HttpsError("unavailable", `Lavagna non disponibile: HTTP ${response.status}`);
  const contentType = response.headers.get("content-type") || "";
  const text = await response.text();
  if (contentType.includes("application/json") || text.trim().startsWith("{") || text.trim().startsWith("[")) {
    return extractJsonFromLavagnaPayload(JSON.parse(text));
  }
  return extractLavagnaRowsFromHtml(text);
}

async function syncLavagnaRowsToFirestore(db, rows) {
  const commesseSnapshot = await db.collection("commesse").get();
  const commesseByCode = new Map();
  commesseSnapshot.forEach((doc) => {
    const code = String(doc.data()?.codice || "").trim().toLowerCase();
    if (code) commesseByCode.set(code, { id: doc.id, nome: doc.data()?.nome || "Commessa" });
  });
  const grouped = new Map();
  const skippedUnknownCodes = new Set();
  rows.map(normalizeLavagnaRow).forEach((item) => {
    if (!item.codiceCantiere || !item.dateKey || !isSquadraRowFilledServer(item.squadra)) return;
    const commessa = commesseByCode.get(item.codiceCantiere.toLowerCase());
    if (!commessa) {
      skippedUnknownCodes.add(item.codiceCantiere);
      return;
    }
    const key = `${item.dateKey}__${commessa.id}`;
    if (!grouped.has(key)) grouped.set(key, { commessa, dateKey: item.dateKey, rows: [] });
    grouped.get(key).rows.push(item.squadra);
  });
  const now = admin.firestore.FieldValue.serverTimestamp();
  let batch = db.batch();
  let writes = 0;
  for (const [docId, group] of grouped.entries()) {
    const payload = { commessaId: group.commessa.id, commessaNome: group.commessa.nome, riferimentoData: group.dateKey, dateKey: group.dateKey, squadre: group.rows, source: "eggs-lavagna", updatedAt: now, updatedBy: "sync-lavagna" };
    batch.set(db.collection("squadreStorico").doc(docId), payload, { merge: true });
    batch.set(db.collection("squadreCommesse").doc(group.commessa.id), payload, { merge: true });
    writes += 2;
    if (writes >= 450) { await batch.commit(); batch = db.batch(); writes = 0; }
  }
  if (writes) await batch.commit();
  return {
    matched: grouped.size,
    importedRows: Array.from(grouped.values()).reduce((sum, group) => sum + group.rows.length, 0),
    skippedUnknownCodes: skippedUnknownCodes.size
  };
}

function isSquadraRowFilledServer(row) {
  return Boolean(row?.caposquadra || row?.personale || row?.mezzi || row?.impianti || row?.note || row?.orario || row?.orarioFine);
}

exports.syncLavagnaSquadre = functions.https.onCall(async (_data, context) => {
  const db = admin.firestore();
  await assertAdmin(context, db);
  const rawRows = await fetchLavagnaSource();
  const result = await syncLavagnaRowsToFirestore(db, rawRows);
  await db.collection("appConfig").doc("lavagnaSync").set({ ...result, lastRunAt: admin.firestore.FieldValue.serverTimestamp(), sourceUrl: getLavagnaConfig().url }, { merge: true });
  return result;
});

exports.scheduledSyncLavagnaSquadre = functions.pubsub.schedule("every 30 minutes").timeZone("Europe/Rome").onRun(async () => {
  const db = admin.firestore();
  const rawRows = await fetchLavagnaSource();
  const result = await syncLavagnaRowsToFirestore(db, rawRows);
  await db.collection("appConfig").doc("lavagnaSync").set({ ...result, lastRunAt: admin.firestore.FieldValue.serverTimestamp(), sourceUrl: getLavagnaConfig().url }, { merge: true });
  return result;
});
