const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");
const crypto = require("crypto");
const { Readable } = require("stream");
const { google } = require("googleapis");
const { CACHE_TTL_MS: FUEL_CACHE_TTL_MS, downloadNationalSnapshot } = require("./fuel-stations-cache");

admin.initializeApp();

const ADMIN_EMAIL = "ionut29019@gmail.com";
const CENTRAL_DRIVE_ROOT_FOLDER_ID = "1s6qmv2SsiTUbCjqFX4yIk4VoPQayFrU0";
const CENTRAL_DRIVE_ROOT_FOLDER_NAME = "Varga Cantieri";
const MAX_UPLOAD_BYTES = 20 * 1024 * 1024;
const FCM_BATCH_SIZE = 500;
const INVALID_FCM_TOKEN_CODES = new Set([
  "messaging/invalid-registration-token",
  "messaging/registration-token-not-registered"
]);

function chunkItems(items, size) {
  const chunks = [];
  for (let index = 0; index < items.length; index += size) chunks.push(items.slice(index, index + size));
  return chunks;
}

async function removeInvalidPushTokens(db, invalidTokens) {
  if (!invalidTokens.size) return;
  const users = await db.collection("platformUsers").where("pushToken", "in", [...invalidTokens].slice(0, 30)).get();
  const batch = db.batch();
  users.forEach((user) => batch.update(user.ref, {
    pushToken: admin.firestore.FieldValue.delete(),
    pushTokenInvalidatedAt: admin.firestore.FieldValue.serverTimestamp()
  }));
  if (!users.empty) await batch.commit();
}

exports.notifyAllUsersWhenImpiantoDone = functions.region("europe-west1").firestore
  .document("appNotifications/{notificationId}")
  .onCreate(async (snapshot) => {
    const event = snapshot.data() || {};
    if (event.eventType !== "impianto-done") return null;

    const db = admin.firestore();
    const users = await db.collection("platformUsers").get();
    const tokens = [...new Set(users.docs.map((user) => String(user.data().pushToken || "").trim()).filter(Boolean))];
    if (!tokens.length) {
      console.info("Notifica FATTO non inviata: nessun dispositivo Android registrato.", { notificationId: snapshot.id });
      return null;
    }

    const invalidTokens = new Set();
    let successCount = 0;
    let failureCount = 0;
    for (const tokenBatch of chunkItems(tokens, FCM_BATCH_SIZE)) {
      const response = await admin.messaging().sendEachForMulticast({
        tokens: tokenBatch,
        notification: {
          title: String(event.title || "Impianto completato").slice(0, 120),
          body: String(event.body || "Un utente ha premuto FATTO su un impianto.").slice(0, 500)
        },
        data: {
          eventType: "impianto-done",
          notificationId: snapshot.id,
          commessaId: String(event.commessaId || ""),
          impiantoKey: String(event.impiantoKey || "")
        },
        android: {
          priority: "high",
          notification: {
            channelId: "hera_operational_updates",
            sound: "default",
            tag: `impianto-done-${snapshot.id}`
          }
        }
      });
      successCount += response.successCount;
      failureCount += response.failureCount;
      response.responses.forEach((result, index) => {
        if (!result.success && INVALID_FCM_TOKEN_CODES.has(result.error?.code)) invalidTokens.add(tokenBatch[index]);
      });
    }

    for (const invalidBatch of chunkItems([...invalidTokens], 30)) {
      await removeInvalidPushTokens(db, new Set(invalidBatch));
    }
    console.info("Notifica Android FATTO inviata.", { notificationId: snapshot.id, successCount, failureCount });
    return null;
  });

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


const FUEL_CACHE_OBJECT = "public-cache/fuel-stations-italy.json";
let fuelSnapshotMemory = null;
let fuelSnapshotRefreshPromise = null;

function isFreshFuelSnapshot(snapshot) {
  return Array.isArray(snapshot?.stations)
    && snapshot.stations.length > 0
    && Date.now() - Number(snapshot.updatedAt || 0) < FUEL_CACHE_TTL_MS;
}

async function readStoredFuelSnapshot() {
  if (isFreshFuelSnapshot(fuelSnapshotMemory)) return fuelSnapshotMemory;
  try {
    const file = admin.storage().bucket().file(FUEL_CACHE_OBJECT);
    const [exists] = await file.exists();
    if (!exists) return null;
    const [buffer] = await file.download();
    const snapshot = JSON.parse(buffer.toString("utf8"));
    if (!Array.isArray(snapshot?.stations) || !snapshot.stations.length) return null;
    fuelSnapshotMemory = snapshot;
    return snapshot;
  } catch (error) {
    console.warn("Archivio distributori su Cloud Storage non leggibile:", error.message || error);
    return null;
  }
}

async function writeStoredFuelSnapshot(snapshot) {
  try {
    await admin.storage().bucket().file(FUEL_CACHE_OBJECT).save(JSON.stringify(snapshot), {
      resumable: false,
      metadata: {
        contentType: "application/json",
        cacheControl: "private,max-age=3600"
      }
    });
  } catch (error) {
    console.warn("Archivio distributori non salvato su Cloud Storage:", error.message || error);
  }
  fuelSnapshotMemory = snapshot;
  return snapshot;
}

async function getNationalFuelSnapshot(force = false) {
  if (fuelSnapshotRefreshPromise) return fuelSnapshotRefreshPromise;
  const stored = await readStoredFuelSnapshot();
  if (!force && isFreshFuelSnapshot(stored)) return stored;

  fuelSnapshotRefreshPromise = downloadNationalSnapshot(fetch)
    .then(writeStoredFuelSnapshot)
    .finally(() => {
      fuelSnapshotRefreshPromise = null;
    });
  return fuelSnapshotRefreshPromise;
}

exports.refreshFuelStationsItaly = functions
  .runWith({ timeoutSeconds: 120, memory: "512MB", maxInstances: 2 })
  .pubsub.schedule("30 3 * * *")
  .timeZone("Europe/Rome")
  .onRun(async () => {
    const snapshot = await getNationalFuelSnapshot(true);
    console.info("Archivio distributori MIMIT aggiornato", {
      stations: snapshot.stations.length,
      extractionDate: snapshot.extractionDate
    });
    return null;
  });

exports.getFuelStationsItaly = functions
  .runWith({ timeoutSeconds: 120, memory: "512MB", maxInstances: 5 })
  .https.onRequest(async (request, response) => {
    response.set("Access-Control-Allow-Origin", "*");
    response.set("Cache-Control", "public,max-age=3600,s-maxage=3600,stale-while-revalidate=86400");
    if (request.method === "OPTIONS") {
      response.set("Access-Control-Allow-Methods", "GET");
      response.set("Access-Control-Allow-Headers", "Content-Type");
      response.status(204).send("");
      return;
    }
    if (request.method !== "GET") {
      response.status(405).json({ error: "Metodo non consentito" });
      return;
    }
    try {
      const snapshot = await getNationalFuelSnapshot(false);
      response.status(200).json(snapshot);
    } catch (error) {
      console.error("Archivio distributori MIMIT non disponibile:", error);
      response.status(503).json({ error: "Archivio distributori temporaneamente non disponibile" });
    }
  });
