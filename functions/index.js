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

function getTomorrowApiKey() {
  return String(
    process.env.TOMORROW_API_KEY
    || functions.config().tomorrow?.api_key
    || functions.config().tomorrow?.key
    || ""
  ).trim();
}

exports.getTomorrowWeatherForPlant = functions.https.onCall(async (data, context) => {
  if (!context.auth) {
    throw new functions.https.HttpsError("unauthenticated", "Login richiesto.");
  }

  const lat = Number(data?.lat);
  const lon = Number(data?.lon);
  if (!Number.isFinite(lat) || !Number.isFinite(lon)) {
    throw new functions.https.HttpsError("invalid-argument", "Coordinate impianto non valide.");
  }

  const apiKey = getTomorrowApiKey();
  if (!apiKey) {
    throw new functions.https.HttpsError("failed-precondition", "TOMORROW_API_KEY non configurata.");
  }

  const fields = [
    "temperature",
    "weatherCode",
    "precipitationProbability",
    "precipitationIntensity",
    "rainIntensity",
    "windSpeed",
    "windDirection",
    "windGust"
  ];
  const params = new URLSearchParams({
    location: `${lat},${lon}`,
    timesteps: "1h",
    units: "metric",
    fields: fields.join(","),
    apikey: apiKey
  });

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 10000);
  try {
    const response = await fetch(`https://api.tomorrow.io/v4/weather/forecast?${params.toString()}`, {
      method: "GET",
      signal: controller.signal
    });
    const text = await response.text();
    if (!response.ok) {
      throw new functions.https.HttpsError("unavailable", `Tomorrow.io ${response.status}: ${text.slice(0, 160)}`);
    }
    const parsed = JSON.parse(text);
    console.log("[getTomorrowWeatherForPlant] coordinate usate", { lat, lon });
    console.log("[getTomorrowWeatherForPlant] API response", parsed);
    return parsed;
  } catch (error) {
    if (error instanceof functions.https.HttpsError) throw error;
    throw new functions.https.HttpsError("unavailable", error?.message || "Meteo Tomorrow.io non disponibile");
  } finally {
    clearTimeout(timeoutId);
  }
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
