"use strict";

const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");

const REGION = "europe-west1";
const ADMIN_EMAIL = "ionut29019@gmail.com";
const MONTHLY_EMAIL_LIMIT = 2500;
const COUNTER_COLLECTION = "systemCounters";
const COUNTER_PREFIX = "errorEmails_";

function monthKeyRome(date = new Date()) {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/Rome",
    year: "numeric",
    month: "2-digit"
  }).formatToParts(date);
  const year = parts.find((part) => part.type === "year")?.value || String(date.getUTCFullYear());
  const month = parts.find((part) => part.type === "month")?.value || String(date.getUTCMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}

function ensureAdmin(context) {
  if (!context.auth?.uid) {
    throw new functions.https.HttpsError("unauthenticated", "Accesso necessario.");
  }
  const email = String(context.auth.token?.email || "").trim().toLowerCase();
  if (email !== ADMIN_EMAIL) {
    throw new functions.https.HttpsError("permission-denied", "Funzione riservata all'amministratore.");
  }
}

exports.getErrorEmailUsage = functions
  .region(REGION)
  .https.onCall(async (_data, context) => {
    ensureAdmin(context);

    const month = monthKeyRome();
    const ref = admin.firestore().collection(COUNTER_COLLECTION).doc(`${COUNTER_PREFIX}${month}`);
    const snapshot = await ref.get();
    const data = snapshot.data() || {};
    const sent = Math.max(0, Number(data.sentCount) || 0);
    const limit = Math.max(1, Number(data.limit) || MONTHLY_EMAIL_LIMIT);
    const remaining = Math.max(0, limit - sent);
    const percentage = Math.min(100, Math.round((sent / limit) * 1000) / 10);
    const updatedAt = data.updatedAt?.toDate?.()?.toISOString?.() || null;

    return {
      month,
      sent,
      limit,
      remaining,
      percentage,
      updatedAt
    };
  });
