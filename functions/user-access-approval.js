"use strict";

const crypto = require("node:crypto");
const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");

const REGION = "europe-west1";
const BUILT_IN_ADMIN_EMAIL = "ionut29019@gmail.com";
const APP_NAME = "Varga Cantieri";
const APP_URL = "https://creative-syrniki-dddbae.netlify.app";
const ANDROID_URL = "https://play.google.com/store/apps/details?id=it.vargacantieri.hera";
const RESEND_ENDPOINT = "https://api.resend.com/emails";

function optionalEmailConfiguration() {
  let legacy = {};
  try {
    legacy = functions.config()?.resend || {};
  } catch (_) {}
  return {
    apiKey: String(process.env.RESEND_API_KEY || legacy.api_key || "").trim(),
    from: String(process.env.ERROR_REPORT_FROM || legacy.from || "").trim()
  };
}

function cleanText(value, max = 500) {
  return String(value || "").replace(/[\u0000-\u001f\u007f]+/g, " ").trim().slice(0, max);
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function normalizeRequestId(value) {
  const normalized = String(value || "").trim().replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 80);
  return normalized || crypto.randomUUID();
}

async function getAdminEmails(db) {
  const snapshot = await db.collection("appConfig").doc("adminUsers").get();
  const configured = snapshot.exists && Array.isArray(snapshot.data().emails)
    ? snapshot.data().emails
    : [];
  return new Set(
    [BUILT_IN_ADMIN_EMAIL, ...configured]
      .map((email) => String(email || "").trim().toLowerCase())
      .filter(Boolean)
  );
}

async function resolveAdministrator(context, db, requestedName) {
  if (!context.auth?.uid) {
    throw new functions.https.HttpsError("unauthenticated", "Accedi come amministratore e riprova.");
  }
  const email = String(context.auth.token?.email || "").trim().toLowerCase();
  const adminEmails = await getAdminEmails(db);
  if (!email || !adminEmails.has(email)) {
    throw new functions.https.HttpsError("permission-denied", "Solo un amministratore può sbloccare gli utenti.");
  }

  let displayName = cleanText(context.auth.token?.name, 160);
  if (!displayName) {
    try {
      displayName = cleanText((await admin.auth().getUser(context.auth.uid)).displayName, 160);
    } catch (_) {}
  }
  return {
    uid: context.auth.uid,
    email,
    displayName: displayName || cleanText(requestedName, 160) || email
  };
}

function firstValue(source, keys) {
  for (const key of keys) {
    const value = cleanText(source?.[key], 180);
    if (value) return value;
  }
  return "";
}

function buildApprovalMessage({ userName, userEmail, administratorName }) {
  return [
    `✅ ACCESSO A ${APP_NAME.toUpperCase()} APPROVATO`,
    "",
    `Ciao ${userName || ""},`,
    "",
    `l’amministratore ${administratorName} ha accettato la tua richiesta. Ora puoi accedere all’app ${APP_NAME}.`,
    "",
    "A COSA SERVE L’APP",
    "• consultare le commesse e gli impianti di lavoro;",
    "• vedere squadre, attività e informazioni operative;",
    "• aprire la navigazione verso gli impianti;",
    "• consultare documenti, comunicazioni e aggiornamenti autorizzati.",
    "",
    "COME ACCEDERE",
    `Email: ${userEmail}`,
    "Password: usa la password che hai scelto durante la registrazione.",
    "Non comunicare la password ad altre persone.",
    "",
    "INSTALLAZIONE SU ANDROID",
    `1. Apri Google Play: ${ANDROID_URL}`,
    "2. Premi Installa.",
    `3. Apri ${APP_NAME} e accedi con l’email e la password scelte.`,
    "",
    "INSTALLAZIONE SU IPHONE",
    `1. Apri con Safari: ${APP_URL}`,
    "2. Premi Condividi (quadrato con freccia verso l’alto).",
    "3. Premi Aggiungi alla schermata Home e poi Aggiungi.",
    `4. Apri l’icona ${APP_NAME} e accedi con l’email e la password scelte.`,
    "",
    `Accesso web: ${APP_URL}`,
    "",
    "Benvenuto e buon lavoro!"
  ].join("\n");
}

function buildApprovalHtml(message) {
  const lines = String(message || "").split("\n");
  const content = lines.map((line) => {
    if (!line) return "<div style=\"height:10px\"></div>";
    if (/^(✅|A COSA SERVE|COME ACCEDERE|INSTALLAZIONE SU|Accesso web:)/.test(line)) {
      return `<p style="margin:12px 0 5px;font-weight:800">${escapeHtml(line)}</p>`;
    }
    const linked = escapeHtml(line)
      .replace(escapeHtml(ANDROID_URL), `<a href="${ANDROID_URL}">${ANDROID_URL}</a>`)
      .replace(escapeHtml(APP_URL), `<a href="${APP_URL}">${APP_URL}</a>`);
    return `<p style="margin:4px 0;line-height:1.45">${linked}</p>`;
  }).join("");
  return `<!doctype html><html><body style="margin:0;padding:20px;background:#f3f6fa;font-family:Arial,sans-serif;color:#172033"><div style="max-width:680px;margin:auto;background:#fff;border-radius:18px;padding:24px;box-shadow:0 10px 35px rgba(15,23,42,.10)">${content}</div></body></html>`;
}

async function sendApprovalEmail({ apiKey, from, to, message, requestId }) {
  const response = await fetch(RESEND_ENDPOINT, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
      "User-Agent": "varga-cantieri-user-approval/1.0",
      "Idempotency-Key": `varga-user-approval-${requestId}`
    },
    body: JSON.stringify({
      from,
      to: [to],
      subject: `✅ Accesso a ${APP_NAME} approvato`,
      text: message,
      html: buildApprovalHtml(message)
    })
  });
  if (!response.ok) {
    const detail = cleanText(await response.text().catch(() => ""), 500);
    throw new Error(`Servizio email ${response.status}${detail ? `: ${detail}` : ""}`);
  }
}

exports.approveUserAccess = functions
  .region(REGION)
  .https.onCall(async (data, context) => {
    const db = admin.firestore();
    const administrator = await resolveAdministrator(context, db, data?.administratorName);
    const targetUid = cleanText(data?.targetUid, 160);
    const requestId = normalizeRequestId(data?.requestId);
    if (!targetUid) {
      throw new functions.https.HttpsError("invalid-argument", "Utente da sbloccare non valido.");
    }

    const userRef = db.collection("platformUsers").doc(targetUid);
    const snapshot = await userRef.get();
    if (!snapshot.exists) {
      throw new functions.https.HttpsError("not-found", "Il profilo utente non esiste più.");
    }
    const profile = snapshot.data() || {};
    const userEmail = cleanText(profile.email, 254).toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(userEmail)) {
      throw new functions.https.HttpsError("failed-precondition", "Il profilo non contiene un indirizzo email valido.");
    }
    const userName = firstValue(profile, ["nomeCompleto", "displayName", "fullName"]) || userEmail;
    const phone = firstValue(profile, ["whatsapp", "whatsappPhone", "telefono", "cellulare", "phone", "phoneNumber"]);
    const message = buildApprovalMessage({
      userName,
      userEmail,
      administratorName: administrator.displayName
    });

    const approvalPatch = {
      statoAccount: "attivo",
      accountStatus: "attivo",
      role: "user",
      ruolo: "user",
      approvatoAt: admin.firestore.FieldValue.serverTimestamp(),
      approvedAt: admin.firestore.FieldValue.serverTimestamp(),
      approvatoDa: administrator.uid,
      approvatoDaEmail: administrator.email,
      approvatoDaNome: administrator.displayName,
      approvedBy: administrator.email || administrator.uid,
      banned: false
    };
    const auditRef = db.collection("userAccessAudit").doc(`approval_${targetUid}_${requestId}`);
    const batch = db.batch();
    batch.update(userRef, approvalPatch);
    batch.set(auditRef, {
      userId: targetUid,
      action: "approvazione",
      reason: "",
      administratorUid: administrator.uid,
      administratorEmail: administrator.email,
      administratorName: administrator.displayName,
      requestId,
      createdAt: admin.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    await batch.commit();

    let emailSent = false;
    let emailError = "";
    const { apiKey, from } = optionalEmailConfiguration();
    if (!apiKey || !from) {
      emailError = "Invio automatico non configurato: usa il pulsante INVIA EMAIL.";
    } else {
      try {
        await sendApprovalEmail({ apiKey, from, to: userEmail, message, requestId });
        emailSent = true;
      } catch (error) {
        emailError = "Invio email non riuscito. Usa il pulsante WhatsApp.";
        console.error("Email approvazione utente non inviata.", {
          targetUid,
          requestId,
          message: cleanText(error?.message, 500)
        });
      }
    }

    return {
      approved: true,
      emailSent,
      emailError,
      userName,
      userEmail,
      phone,
      administratorName: administrator.displayName,
      message,
      requestId
    };
  });

Object.defineProperty(module.exports, "__test", {
  enumerable: false,
  value: { buildApprovalMessage, buildApprovalHtml, normalizeRequestId }
});
