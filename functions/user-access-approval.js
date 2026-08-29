"use strict";

const crypto = require("node:crypto");
const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");
const { defineJsonSecret } = require("firebase-functions/params");

const REGION = "europe-west1";
const BUILT_IN_ADMIN_EMAIL = "ionut29019@gmail.com";
const APP_NAME = "Varga Cantieri";
const APP_URL = "https://creative-syrniki-dddbae.netlify.app";
const APP_LOGO_URL = `${APP_URL}/icons/varga-cantieri-192.png`;
const ANDROID_URL = "https://play.google.com/store/apps/details?id=it.vargacantieri.hera";
const RESEND_ENDPOINT = "https://api.resend.com/emails";
const RUNTIME_CONFIG = defineJsonSecret("RUNTIME_CONFIG");

function optionalEmailConfiguration() {
  const runtimeConfig = RUNTIME_CONFIG.value() || {};
  const resendConfig = runtimeConfig.resend || {};
  return {
    apiKey: String(resendConfig.api_key || "").trim(),
    from: String(resendConfig.from || "").trim()
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
    `🌿 BENVENUTO IN ${APP_NAME.toUpperCase()}`,
    "✅ La tua richiesta di accesso è stata approvata",
    "",
    `Ciao ${userName || ""},`,
    "",
    `ti diamo il benvenuto in ${APP_NAME}!`,
    `L’amministratore ${administratorName} ha accettato la tua richiesta e ora puoi accedere all’app.`,
    "",
    "COME ACCEDERE",
    `Email: ${userEmail}`,
    "Password: usa la password che hai scelto durante la registrazione.",
    "Per la tua sicurezza, non comunicare la password ad altre persone.",
    "",
    "A COSA SERVE L’APP",
    "• consultare le commesse e gli impianti di lavoro;",
    "• vedere squadre, attività e informazioni operative;",
    "• aprire la navigazione verso gli impianti;",
    "• consultare documenti, comunicazioni e aggiornamenti autorizzati.",
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
    "Siamo felici di averti con noi. Benvenuto e buon lavoro!",
    "",
    `${administratorName}`,
    `Amministratore ${APP_NAME}`
  ].join("\n");
}

function buildApprovalHtml(message, { userName, userEmail, administratorName } = {}) {
  const safeName = escapeHtml(cleanText(userName, 180) || "nuovo utente");
  const safeEmail = escapeHtml(cleanText(userEmail, 254));
  const safeAdministrator = escapeHtml(cleanText(administratorName, 180) || "Amministratore");
  void message;

  return `<!doctype html>
<html lang="it">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <meta name="color-scheme" content="light">
  <title>Benvenuto in ${APP_NAME}</title>
  <style>
    @media only screen and (max-width:620px) {
      .email-shell { width:100% !important; }
      .email-pad { padding-left:22px !important; padding-right:22px !important; }
      .install-column { display:block !important; width:100% !important; }
      .install-spacer { display:block !important; height:12px !important; width:100% !important; }
      .email-title { font-size:29px !important; line-height:35px !important; }
      .cta-button { display:block !important; }
    }
  </style>
</head>
<body style="margin:0;padding:0;background:#edf4ef;color:#173022;font-family:Arial,Helvetica,sans-serif;-webkit-text-size-adjust:100%;">
  <div style="display:none;max-height:0;overflow:hidden;opacity:0;color:transparent;">Il tuo accesso a ${APP_NAME} è stato approvato. Ti diamo il benvenuto!</div>
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;background:#edf4ef;">
    <tr>
      <td align="center" style="padding:28px 12px;">
        <table role="presentation" class="email-shell" width="620" cellspacing="0" cellpadding="0" border="0" style="width:620px;max-width:620px;background:#ffffff;border-radius:24px;overflow:hidden;box-shadow:0 14px 45px rgba(22,73,43,.14);">
          <tr>
            <td class="email-pad" align="center" style="padding:38px 46px 34px;background:#146c3b;background-image:linear-gradient(135deg,#0f6b3a 0%,#259653 100%);">
              <div style="display:inline-block;padding:10px;background:#ffffff;border-radius:24px;box-shadow:0 8px 24px rgba(0,0,0,.16);">
                <img src="${APP_LOGO_URL}" width="88" height="88" alt="Logo ${APP_NAME}" style="display:block;width:88px;height:88px;border:0;border-radius:16px;object-fit:cover;">
              </div>
              <p style="margin:22px 0 10px;color:#d9fbe7;font-size:13px;font-weight:800;letter-spacing:1.7px;text-transform:uppercase;">Accesso approvato</p>
              <h1 class="email-title" style="margin:0;color:#ffffff;font-size:34px;line-height:40px;font-weight:800;">Ti diamo il benvenuto!</h1>
              <p style="margin:12px 0 0;color:#e8fff0;font-size:17px;line-height:26px;">Ora fai parte di <strong>${APP_NAME}</strong></p>
            </td>
          </tr>
          <tr>
            <td class="email-pad" style="padding:38px 46px 12px;">
              <p style="margin:0 0 14px;font-size:21px;line-height:30px;font-weight:800;color:#153c27;">Ciao ${safeName},</p>
              <p style="margin:0;color:#476053;font-size:16px;line-height:26px;">La tua richiesta è stata accettata da <strong style="color:#173c28;">${safeAdministrator}</strong>. Il tuo account è attivo e puoi iniziare subito a usare l’app.</p>
            </td>
          </tr>
          <tr>
            <td class="email-pad" align="center" style="padding:22px 46px 30px;">
              <a class="cta-button" href="${APP_URL}" target="_blank" style="display:inline-block;background:#178447;color:#ffffff;text-decoration:none;font-size:17px;font-weight:800;line-height:20px;padding:16px 30px;border-radius:12px;box-shadow:0 7px 18px rgba(23,132,71,.24);">APRI VARGA CANTIERI</a>
              <p style="margin:12px 0 0;color:#718278;font-size:12px;line-height:18px;">Il pulsante apre l’accesso web all’app</p>
            </td>
          </tr>
          <tr>
            <td class="email-pad" style="padding:0 46px 30px;">
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;background:#f1f8f3;border:1px solid #d5eadb;border-radius:16px;">
                <tr>
                  <td style="padding:21px 22px;">
                    <p style="margin:0 0 11px;color:#146c3b;font-size:13px;font-weight:800;letter-spacing:1px;text-transform:uppercase;">I tuoi dati di accesso</p>
                    <p style="margin:0 0 8px;color:#173c28;font-size:15px;line-height:23px;"><strong>Email:</strong> ${safeEmail}</p>
                    <p style="margin:0;color:#526a5d;font-size:14px;line-height:22px;"><strong>Password:</strong> usa quella che hai scelto durante la registrazione.</p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td class="email-pad" style="padding:0 46px 30px;">
              <h2 style="margin:0 0 16px;color:#153c27;font-size:20px;line-height:28px;">Con l’app puoi</h2>
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
                <tr><td width="34" valign="top" style="padding:0 0 12px;font-size:21px;">📋</td><td valign="top" style="padding:1px 0 12px;color:#4d6357;font-size:15px;line-height:23px;">Consultare commesse e impianti di lavoro</td></tr>
                <tr><td width="34" valign="top" style="padding:0 0 12px;font-size:21px;">👷</td><td valign="top" style="padding:1px 0 12px;color:#4d6357;font-size:15px;line-height:23px;">Vedere squadre, attività e informazioni operative</td></tr>
                <tr><td width="34" valign="top" style="padding:0 0 12px;font-size:21px;">🧭</td><td valign="top" style="padding:1px 0 12px;color:#4d6357;font-size:15px;line-height:23px;">Avviare la navigazione verso gli impianti</td></tr>
                <tr><td width="34" valign="top" style="padding:0;font-size:21px;">📄</td><td valign="top" style="padding:1px 0 0;color:#4d6357;font-size:15px;line-height:23px;">Consultare documenti e comunicazioni autorizzate</td></tr>
              </table>
            </td>
          </tr>
          <tr>
            <td class="email-pad" style="padding:0 46px 32px;">
              <h2 style="margin:0 0 16px;color:#153c27;font-size:20px;line-height:28px;">Installa l’app sul telefono</h2>
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
                <tr>
                  <td class="install-column" width="49%" valign="top" style="width:49%;background:#f7f9f8;border:1px solid #e1e9e4;border-radius:14px;padding:18px;">
                    <p style="margin:0 0 8px;color:#173c28;font-size:16px;font-weight:800;">🤖 Android</p>
                    <p style="margin:0 0 15px;color:#5b6e63;font-size:13px;line-height:20px;">Apri Google Play e premi <strong>Installa</strong>.</p>
                    <a href="${ANDROID_URL}" target="_blank" style="color:#13703d;font-size:14px;font-weight:800;text-decoration:underline;">Apri Google Play →</a>
                  </td>
                  <td class="install-spacer" width="2%" style="width:2%;font-size:1px;">&nbsp;</td>
                  <td class="install-column" width="49%" valign="top" style="width:49%;background:#f7f9f8;border:1px solid #e1e9e4;border-radius:14px;padding:18px;">
                    <p style="margin:0 0 8px;color:#173c28;font-size:16px;font-weight:800;">🍎 iPhone</p>
                    <p style="margin:0 0 15px;color:#5b6e63;font-size:13px;line-height:20px;">Apri con Safari, premi <strong>Condividi</strong> e poi <strong>Aggiungi alla schermata Home</strong>.</p>
                    <a href="${APP_URL}" target="_blank" style="color:#13703d;font-size:14px;font-weight:800;text-decoration:underline;">Apri con Safari →</a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td class="email-pad" style="padding:0 46px 34px;">
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;background:#fff8e8;border:1px solid #f2dfae;border-radius:14px;">
                <tr><td width="42" valign="top" style="padding:17px 0 17px 18px;font-size:21px;">🔐</td><td style="padding:17px 18px 17px 8px;color:#6e5725;font-size:13px;line-height:20px;"><strong>Sicurezza:</strong> non comunicare mai la tua password ad altre persone.</td></tr>
              </table>
            </td>
          </tr>
          <tr>
            <td class="email-pad" style="padding:27px 46px;background:#f5f8f6;border-top:1px solid #e4ece7;">
              <p style="margin:0 0 7px;color:#173c28;font-size:16px;line-height:24px;font-weight:800;">Siamo felici di averti con noi.</p>
              <p style="margin:0 0 20px;color:#587064;font-size:14px;line-height:22px;">Benvenuto e buon lavoro!</p>
              <p style="margin:0;color:#173c28;font-size:14px;line-height:21px;"><strong>${safeAdministrator}</strong><br><span style="color:#718278;">Amministratore ${APP_NAME}</span></p>
            </td>
          </tr>
        </table>
        <p style="margin:18px 0 0;color:#75877c;font-size:11px;line-height:17px;">Hai ricevuto questa e-mail perché è stato approvato un account associato a ${safeEmail || "questo indirizzo"}.</p>
      </td>
    </tr>
  </table>
</body>
</html>`;
}

async function sendApprovalEmail({ apiKey, from, to, message, requestId, userName, administratorName }) {
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
      subject: `Benvenuto in ${APP_NAME} — accesso approvato`,
      text: message,
      html: buildApprovalHtml(message, {
        userName,
        userEmail: to,
        administratorName
      })
    })
  });
  if (!response.ok) {
    const detail = cleanText(await response.text().catch(() => ""), 500);
    throw new Error(`Servizio email ${response.status}${detail ? `: ${detail}` : ""}`);
  }
}

exports.approveUserAccess = functions
  .region(REGION)
  .runWith({ secrets: [RUNTIME_CONFIG] })
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
        await sendApprovalEmail({
          apiKey,
          from,
          to: userEmail,
          message,
          requestId,
          userName,
          administratorName: administrator.displayName
        });
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
