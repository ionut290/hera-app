"use strict";

const admin = require("firebase-admin");
const { onDocumentCreated, onDocumentWritten } = require("firebase-functions/v2/firestore");

const REGION = "europe-west1";
const CHANNEL_ID = "hera_operational_updates";
const FCM_BATCH_SIZE = 500;
const INVALID_TOKEN_CODES = new Set([
  "messaging/invalid-registration-token",
  "messaging/registration-token-not-registered"
]);

function normalize(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("it-IT")
    .replace(/[^a-z0-9@._+-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function chunk(items, size) {
  const result = [];
  for (let index = 0; index < items.length; index += size) result.push(items.slice(index, index + size));
  return result;
}

function getDb() {
  return admin.firestore();
}

async function loadPushUsers() {
  const snapshot = await getDb().collection("platformUsers").get();
  return snapshot.docs.map((doc) => {
    const data = doc.data() || {};
    const labels = new Set([
      doc.id,
      data.uid,
      data.email,
      data.displayName,
      data.firstName && data.lastName ? `${data.firstName} ${data.lastName}` : "",
      data.userName,
      data.nome
    ].map(normalize).filter(Boolean));
    return {
      ref: doc.ref,
      uid: String(data.uid || doc.id),
      email: String(data.email || ""),
      displayName: String(data.displayName || data.userName || data.email || "Utente"),
      pushToken: String(data.pushToken || "").trim(),
      labels
    };
  }).filter((user) => user.pushToken);
}

function userMatches(user, rawValue) {
  const value = normalize(rawValue);
  if (!value) return false;
  if (user.labels.has(value)) return true;
  for (const label of user.labels) {
    if (label && value && (label.includes(value) || value.includes(label))) return true;
  }
  return false;
}

async function clearInvalidTokens(users, invalidTokens) {
  if (!invalidTokens.size) return;
  const batch = getDb().batch();
  let writes = 0;
  users.forEach((user) => {
    if (!invalidTokens.has(user.pushToken)) return;
    batch.set(user.ref, {
      pushToken: admin.firestore.FieldValue.delete(),
      pushTokenInvalidatedAt: admin.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    writes += 1;
  });
  if (writes) await batch.commit();
}

async function sendToUsers(users, { title, body, eventType, data = {}, tag = "hera-update" }) {
  const uniqueByToken = new Map();
  users.forEach((user) => {
    if (user.pushToken) uniqueByToken.set(user.pushToken, user);
  });
  const recipients = [...uniqueByToken.values()];
  if (!recipients.length) return { successCount: 0, failureCount: 0 };

  const invalidTokens = new Set();
  let successCount = 0;
  let failureCount = 0;

  for (const userBatch of chunk(recipients, FCM_BATCH_SIZE)) {
    const response = await admin.messaging().sendEachForMulticast({
      tokens: userBatch.map((user) => user.pushToken),
      notification: {
        title: String(title || "Varga Cantieri").slice(0, 120),
        body: String(body || "Nuovo aggiornamento disponibile.").slice(0, 500)
      },
      data: Object.fromEntries(Object.entries({ eventType, ...data }).map(([key, value]) => [key, String(value ?? "")])),
      android: {
        priority: "high",
        notification: {
          channelId: CHANNEL_ID,
          sound: "default",
          tag: String(tag || "hera-update").slice(0, 100)
        }
      }
    });
    successCount += response.successCount;
    failureCount += response.failureCount;
    response.responses.forEach((result, index) => {
      if (!result.success && INVALID_TOKEN_CODES.has(result.error?.code)) {
        invalidTokens.add(userBatch[index].pushToken);
      }
    });
  }

  await clearInvalidTokens(recipients, invalidTokens);
  return { successCount, failureCount };
}

function getChatBody(message) {
  const text = String(message.text || "").trim();
  if (text) return text;
  if (message.type === "voice") return "Ti ha inviato un messaggio vocale.";
  if (message.type === "image") return "Ti ha inviato una foto.";
  if (message.type === "video") return "Ti ha inviato un video.";
  return "Ti ha inviato un nuovo messaggio.";
}

exports.notifyAndroidChatMessage = onDocumentCreated(
  { document: "chatMessages/{messageId}", region: REGION },
  async (event) => {
    const snapshot = event.data;
    if (!snapshot) return null;
    const message = snapshot.data() || {};
    const users = await loadPushUsers();
    const senderId = normalize(message.senderId);
    const senderEmail = normalize(message.senderEmail);
    const recipientId = String(message.recipientId || "").trim();

    const recipients = users.filter((user) => {
      const isSender = user.labels.has(senderId) || user.labels.has(senderEmail);
      if (isSender) return false;
      return recipientId ? userMatches(user, recipientId) : true;
    });

    const senderName = String(message.senderName || "Nuovo messaggio").trim();
    const result = await sendToUsers(recipients, {
      title: `💬 Messaggio da ${senderName}`,
      body: getChatBody(message),
      eventType: "chat-message",
      data: {
        messageId: event.params.messageId,
        senderId: message.senderId || "",
        recipientId
      },
      tag: recipientId ? `chat-${event.params.messageId}` : "chat-generale"
    });
    console.info("Notifica chat Android inviata.", { messageId: event.params.messageId, ...result });
    return null;
  }
);

function extractHoursRows(report) {
  const rows = [];
  (Array.isArray(report.entries) ? report.entries : []).forEach((entry) => {
    const commessaName = String(entry.commessaName || entry.commessaNome || entry.nomeCommessa || "Commessa");
    (Array.isArray(entry.rows) ? entry.rows : []).forEach((row) => {
      const operatorName = row.operatore || row.operatorName || row.nomeOperatore || row.email || row.userEmail || row.uid || row.userId;
      const hours = Number(row.ore || row.hours || 0);
      if (operatorName && hours > 0) rows.push({ operatorName: String(operatorName), hours, commessaName });
    });
  });
  return rows;
}

exports.notifyAndroidHoursInserted = onDocumentCreated(
  { document: "oreReports/{reportId}", region: REGION },
  async (event) => {
    const snapshot = event.data;
    if (!snapshot) return null;
    const report = snapshot.data() || {};
    const rows = extractHoursRows(report);
    if (!rows.length) return null;

    const users = await loadPushUsers();
    for (const user of users) {
      const matchingRows = rows.filter((row) => userMatches(user, row.operatorName));
      if (!matchingRows.length) continue;
      const total = matchingRows.reduce((sum, row) => sum + row.hours, 0);
      const commesse = [...new Set(matchingRows.map((row) => row.commessaName).filter(Boolean))];
      await sendToUsers([user], {
        title: "⏱️ Ore inserite",
        body: `${total.toLocaleString("it-IT")} ore inserite per il ${String(report.date || "giorno indicato")}${commesse.length ? ` • ${commesse.slice(0, 2).join(", ")}` : ""}`,
        eventType: "hours-inserted",
        data: { reportId: event.params.reportId, date: report.date || "" },
        tag: `hours-${event.params.reportId}-${user.uid}`
      });
    }
    return null;
  }
);

function extractSquadraOperators(documentData) {
  const values = new Set();
  (Array.isArray(documentData?.squadre) ? documentData.squadre : []).forEach((row) => {
    extractSquadraRowOperators(row).forEach((value) => values.add(value));
  });
  return values;
}

function extractSquadraRowOperators(row = {}) {
  const values = new Set();
  [row.caposquadra, row.caposquadraEmail, row.caposquadraUid].forEach((value) => {
    if (String(value || "").trim()) values.add(String(value).trim());
  });
  const personnel = row.personale || row.operatori || row.operators || [];
  if (Array.isArray(personnel)) {
    personnel.forEach((value) => {
      const label = typeof value === "object"
        ? value.uid || value.email || value.displayName || value.nome
        : value;
      if (String(label || "").trim()) values.add(String(label).trim());
    });
  } else {
    String(personnel || "").split(/[;,\n|]+/).map((value) => value.trim()).filter(Boolean).forEach((value) => values.add(value));
  }
  return values;
}

function extractChangedSquadraAlerts(before, after) {
  const beforeRows = Array.isArray(before?.squadre) ? before.squadre : [];
  const afterRows = Array.isArray(after?.squadre) ? after.squadre : [];
  const sameDate = String(before?.riferimentoData || before?.dateKey || "") === String(after?.riferimentoData || after?.dateKey || "");
  return afterRows.map((row, index) => {
    const alertText = String(row?.avviso || "").trim();
    const previousAlertText = String(beforeRows[index]?.avviso || "").trim();
    if (!alertText || (sameDate && normalize(alertText) === normalize(previousAlertText))) return null;
    return {
      index,
      alertText,
      operators: [...extractSquadraRowOperators(row)]
    };
  }).filter(Boolean);
}

exports.notifyAndroidSquadraAssignment = onDocumentWritten(
  { document: "squadreCommesse/{commessaId}", region: REGION },
  async (event) => {
    const before = event.data?.before?.exists ? event.data.before.data() || {} : {};
    const after = event.data?.after?.exists ? event.data.after.data() || {} : {};
    if (!event.data?.after?.exists) return null;

    const beforeOperators = extractSquadraOperators(before);
    const afterOperators = extractSquadraOperators(after);
    const added = [...afterOperators].filter((operator) => ![...beforeOperators].some((oldValue) => normalize(oldValue) === normalize(operator)));
    const changedAlerts = extractChangedSquadraAlerts(before, after);
    if (!added.length && !changedAlerts.length) return null;

    const users = await loadPushUsers();
    const commessaName = String(after.commessaNome || "Commessa");
    const date = String(after.riferimentoData || after.dateKey || "");
    for (const user of users) {
      if (added.some((operator) => userMatches(user, operator))) {
        await sendToUsers([user], {
          title: "👷 Aggiunto a una squadra",
          body: `Sei stato aggiunto alla squadra di ${commessaName}${date ? ` per il ${date}` : ""}.`,
          eventType: "squadra-assigned",
          data: { commessaId: event.params.commessaId, date },
          tag: `squadra-${event.params.commessaId}-${date}-${user.uid}`
        });
      }
      for (const squadAlert of changedAlerts) {
        if (!squadAlert.operators.some((operator) => userMatches(user, operator))) continue;
        await sendToUsers([user], {
          title: "⚠️ Avviso squadra",
          body: `${commessaName}${date ? ` • ${date}` : ""}: ${squadAlert.alertText}`,
          eventType: "squadra-alert",
          data: {
            commessaId: event.params.commessaId,
            date,
            squadraIndex: squadAlert.index + 1,
            alertText: squadAlert.alertText
          },
          tag: `squadra-alert-${event.params.commessaId}-${date}-${squadAlert.index + 1}-${user.uid}`
        });
      }
    }
    return null;
  }
);
