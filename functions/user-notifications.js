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
const CALENDAR_ABSENCE_TYPES = new Set(["ferie", "permesso", "malattia"]);

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
      // Firestore document ID is the Auth UID. Profile data.uid is not an
      // authority and could be stale or self-edited in a legacy client.
      uid: doc.id,
      email: String(data.email || ""),
      displayName: String(data.displayName || data.userName || data.email || "Utente"),
      pushToken: String(data.pushToken || "").trim(),
      labels
    };
  }); // Keep profiles without push tokens for ambiguity checks.
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

/**
 * Only deliver a personal push when ONE user can be identified among ALL
 * platform profiles, including profiles without a push token. Prefer exact
 * UID/email/name identity labels; ambiguous exact or fuzzy matches fail closed.
 */
function resolveUniquePushUser(users, rawValue) {
  const value = normalize(rawValue);
  if (!value) return null;
  const exact = users.filter((user) => user.labels.has(value));
  if (exact.length) return exact.length === 1 ? exact[0] : null;
  const fuzzy = users.filter((user) => userMatches(user, rawValue));
  return fuzzy.length === 1 ? fuzzy[0] : null;
}

/**
 * For private chats, never try fuzzy names. A recipientId must be a unique
 * authenticated Firebase UID, or a unique exact legacy email. Unmatched or
 * ambiguous addresses receive no push at all.
 */
function resolvePrivateChatRecipient(users, recipientId) {
  const raw = String(recipientId || "").trim();
  if (!raw) return null;
  const emailValue = raw.includes("@") ? raw.toLowerCase() : "";
  const exact = users.filter((user) =>
    // Firebase Auth UIDs are case-sensitive; never lowercase a recipient UID.
    user.uid === raw ||
    (emailValue && user.email.toLowerCase() === emailValue)
  );
  return exact.length === 1 ? exact[0] : null;
}

// Shared/stale FCM tokens across multiple Firebase profiles are unsafe for
// personal pushes. Do not send until a token is uniquely tied to one profile.
function hasUniquePushTokenOwner(users, user) {
  if (!user || !user.pushToken) return false;
  return users.filter((candidate) => candidate.pushToken === user.pushToken).length === 1;
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
    const senderId = String(message.senderId || "").trim();
    const senderEmail = String(message.senderEmail || "").trim().toLowerCase();
    const recipientId = String(message.recipientId || "").trim();

    const directRecipient = recipientId ? resolvePrivateChatRecipient(users, recipientId) : null;
    const recipients = (recipientId
      ? (hasUniquePushTokenOwner(users, directRecipient) ? [directRecipient] : [])
      : users)
      .filter((user) => {
        const isSender = user.uid === senderId ||
          (senderEmail && user.email.toLowerCase() === senderEmail);
        return !isSender;
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
    const matchedRowsByUser = new Map();
    for (const row of rows) {
      const user = resolveUniquePushUser(users, row.operatorName);
      if (!hasUniquePushTokenOwner(users, user)) continue; // Never guess a person or device.
      if (!matchedRowsByUser.has(user.uid)) matchedRowsByUser.set(user.uid, { user, rows: [] });
      matchedRowsByUser.get(user.uid).rows.push(row);
    }
    for (const { user, rows: matchingRows } of matchedRowsByUser.values()) {
      const total = matchingRows.reduce((sum, row) => sum + row.hours, 0);
      const commesse = [...new Set(matchingRows.map((row) => row.commessaName).filter(Boolean))];
      await sendToUsers([user], {
        title: "⏱️ Ore inserite",
        body: `${String(report.createdByName || report.createdBy || report.authorName || "Varga")} ti ha segnato ${total.toLocaleString("it-IT")} ore${commesse.length ? ` in ${commesse.join(", ")}` : ""} per il ${String(report.date || "giorno indicato")}.`,
        eventType: "hours-inserted",
        data: {
          reportId: event.params.reportId,
          date: report.date || "",
          fullMessage: `${String(report.createdByName || report.createdBy || report.authorName || "Varga")} ti ha segnato ${total.toLocaleString("it-IT")} ore${commesse.length ? ` in ${commesse.join(", ")}` : ""} per il ${String(report.date || "giorno indicato")}.`
        },
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

exports.notifyAndroidSquadraAssignment = onDocumentWritten(
  { document: "squadreCommesse/{commessaId}", region: REGION },
  async (event) => {
    const before = event.data?.before?.exists ? event.data.before.data() || {} : {};
    const after = event.data?.after?.exists ? event.data.after.data() || {} : {};
    if (!event.data?.after?.exists) return null;

    const beforeOperators = extractSquadraOperators(before);
    const afterOperators = extractSquadraOperators(after);
    const added = [...afterOperators].filter((operator) => ![...beforeOperators].some((oldValue) => normalize(oldValue) === normalize(operator)));
    if (!added.length) return null;

    const users = await loadPushUsers();
    const commessaName = String(after.commessaNome || "Commessa");
    const date = String(after.riferimentoData || after.dateKey || "");
    const uniquelyResolved = new Map();
    for (const operator of added) {
      const user = resolveUniquePushUser(users, operator);
      if (hasUniquePushTokenOwner(users, user)) uniquelyResolved.set(user.uid, user);
    }
    for (const user of uniquelyResolved.values()) {
        await sendToUsers([user], {
          title: "👷 Aggiunto a una squadra",
          body: `Sei stato aggiunto alla squadra di ${commessaName}${date ? ` per il ${date}` : ""}.`,
          eventType: "squadra-assigned",
          data: { commessaId: event.params.commessaId, date },
          tag: `squadra-${event.params.commessaId}-${date}-${user.uid}`
        });
    }
    return null;
  }
);

function getCalendarEventParticipants(calendarEvent = {}) {
  const snapshots = Array.isArray(calendarEvent.participantSnapshots)
    ? calendarEvent.participantSnapshots.map((participant) => participant?.name || participant?.email || participant?.id)
    : [];
  const textParticipants = String(calendarEvent.participants || "").split(/[;,\n|]+/);
  return [...new Set([
    ...snapshots,
    ...textParticipants,
    ...(Array.isArray(calendarEvent.participantIds) ? calendarEvent.participantIds : []),
    ...(Array.isArray(calendarEvent.participantEmails) ? calendarEvent.participantEmails : [])
  ].map((value) => String(value || "").trim()).filter(Boolean))];
}

function calendarEventMatchesOperator(calendarEvent, operator) {
  const operatorKey = normalize(operator);
  if (!operatorKey) return false;
  return getCalendarEventParticipants(calendarEvent).some((participant) => {
    const participantKey = normalize(participant);
    return participantKey && (participantKey === operatorKey || participantKey.includes(operatorKey) || operatorKey.includes(participantKey));
  });
}

function enumerateDateKeys(startDate, endDate) {
  const start = new Date(`${String(startDate || "")}T12:00:00`);
  const end = new Date(`${String(endDate || startDate || "")}T12:00:00`);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || end < start) return [];
  const result = [];
  for (const cursor = new Date(start); cursor <= end && result.length < 370; cursor.setDate(cursor.getDate() + 1)) {
    result.push([
      cursor.getFullYear(),
      String(cursor.getMonth() + 1).padStart(2, "0"),
      String(cursor.getDate()).padStart(2, "0")
    ].join("-"));
  }
  return result;
}

function formatAbsenceLabel(calendarEvent = {}) {
  return {
    ferie: "Ferie",
    permesso: "Permesso",
    malattia: "Malattia"
  }[calendarEvent.type] || "Assenza";
}

function buildAutomaticAbsenceText(calendarEvent, operatorNames, dateKey) {
  return [...new Set(operatorNames)].map((name) => `⛔ ${name} assente: ${formatAbsenceLabel(calendarEvent)} il ${dateKey}`).join("\n");
}

async function syncCalendarEventAcrossSquads(eventId, before, after) {
  const beforeDates = enumerateDateKeys(before?.startDate, before?.endDate);
  const afterIsAbsence = Boolean(after && CALENDAR_ABSENCE_TYPES.has(String(after.type || "")));
  const afterDates = afterIsAbsence ? enumerateDateKeys(after.startDate, after.endDate) : [];
  const affectedDates = [...new Set([...beforeDates, ...afterDates])];
  if (!affectedDates.length) return;
  const db = getDb();
  const existingAlerts = await db.collection("userAlerts").where("calendarEventId", "==", eventId).get();
  const deleteBatch = db.batch();
  let deleteCount = 0;
  existingAlerts.forEach((doc) => {
    if (String(doc.data()?.source || "") !== "calendar-absence") return;
    deleteBatch.delete(doc.ref);
    deleteCount += 1;
  });
  if (deleteCount) await deleteBatch.commit();

  for (const dateBatch of chunk(affectedDates, 10)) {
    const snapshot = await db.collection("squadreStorico").where("dateKey", "in", dateBatch).get();
    for (const squadDoc of snapshot.docs) {
      const squadData = squadDoc.data() || {};
      const dateKey = String(squadData.dateKey || squadData.riferimentoData || "");
      const rows = Array.isArray(squadData.squadre) ? squadData.squadre : [];
      let changed = false;
      const matchedRows = [];
      const nextRows = rows.map((row, index) => {
        const previousAlerts = Array.isArray(row.calendarAbsenceAlerts) ? row.calendarAbsenceAlerts : [];
        const nextAlerts = previousAlerts.filter((alert) => String(alert?.eventId || "") !== eventId);
        const operators = [...extractSquadraRowOperators(row)];
        const matchingOperators = afterIsAbsence && afterDates.includes(dateKey)
          ? operators.filter((operator) => calendarEventMatchesOperator(after, operator))
          : [];
        if (matchingOperators.length) {
          const text = buildAutomaticAbsenceText(after, matchingOperators, dateKey);
          nextAlerts.push({
            eventId,
            text,
            operatorNames: matchingOperators,
            type: String(after.type || ""),
            startDate: String(after.startDate || ""),
            endDate: String(after.endDate || "")
          });
          matchedRows.push({ index, text, row, matchingOperators });
        }
        const avvisoAutomaticoAssenze = nextAlerts.map((alert) => String(alert?.text || "").trim()).filter(Boolean).join("\n");
        if (normalize(avvisoAutomaticoAssenze) !== normalize(row.avvisoAutomaticoAssenze || "") || nextAlerts.length !== previousAlerts.length) changed = true;
        return { ...row, calendarAbsenceAlerts: nextAlerts, avvisoAutomaticoAssenze };
      });
      if (!changed) continue;

      await squadDoc.ref.set({
        squadre: nextRows,
        calendarAbsenceUpdatedAt: admin.firestore.FieldValue.serverTimestamp()
      }, { merge: true });

      const commessaId = String(squadData.commessaId || "");
      if (commessaId) {
        const currentRef = db.collection("squadreCommesse").doc(commessaId);
        const currentSnap = await currentRef.get();
        const currentData = currentSnap.exists ? currentSnap.data() || {} : {};
        if (String(currentData.riferimentoData || currentData.dateKey || "") === dateKey) {
          await currentRef.set({
            squadre: nextRows,
            calendarAbsenceUpdatedAt: admin.firestore.FieldValue.serverTimestamp()
          }, { merge: true });
        }
      }

      for (const match of matchedRows) {
        const alertId = `calendar-absence-${eventId}-${squadDoc.id}-${match.index + 1}`.replace(/[^a-zA-Z0-9_-]/g, "_").slice(0, 500);
        await db.collection("userAlerts").doc(alertId).set({
          title: "⛔ Operatore assente nella squadra",
          message: `${squadData.commessaNome || "Commessa"} • Squadra ${match.index + 1} • ${dateKey}\n\n${match.text}`,
          alertText: match.text,
          source: "calendar-absence",
          calendarEventId: eventId,
          commessaId,
          commessaNome: squadData.commessaNome || "Commessa",
          squadraIndex: match.index + 1,
          targetMemberNames: [...extractSquadraRowOperators(match.row)],
          scheduledDateKey: dateKey,
          createdDateKey: new Date().toISOString().slice(0, 10),
          status: "active",
          createdAt: admin.firestore.FieldValue.serverTimestamp(),
          createdBy: "calendar-system"
        }, { merge: true });
      }
    }
  }
}

exports.syncCalendarAbsenceSquadraAlerts = onDocumentWritten(
  { document: "calendarEvents/{eventId}", region: REGION },
  async (event) => {
    const before = event.data?.before?.exists ? event.data.before.data() || {} : null;
    const after = event.data?.after?.exists ? event.data.after.data() || {} : null;
    await syncCalendarEventAcrossSquads(event.params.eventId, before, after);
    return null;
  }
);
