"use strict";

const admin = require("firebase-admin");

const CHANNEL_ID = "hera_operational_updates";
const FCM_BATCH_SIZE = 500;
const CALENDAR_ABSENCE_TYPES = new Set(["ferie", "permesso", "malattia"]);
const BUILT_IN_ADMIN_EMAILS = new Set(["ionut29019@gmail.com"]);
const INVALID_TOKEN_CODES = new Set([
  "messaging/invalid-registration-token",
  "messaging/registration-token-not-registered"
]);

admin.initializeApp({
  credential: admin.credential.applicationDefault()
});

function chunk(items, size) {
  const result = [];
  for (let index = 0; index < items.length; index += size) result.push(items.slice(index, index + size));
  return result;
}

function getRomeDateParts(date = new Date()) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/Rome",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    hourCycle: "h23"
  }).formatToParts(date);
  return Object.fromEntries(parts.map(({ type, value }) => [type, value]));
}

function getTomorrowRomeDateKey(date = new Date()) {
  const parts = getRomeDateParts(date);
  const currentDay = new Date(Date.UTC(Number(parts.year), Number(parts.month) - 1, Number(parts.day)));
  currentDay.setUTCDate(currentDay.getUTCDate() + 1);
  return currentDay.toISOString().slice(0, 10);
}

function getCalendarEventParticipants(calendarEvent = {}) {
  const snapshots = Array.isArray(calendarEvent.participantSnapshots) ? calendarEvent.participantSnapshots : [];
  const snapshotNames = snapshots.flatMap((participant) => [participant?.id, participant?.uid, participant?.name, participant?.email]);
  const legacyNames = Array.isArray(calendarEvent.participantNames)
    ? calendarEvent.participantNames
    : String(calendarEvent.participants || "").split(/[,;\n]+/);
  return [...new Set([
    ...snapshotNames,
    ...(Array.isArray(calendarEvent.participantIds) ? calendarEvent.participantIds : []),
    ...(Array.isArray(calendarEvent.participantEmails) ? calendarEvent.participantEmails : []),
    ...legacyNames
  ].map((value) => String(value || "").trim()).filter(Boolean))];
}

function normalizePerson(value) {
  return String(value || "").trim().toLocaleLowerCase("it-IT").replace(/\s+/g, " ");
}

function getSquadraOperators(squadData = {}) {
  const operators = [];
  for (const row of Array.isArray(squadData.squadre) ? squadData.squadre : []) {
    operators.push(row?.caposquadra, row?.caposquadraEmail, row?.caposquadraUid);
    const personnel = row?.personale || row?.operatori || row?.operators || [];
    if (Array.isArray(personnel)) {
      for (const person of personnel) {
        if (person && typeof person === "object") {
          operators.push(person.uid, person.id, person.email, person.displayName, person.nome);
        } else {
          operators.push(person);
        }
      }
    } else {
      operators.push(...String(personnel || "").split(/[;,\n|]+/));
    }
  }
  return [...new Set(operators.map(normalizePerson).filter(Boolean))];
}

function absenceEventHasSquadraAssignment(calendarEvent, squadDocuments) {
  if (!CALENDAR_ABSENCE_TYPES.has(String(calendarEvent?.type || "").trim().toLowerCase())) return true;
  const participants = getCalendarEventParticipants(calendarEvent).map(normalizePerson).filter(Boolean);
  if (!participants.length) return false;
  return squadDocuments.some((squadData) => {
    const operators = getSquadraOperators(squadData);
    return participants.some((participant) => operators.some((operator) => (
      participant === operator || participant.includes(operator) || operator.includes(participant)
    )));
  });
}

function formatEventType(calendarEvent = {}) {
  const labels = {
    ferie: "Ferie",
    permesso: "Permesso",
    malattia: "Malattia",
    intervento: "Intervento",
    scadenza: "Scadenza",
    altro: "Evento"
  };
  return labels[String(calendarEvent.type || "").trim().toLowerCase()] || "Evento calendario";
}

async function loadAdminUsers(db) {
  const [usersSnapshot, configSnapshot] = await Promise.all([
    db.collection("platformUsers").get(),
    db.collection("appConfig").doc("adminUsers").get()
  ]);
  const configuredEmails = configSnapshot.exists && Array.isArray(configSnapshot.data()?.emails)
    ? configSnapshot.data().emails
    : [];
  const adminEmails = new Set([
    ...BUILT_IN_ADMIN_EMAILS,
    ...configuredEmails.map((email) => String(email || "").trim().toLowerCase())
  ]);
  return usersSnapshot.docs.map((doc) => {
    const data = doc.data() || {};
    return {
      ref: doc.ref,
      email: String(data.email || "").trim().toLowerCase(),
      pushToken: String(data.pushToken || "").trim()
    };
  }).filter((user) => adminEmails.has(user.email));
}

async function clearInvalidTokens(db, users, invalidTokens) {
  const invalidUsers = users.filter((user) => invalidTokens.has(user.pushToken));
  if (!invalidUsers.length) return;
  const batch = db.batch();
  invalidUsers.forEach((user) => batch.set(user.ref, {
    pushToken: admin.firestore.FieldValue.delete(),
    pushTokenInvalidatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true }));
  await batch.commit();
}

async function sendPushToAdmins(db, admins, calendarEvent, eventId, tomorrowKey, body) {
  const recipients = [...new Map(
    admins.filter((user) => user.pushToken).map((user) => [user.pushToken, user])
  ).values()];
  const invalidTokens = new Set();
  let successCount = 0;
  let failureCount = 0;

  for (const userBatch of chunk(recipients, FCM_BATCH_SIZE)) {
    const response = await admin.messaging().sendEachForMulticast({
      tokens: userBatch.map((user) => user.pushToken),
      notification: {
        title: "📅 Evento in calendario domani",
        body: body.slice(0, 500)
      },
      data: {
        eventType: "calendar-admin-reminder",
        eventId,
        startDate: tomorrowKey
      },
      android: {
        priority: "high",
        notification: {
          channelId: CHANNEL_ID,
          sound: "default",
          tag: `calendar-admin-${eventId}-${tomorrowKey}`
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

  await clearInvalidTokens(db, admins, invalidTokens);
  return { successCount, failureCount, title: calendarEvent.title || formatEventType(calendarEvent) };
}

async function run() {
  const now = new Date();
  const romeParts = getRomeDateParts(now);
  const isManualRun = process.env.GITHUB_EVENT_NAME === "workflow_dispatch";
  const isDryRun = String(process.env.DRY_RUN || "").toLowerCase() === "true";
  if (!isManualRun && romeParts.hour !== "07") {
    console.log(`Esecuzione ignorata: in Italia sono le ${romeParts.hour}, non le 07.`);
    return;
  }

  const tomorrowKey = getTomorrowRomeDateKey(now);
  const db = admin.firestore();
  const events = await db.collection("calendarEvents").where("startDate", "==", tomorrowKey).get();
  if (events.empty) {
    console.log(`Nessun evento da notificare per ${tomorrowKey}.`);
    return;
  }
  if (isDryRun) {
    console.log(`Verifica completata: trovati ${events.size} eventi per ${tomorrowKey}; nessuna notifica inviata.`);
    return;
  }
  const admins = await loadAdminUsers(db);
  const hasAbsenceEvents = events.docs.some((eventDoc) => CALENDAR_ABSENCE_TYPES.has(String(eventDoc.data()?.type || "").trim().toLowerCase()));
  const squadDocuments = hasAbsenceEvents
    ? (await db.collection("squadreStorico").where("dateKey", "==", tomorrowKey).get()).docs.map((doc) => doc.data() || {})
    : [];
  let notifiedEvents = 0;

  for (const eventDoc of events.docs) {
    const calendarEvent = eventDoc.data() || {};
    if (String(calendarEvent.adminReminderDateKey || "") === tomorrowKey) continue;
    if (!absenceEventHasSquadraAssignment(calendarEvent, squadDocuments)) {
      console.log(`Promemoria assenza ignorato per ${eventDoc.id}: nessun partecipante assegnato a una squadra il ${tomorrowKey}.`);
      continue;
    }
    const people = getCalendarEventParticipants(calendarEvent).slice(0, 4).join(", ");
    const body = [
      String(calendarEvent.title || formatEventType(calendarEvent)),
      people,
      calendarEvent.commessaName || calendarEvent.worksite || "",
      calendarEvent.impiantoName || "",
      calendarEvent.location || ""
    ].filter(Boolean).join(" • ");
    const pushResult = await sendPushToAdmins(db, admins, calendarEvent, eventDoc.id, tomorrowKey, body);

    await db.collection("userAlerts").doc(`calendar-admin-${eventDoc.id}-${tomorrowKey}`).set({
      title: "📅 Evento in calendario domani",
      message: body,
      source: "calendar-admin-reminder",
      calendarEventId: eventDoc.id,
      scheduledDateKey: tomorrowKey,
      sendToAdmins: true,
      status: "active",
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
      createdBy: "calendar-system"
    }, { merge: true });
    await eventDoc.ref.set({
      adminReminderDateKey: tomorrowKey,
      adminReminderSentAt: admin.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    notifiedEvents += 1;
    console.log(`Promemoria creato per ${eventDoc.id}: ${pushResult.successCount} push inviati, ${pushResult.failureCount} falliti.`);
  }

  console.log(`Promemoria completato: ${notifiedEvents} eventi elaborati per ${tomorrowKey}.`);
}

if (require.main === module) {
  run().then(() => process.exit(0)).catch((error) => {
    console.error("Promemoria calendario fallito:", error);
    process.exit(1);
  });
}

module.exports = {
  getRomeDateParts,
  getTomorrowRomeDateKey,
  absenceEventHasSquadraAssignment,
  run
};
