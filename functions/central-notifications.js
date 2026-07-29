"use strict";
const { onCall, HttpsError } = require("firebase-functions/v2/https");
const { getFirestore, FieldValue } = require("firebase-admin/firestore");

const TYPES = new Set(["IMPIANTO_COMPLETATO", "NAVIGAZIONE", "ORE", "SQUADRA", "MEZZO", "SEGNALAZIONE", "INFO"]);
exports.createCentralNotification = onCall({ region: "europe-west1" }, async (request) => {
  if (!request.auth) throw new HttpsError("unauthenticated", "Accesso richiesto.");
  const input = request.data || {};
  if (!TYPES.has(String(input.type || ""))) throw new HttpsError("invalid-argument", "Tipo notifica non consentito.");
  if (["GLOBALE", "ADMIN"].includes(input.scopeType) || input.priority === "CRITICA") throw new HttpsError("permission-denied", "Notifica riservata ai servizi autorizzati.");
  const recipients = [...new Set([request.auth.uid, ...(Array.isArray(input.recipientUserIds) ? input.recipientUserIds : [])])].slice(0, 100);
  const dedupeKey = String(input.dedupeKey || "").slice(0, 300), db = getFirestore();
  const ref = dedupeKey ? db.collection("notifications").doc(Buffer.from(dedupeKey).toString("base64url").slice(0, 140)) : db.collection("notifications").doc();
  const value = {
    id: ref.id, type: String(input.type), priority: String(input.priority || "NORMALE"), title: String(input.title || "Notifica").slice(0, 160),
    preview: String(input.preview || input.message || "").slice(0, 500), message: String(input.message || input.preview || "").slice(0, 4000),
    actorId: request.auth.uid, actorName: String(input.actorName || request.auth.token.name || request.auth.token.email || "Operatore").slice(0, 160),
    scopeType: String(input.scopeType || "PERSONALE"), recipientUserIds: recipients,
    commessaId: String(input.commessaId || ""), commessaName: String(input.commessaName || ""), squadraId: String(input.squadraId || ""), squadraName: String(input.squadraName || ""),
    impiantoId: String(input.impiantoId || ""), impiantoName: String(input.impiantoName || ""), actionType: String(input.actionType || ""), actionTarget: String(input.actionTarget || ""),
    dedupeKey, expiresAt: input.expiresAt || null, readBy: {}, deletedBy: {}, metadata: input.metadata || {}, createdAt: FieldValue.serverTimestamp(), updatedAt: FieldValue.serverTimestamp()
  };
  try { await ref.create(value); } catch (error) { if (error.code !== 6 && error.code !== "already-exists") throw error; }
  return { id: ref.id };
});
