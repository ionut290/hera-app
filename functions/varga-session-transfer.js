"use strict";

const functions = require("firebase-functions/v1");
const admin = require("firebase-admin");

if (!admin.apps.length) admin.initializeApp();

exports.createVargaCantieriTransferToken = functions
  .region("europe-west1")
  .runWith({ timeoutSeconds: 15, memory: "256MB" })
  .https.onCall(async (_data, context) => {
    if (!context.auth?.uid) {
      throw new functions.https.HttpsError("unauthenticated", "Accesso al Gestionale richiesto.");
    }
    const token = await admin.auth().createCustomToken(context.auth.uid, { vargaSessionTransfer: true });
    return { token };
  });
