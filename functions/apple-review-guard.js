"use strict";

// Shared server-side gate: Firestore/Storage Rules DO NOT apply to Admin SDK
// operations in Cloud Functions. Check the signed Firebase ID token directly.
const { HttpsError } = require("firebase-functions/v2/https");

function assertNotAppleReviewer(auth) {
  if (auth?.token?.apple_review === true) {
    throw new HttpsError(
      "permission-denied",
      "Un account Apple Review può accedere soltanto ai dati dimostrativi e non può utilizzare i servizi aziendali."
    );
  }
}

module.exports = { assertNotAppleReviewer };
