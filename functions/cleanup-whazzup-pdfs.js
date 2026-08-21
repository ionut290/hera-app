"use strict";

const { onSchedule } = require("firebase-functions/v2/scheduler");
const { getFirestore, Timestamp } = require("firebase-admin/firestore");
const { getStorage } = require("firebase-admin/storage");

const SOURCE = "whazzup-impianto-pdf";
const BATCH_LIMIT = 200;

exports.cleanupExpiredWhazzupPdfs = onSchedule(
  {
    schedule: "every day 03:20",
    timeZone: "Europe/Rome",
    region: "europe-west1",
    retryCount: 1
  },
  async () => {
    const db = getFirestore();
    const bucket = getStorage().bucket();
    const now = Timestamp.now();
    let deleted = 0;
    let scanned = 0;

    while (true) {
      const snapshot = await db.collection("documents")
        .where("expiresAt", "<=", now)
        .limit(BATCH_LIMIT)
        .get();

      if (snapshot.empty) break;
      const expiredWhazzup = snapshot.docs.filter((doc) => doc.get("source") === SOURCE);
      scanned += snapshot.size;

      for (const doc of expiredWhazzup) {
        const data = doc.data() || {};
        const storagePath = String(data.storagePath || "").trim();
        if (storagePath) {
          await bucket.file(storagePath).delete({ ignoreNotFound: true }).catch((error) => {
            console.warn("PDF Whazzup: file Storage non eliminato", { documentId: doc.id, storagePath, error: error?.message || error });
          });
        }
        await doc.ref.delete();
        deleted += 1;
      }

      // Se il batch contiene documenti scaduti di altri moduli, non li tocchiamo.
      // Per evitare un loop sullo stesso batch, usciamo: il cleanup client dei PDF
      // mantiene comunque nascosti i record oltre la scadenza, e il job successivo
      // riproverà dopo eventuali eliminazioni dei documenti di altri moduli.
      if (expiredWhazzup.length < snapshot.size || snapshot.size < BATCH_LIMIT) break;
    }

    console.log("Cleanup PDF Whazzup completato", { scanned, deleted, source: SOURCE });
  }
);
