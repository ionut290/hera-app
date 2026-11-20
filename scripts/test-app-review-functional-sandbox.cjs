"use strict";

// This is NOT an Apple reviewer account or a full iPhone test.
// Exercise real Firestore security rules with fully fictitious records in a
// LOCAL emulator only. No Google credentials, Firebase Admin, or production APIs.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const { doc, collection, getDoc, getDocs, setDoc, updateDoc, deleteDoc } =
  require("firebase/firestore");
const { initializeTestEnvironment, assertSucceeds, assertFails } =
  require("@firebase/rules-unit-testing");

const PROJECT = "demo-apple-functional-sandbox";
const OPERATOR = "demoOperator";
const OTHER = "demoColleague";
const ADMIN = "demoAdministrator";

async function main() {
  assert.equal(process.env.GOOGLE_CLOUD_PROJECT, PROJECT,
    "Refuse to run outside the dedicated synthetic emulator project");
  const host = process.env.FIRESTORE_EMULATOR_HOST || "";
  assert.match(host, /^(localhost|127\.0\.0\.1):[0-9]+$/,
    "Refuse any non-local Firestore endpoint");
  const [hostname, port] = host.split(":");
  const env = await initializeTestEnvironment({
    projectId: PROJECT,
    firestore: {
      host: hostname, port: Number(port),
      rules: fs.readFileSync("firestore.rules", "utf8")
    }
  });

  try {
    await env.withSecurityRulesDisabled(async (ctx) => {
      const db = ctx.firestore();
      const entries = [
        ["platformUsers/" + OPERATOR, {
          email: "operator@example.test", statoAccount: "attivo",
          accountStatus: "attivo", role: "user", banned: false
        }],
        ["platformUsers/" + OTHER, {
          email: "colleague@example.test", statoAccount: "attivo",
          role: "user", banned: false
        }],
        ["platformUsers/demoPending", {
          email: "pending@example.test", statoAccount: "in_attesa",
          role: "user", banned: false
        }],
        ["appConfig/adminUsers", { emails: ["demo-admin@example.test"] }],
        ["commesse/demoJobAlpha", {
          name: "Cantiere immaginario A", synthetic: true
        }],
        ["commesse/demoJobBeta", {
          name: "Cantiere immaginario B", synthetic: true
        }],
        ["commesse/demoJobAlpha/impianti/demoPlantOne", {
          name: "Impianto fittizio 1", lat: 0, lng: 0, synthetic: true
        }],
        ["commesse/demoJobAlpha/impianti/demoPlantTwo", {
          name: "Impianto fittizio 2", lat: 0, lng: 0, synthetic: true
        }],
        ["commesse/demoJobBeta/impianti/demoPlantThree", {
          name: "Impianto fittizio 3", lat: 0, lng: 0, synthetic: true
        }],
        ["squadre/demoTeam", {
          name: "Squadra immaginaria", members: [OPERATOR, OTHER], synthetic: true
        }],
        ["personale/demoPersonOne", {
          displayName: "Operatore Fittizio Uno", synthetic: true
        }],
        ["personale/demoPersonTwo", {
          displayName: "Operatore Fittizio Due", synthetic: true
        }],
        ["mezzi/demoVehicle", { name: "Mezzo fittizio", synthetic: true }],
        ["oreReports/demoHours", {
          userId: OPERATOR, date: "2026-01-01", hours: 4, synthetic: true
        }],
        ["calendarEvents/demoEvent", {
          createdByUid: OPERATOR, type: "intervento", title: "Evento di prova",
          startDate: "2026-01-01", endDate: "2026-01-01", synthetic: true
        }],
        ["documents/demoSharedPdf", {
          id: "demoSharedPdf", visibility: "global", ownerUserId: OPERATOR,
          fileUrl: "demo:pdf-only-no-file-upload", synthetic: true
        }],
        ["globalCommesse/demoGlobal", {
          name: "Commessa globale fittizia", synthetic: true
        }],
        ["preventiviPriceLists/demoPrice", {
          description: "Prezzo fittizio", amount: 25, synthetic: true
        }],
        ["notifications/demoGlobalNotice", {
          scopeType: "GLOBALE", recipientUserIds: [], synthetic: true
        }],
        ["appleReviewDemo/demoSummary", {
          title: "Demo statica isolata", synthetic: true
        }]
      ];
      for (const [path, data] of entries) await setDoc(doc(db, path), data);
    });

    const operator = env.authenticatedContext(OPERATOR, {
      email: "operator@example.test", email_verified: true
    }).firestore();
    const admin = env.authenticatedContext(ADMIN, {
      email: "demo-admin@example.test", email_verified: true
    }).firestore();
    const reviewer = env.authenticatedContext(OPERATOR, {
      email: "operator@example.test", email_verified: true,
      apple_review: true, appReviewDemo: true
    }).firestore();
    const pending = env.authenticatedContext("demoPending", {
      email: "pending@example.test", email_verified: true
    }).firestore();
    const anonymous = env.unauthenticatedContext().firestore();

    // Synthetic access to the same module collections used by the real app.
    const ordinaryReadable = [
      "commesse/demoJobAlpha", "commesse/demoJobBeta",
      "commesse/demoJobAlpha/impianti/demoPlantOne",
      "commesse/demoJobAlpha/impianti/demoPlantTwo",
      "commesse/demoJobBeta/impianti/demoPlantThree",
      "squadre/demoTeam", "personale/demoPersonOne",
      "personale/demoPersonTwo", "mezzi/demoVehicle",
      "oreReports/demoHours", "calendarEvents/demoEvent",
      "documents/demoSharedPdf", "globalCommesse/demoGlobal",
      "notifications/demoGlobalNotice"
    ];
    for (const path of ordinaryReadable) {
      await assertSucceeds(getDoc(doc(operator, path)));
      await assertFails(getDoc(doc(reviewer, path)));
      await assertFails(getDoc(doc(pending, path)));
      await assertFails(getDoc(doc(anonymous, path)));
    }
    const jobs = await assertSucceeds(getDocs(collection(operator, "commesse")));
    assert.equal(jobs.size, 2, "Both fictitious commesse must appear in listings");

    // Exercise reversible writes on ONLY local fixture records.
    await assertSucceeds(setDoc(doc(operator, "oreReports/demoNewHours"), {
      userId: OPERATOR, date: "2026-01-02", hours: 2
    }));
    await assertSucceeds(updateDoc(doc(operator, "oreReports/demoNewHours"), {
      hours: 3
    }));
    assert.equal((await getDoc(doc(operator, "oreReports/demoNewHours"))).data().hours, 3);
    await assertSucceeds(deleteDoc(doc(operator, "oreReports/demoNewHours")));

    await assertSucceeds(setDoc(doc(operator, "calendarEvents/demoNewEvent"), {
      createdByUid: OPERATOR, startDate: "2026-01-02",
      endDate: "2026-01-02", title: "Prova calendario", type: "intervento"
    }));
    await assertSucceeds(updateDoc(doc(operator, "calendarEvents/demoNewEvent"), {
      title: "Calendario verificato"
    }));
    await assertSucceeds(deleteDoc(doc(operator, "calendarEvents/demoNewEvent")));

    await assertSucceeds(setDoc(doc(operator, "commesse/demoJobAlpha/impianti/demoNewPlant"), {
      name: "Impianto SOLO emulator", lat: 0, lng: 0
    }));
    await assertSucceeds(updateDoc(doc(operator, "commesse/demoJobAlpha/impianti/demoNewPlant"), {
      name: "Impianto aggiornato SOLO emulator"
    }));
    await assertSucceeds(deleteDoc(doc(operator, "commesse/demoJobAlpha/impianti/demoNewPlant")));

    const fakeDoc = {
      id: "demoPersonalPdf", ownerUserId: OPERATOR, createdBy: OPERATOR,
      visibility: "personal", sharedToAll: false, sharedUserIds: [],
      commessaIds: [], uploadStatus: "completed",
      fileUrl: "demo:no-real-storage-object",
      storagePath: "documents/demoOperator/demoPersonalPdf/fake.pdf",
      mimeType: "application/pdf", fileSize: 1024,
      versionHistoryEnabled: false
    };
    await assertSucceeds(setDoc(doc(operator, "documents/demoPersonalPdf"), fakeDoc));
    await assertSucceeds(getDoc(doc(operator, "documents/demoPersonalPdf")));
    await assertFails(getDoc(doc(env.authenticatedContext(OTHER, {
      email: "colleague@example.test", email_verified: true
    }).firestore(), "documents/demoPersonalPdf")));
    await assertSucceeds(deleteDoc(doc(operator, "documents/demoPersonalPdf")));

    await assertFails(setDoc(doc(operator, "commesse/forbiddenTopLevel"), {
      name: "Must fail"
    }));
    await assertSucceeds(setDoc(doc(admin, "commesse/demoAdminCreated"), {
      name: "Admin-only fixture"
    }));
    await assertSucceeds(deleteDoc(doc(admin, "commesse/demoAdminCreated")));
    await assertFails(getDoc(doc(operator, "preventiviPriceLists/demoPrice")));
    await assertSucceeds(getDoc(doc(admin, "preventiviPriceLists/demoPrice")));
    await assertFails(getDoc(doc(operator, "appConfig/adminUsers")));
    await assertFails(getDoc(doc(operator, "platformUsers/" + OTHER)));

    // The special reviewer claim must not silently unlock production-style data.
    await assertSucceeds(getDoc(doc(reviewer, "appleReviewDemo/demoSummary")));
    await assertFails(setDoc(doc(reviewer, "appleReviewDemo/demoWrite"), {
      title: "Not permitted"
    }));
    await assertFails(setDoc(doc(reviewer, "oreReports/demoReviewerHours"), {
      userId: OPERATOR, hours: 1
    }));

    console.log("PASS: synthetic commesse, 3 impianti, squadre, personnel, vehicle,");
    console.log("hours, calendar, documents, admin price list and notifications.");
    console.log("PASS: local-only create/update/delete and reviewer/pending/guest isolation.");
    console.log("NOT PROVEN: iPhone UI, GPS/push, Storage files, real login or full Apple access.");
  } finally {
    await env.cleanup();
  }
}

main().catch((error) => {
  console.error("Synthetic App Review emulator verification FAILED:",
    error?.message || "unknown");
  process.exitCode = 1;
});
