"use strict";

// Run only in a local Firestore emulator; never against a Firebase project.
// Example: npx firebase emulators:exec --only firestore --project demo-review-isolation //   "node scripts/test-apple-review-firestore-rules.cjs"
const fs = require("node:fs");
const { doc, collection, getDoc, getDocs, setDoc, updateDoc, deleteDoc, query, where, orderBy, limit, serverTimestamp } = require("firebase/firestore");
const assert = require("node:assert/strict");
const { initializeTestEnvironment, assertFails, assertSucceeds } =
  require("@firebase/rules-unit-testing");

async function run() {
  if (!process.env.FIRESTORE_EMULATOR_HOST) {
    throw new Error("FIRESTORE_EMULATOR_HOST is missing. Never run against production.");
  }
  const [host, port] = process.env.FIRESTORE_EMULATOR_HOST.split(":");
  assert.ok(host === "127.0.0.1" || host === "localhost", "Only a local emulator is permitted.");
  const testEnv = await initializeTestEnvironment({
    projectId: "demo-review-isolation",
    firestore: {
      host,
      port: Number(port),
      rules: fs.readFileSync("firestore.rules", "utf8")
    }
  });

  try {
    await testEnv.withSecurityRulesDisabled(async (context) => {
      const db = context.firestore();
      await setDoc(doc(db, "platformUsers/normalUser"), {
        uid: "normalUser", email: "normal@example.test",
        statoAccount: "attivo", accountStatus: "attivo", role: "user"
      });
      await setDoc(doc(db, "platformUsers/otherUser"), {
        uid: "otherUser", email: "other@example.test",
        statoAccount: "attivo", accountStatus: "attivo", role: "user"
      });
      await setDoc(doc(db, "platformUsers/reviewerUser"), {
        uid: "reviewerUser", email: "reviewer@example.test",
        statoAccount: "attivo", accountStatus: "attivo", role: "user"
      });
      await setDoc(doc(db, "commesse/syntheticJob"), { name: "Synthetic fixture" });
      await setDoc(doc(db, "chatMessages/legacyPrivate"), {
        senderId: "normalUser", createdAt: 0, text: "Historical sender-only record"
      });
      await setDoc(doc(db, "chatMessages/legacyOther"), {
        senderId: "otherUser", createdAt: -1, text: "Historical record from other sender"
      });
      await setDoc(doc(db, "commesse/syntheticJob/impianti/testPlant"), { name: "Plant" });
      await setDoc(doc(db, "personale/staff"), { displayName: "Private staff" });
      await setDoc(doc(db, "utenti/staff"), { displayName: "Private user" });
      await setDoc(doc(db, "oreReports/one"), { date: "2026-11-12" });
      await setDoc(doc(db, "gpsUpdateRequests/gpsNormal"), {
        operatorId: "normalUser", status: "pending", createdAt: 1,
        operatorLat: 44.1, operatorLng: 11.1
      });
      await setDoc(doc(db, "gpsUpdateRequests/gpsOther"), {
        operatorId: "otherUser", status: "pending", createdAt: 2,
        operatorLat: 44.2, operatorLng: 11.2
      });
      await setDoc(doc(db, "notificationEmailQueue/emailOther"), {
        to: "other@example.test", status: "pending", html: "synthetic text"
      });
      await setDoc(doc(db, "appNotifications/normalEvent"), {
        eventType: "impianto-done", createdByUid: "normalUser", createdAt: 3,
        title: "Synthetic work done", body: "Synthetic approved event"
      });
      await setDoc(doc(db, "appNotifications/systemEvent"), {
        eventType: "hours-deadline-missing", createdByUid: "system", createdAt: 4,
        title: "Synthetic system deadline", body: "Synthetic system notice"
      });
      await setDoc(doc(db, "hoursDeadlineAlerts/demoDeadline"), {
        alertId: "demoDeadline", recipient: "all", date: "2026-11-19"
      }); 
      await setDoc(doc(db, "userAlertAcknowledgements/alertA__normalUser"), {
        userId: "normalUser", notificationId: "alertA", acknowledgedAt: 1,
        acknowledgedDateKey: "2026-11-19", userName: "Synthetic worker"
      });
      await setDoc(doc(db, "userAlertAcknowledgements/alertB__otherUser"), {
        userId: "otherUser", notificationId: "alertB", acknowledgedAt: 2,
        acknowledgedDateKey: "2026-11-19", userName: "Other synthetic worker"
      });
      await setDoc(doc(db, "sharedStaticViews/currentTeam"), { demo: false });
      await setDoc(doc(db, "appleReviewDemo/onlySynthetic"), {title: "Synthetic exercise", demo: true});
      // Each fixture is synthetic; the tombstone must stay private after
      // the recipient's account is deleted.
      for (const [id, senderId, recipientId, createdAt] of [
        ["public", "otherUser", "", 1],
        ["incoming", "otherUser", "normalUser", 2],
        ["outgoing", "normalUser", "otherUser", 3],
        ["otherSecret", "otherUser", "reviewerUser", 4],
        ["deletedRecipient", "normalUser", "deleted-recipient-example", 5]
      ]) {
        await setDoc(doc(db, "chatMessages/" + id), {
          senderId, recipientId, createdAt, text: "Synthetic chat " + id
        });
      }
    });

    const anonymous = testEnv.unauthenticatedContext().firestore();
    // Admin privileges are still needed for controlled company moderation.
    const administrator = testEnv.authenticatedContext("adminSynthetic", {
      email: "ionut29019@gmail.com", email_verified: true
    }).firestore();
    const normal = testEnv.authenticatedContext("normalUser", {
      email: "normal@example.test", email_verified: true
    }).firestore();
    const other = testEnv.authenticatedContext("otherUser", {
      email: "other@example.test", email_verified: true
    }).firestore();
    const review = testEnv.authenticatedContext("reviewerUser", {
      email: "reviewer@example.test", email_verified: true, apple_review: true
    }).firestore();
    // Even an accidentally whitelisted admin address must never bypass the claim.
    const reviewAdminEmail = testEnv.authenticatedContext("reviewerUser", {
      email: "ionut29019@gmail.com", email_verified: true, apple_review: true
    }).firestore();

    await assertSucceeds(getDoc(doc(review, "appleReviewDemo/onlySynthetic")));
    await assertSucceeds(getDocs(collection(review, "appleReviewDemo")));
    await assertFails(setDoc(doc(review, "appleReviewDemo/onlySynthetic"), { demo: false }));
    await assertFails(updateDoc(doc(review, "appleReviewDemo/onlySynthetic"), { demo: false }));
    await assertFails(deleteDoc(doc(review, "appleReviewDemo/onlySynthetic")));
    await assertFails(getDoc(doc(normal, "appleReviewDemo/onlySynthetic")));

    // Direct reads and queries must enforce privacy at the Firestore server,
    // not merely by filtering a global listener in the browser.
    await assertFails(getDoc(doc(anonymous, "chatMessages/public")));
    await assertFails(getDocs(collection(anonymous, "chatMessages")));
    await assertFails(getDoc(doc(review, "chatMessages/public")));
    await assertFails(getDoc(doc(reviewAdminEmail, "chatMessages/outgoing")));

    await assertSucceeds(getDoc(doc(normal, "chatMessages/public")));
    await assertSucceeds(getDoc(doc(normal, "chatMessages/incoming")));
    await assertSucceeds(getDoc(doc(normal, "chatMessages/outgoing")));
    await assertSucceeds(getDoc(doc(normal, "chatMessages/deletedRecipient")));
    await assertFails(getDoc(doc(normal, "chatMessages/otherSecret")));
    await assertFails(getDoc(doc(other, "chatMessages/deletedRecipient")));
    // Missing recipientId must NOT make an old private message public.
    await assertSucceeds(getDoc(doc(normal, "chatMessages/legacyPrivate")));
    await assertFails(getDoc(doc(other, "chatMessages/legacyPrivate")));
    await assertFails(getDoc(doc(normal, "chatMessages/legacyOther")));
    await assertSucceeds(getDocs(collection(administrator, "chatMessages")));
    await assertFails(getDocs(collection(normal, "chatMessages"))); // global chat listener is forbidden
    await assertFails(getDocs(query(collection(normal, "chatMessages"), limit(500)))); // limit alone must not bypass per-message rules
    for (const [field, value] of [
      ["recipientId", ""], ["recipientId", "normalUser"], ["senderId", "normalUser"]
    ]) {
      await assertSucceeds(getDocs(query(collection(normal, "chatMessages"),
        where(field, "==", value), orderBy("createdAt", "desc"), limit(500))));
    }
    await assertFails(setDoc(doc(normal, "chatMessages/spoofedSystem"),
      { senderId: "system", recipientId: "", createdAt: 7, text: "Fake system notice" }));
    await assertSucceeds(setDoc(doc(administrator, "chatMessages/validSystemNotice"),
      { senderId: "system", recipientId: "", createdAt: 8, text: "Legitimate synthetic admin notice" }));
    await assertFails(setDoc(doc(normal, "chatMessages/spoofedSender"),
      { senderId: "otherUser", recipientId: "normalUser", text: "Spoofed" }));
    await assertSucceeds(setDoc(doc(normal, "chatMessages/validNew"),
      { senderId: "normalUser", recipientId: "otherUser", createdAt: 6, text: "Valid synthetic" }));
    await assertFails(updateDoc(doc(normal, "chatMessages/outgoing"), { recipientId: "" }));
    await assertFails(deleteDoc(doc(normal, "chatMessages/incoming")));

    // GPS updates and notification emails must not be global Firestore collections.
    await assertFails(getDoc(doc(anonymous, "gpsUpdateRequests/gpsNormal")));
    await assertFails(getDoc(doc(review, "gpsUpdateRequests/gpsNormal")));
    await assertFails(getDoc(doc(normal, "gpsUpdateRequests/gpsOther")));
    await assertSucceeds(getDoc(doc(normal, "gpsUpdateRequests/gpsNormal")));
    await assertSucceeds(getDocs(collection(administrator, "gpsUpdateRequests")));
    await assertFails(getDocs(collection(normal, "gpsUpdateRequests")));
    await assertSucceeds(getDocs(query(collection(normal, "gpsUpdateRequests"),
      where("operatorId", "==", "normalUser"), limit(100))));
    await assertSucceeds(setDoc(doc(normal, "gpsUpdateRequests/gpsNew"), {
      operatorId: "normalUser", status: "pending", operatorLat: 44.3, operatorLng: 11.3
    }));
    await assertFails(setDoc(doc(normal, "gpsUpdateRequests/gpsSpoof"), {
      operatorId: "otherUser", status: "pending", operatorLat: 44.3, operatorLng: 11.3
    }));
    await assertFails(setDoc(doc(normal, "gpsUpdateRequests/gpsForgedApproval"), {
      operatorId: "normalUser", status: "approved"
    }));
    await assertFails(updateDoc(doc(normal, "gpsUpdateRequests/gpsNormal"), { status: "approved" }));
    await assertSucceeds(updateDoc(doc(administrator, "gpsUpdateRequests/gpsNormal"), { status: "approved" }));
    await assertFails(deleteDoc(doc(normal, "gpsUpdateRequests/gpsNormal")));

    await assertFails(getDoc(doc(anonymous, "notificationEmailQueue/emailOther")));
    await assertFails(getDoc(doc(review, "notificationEmailQueue/emailOther")));
    await assertFails(getDoc(doc(normal, "notificationEmailQueue/emailOther")));
    await assertFails(getDocs(collection(normal, "notificationEmailQueue")));
    await assertFails(setDoc(doc(normal, "notificationEmailQueue/emailSpoof"), {
      to: "other@example.test", status: "pending", html: "fake"
    }));
    await assertFails(updateDoc(doc(normal, "notificationEmailQueue/emailOther"), { to: "normal@example.test" }));
    await assertFails(deleteDoc(doc(normal, "notificationEmailQueue/emailOther")));
    await assertSucceeds(getDocs(collection(administrator, "notificationEmailQueue")));
    await assertSucceeds(setDoc(doc(administrator, "notificationEmailQueue/emailAdmin"), {
      to: "normal@example.test", status: "pending", html: "only admin sends"
    }));

    // Notification events remain broadly viewable for work coordination, but
    // only the genuine author can add normal events. No fake system records.
    await assertSucceeds(getDoc(doc(normal, "appNotifications/systemEvent")));
    await assertSucceeds(getDocs(collection(normal, "appNotifications")));
    await assertFails(getDoc(doc(anonymous, "appNotifications/normalEvent")));
    await assertFails(getDoc(doc(review, "appNotifications/normalEvent")));
    await assertSucceeds(setDoc(doc(normal, "appNotifications/validNew"), {
      eventType: "impianto-done", createdByUid: "normalUser",
      createdAt: serverTimestamp(), title: "My synthetic event", body: "Owner"
    }));
    await assertFails(setDoc(doc(normal, "appNotifications/fakeOther"), {
      eventType: "impianto-done", createdByUid: "otherUser",
      createdAt: serverTimestamp(), title: "Fake other worker"
    }));
    await assertFails(setDoc(doc(normal, "appNotifications/fakeSystem"), {
      eventType: "hours-deadline-missing", createdByUid: "system",
      createdAt: serverTimestamp(), title: "Fake system"
    }));
    await assertFails(setDoc(doc(normal, "appNotifications/fakeEvent"), {
      eventType: "unrecognized", createdByUid: "normalUser",
      createdAt: serverTimestamp(), title: "Invalid"
    }));
    await assertFails(updateDoc(doc(normal, "appNotifications/normalEvent"), { title: "Tamper" }));
    await assertFails(deleteDoc(doc(normal, "appNotifications/normalEvent")));
    await assertSucceeds(setDoc(doc(administrator, "appNotifications/systemNew"), {
      eventType: "hours-deadline-missing", createdByUid: "system",
      createdAt: serverTimestamp(), title: "Authentic admin notice"
    }));

    await assertSucceeds(getDoc(doc(normal, "hoursDeadlineAlerts/demoDeadline")));
    await assertFails(getDoc(doc(review, "hoursDeadlineAlerts/demoDeadline")));
    await assertFails(setDoc(doc(normal, "hoursDeadlineAlerts/forgedDeadline"), {
      alertId: "forgedDeadline", recipient: "all"
    }));
    await assertFails(updateDoc(doc(normal, "hoursDeadlineAlerts/demoDeadline"), { date: "2099-01-01" }));
    await assertFails(deleteDoc(doc(normal, "hoursDeadlineAlerts/demoDeadline")));
    await assertSucceeds(setDoc(doc(administrator, "hoursDeadlineAlerts/authenticDeadline"), {
      alertId: "authenticDeadline", recipient: "all"
    }));

    // Employees can read or create their own confirmation only; even an
    // empty own acknowledgement must be readable for get-then-create logic.
    await assertSucceeds(getDoc(doc(normal, "userAlertAcknowledgements/alertA__normalUser")));
    await assertSucceeds(getDoc(doc(normal, "userAlertAcknowledgements/notYet__normalUser")));
    await assertFails(getDoc(doc(normal, "userAlertAcknowledgements/alertB__otherUser")));
    await assertFails(getDoc(doc(anonymous, "userAlertAcknowledgements/alertA__normalUser")));
    await assertFails(getDoc(doc(review, "userAlertAcknowledgements/alertA__normalUser")));
    await assertFails(getDocs(collection(normal, "userAlertAcknowledgements")));
    await assertSucceeds(getDocs(collection(administrator, "userAlertAcknowledgements")));
    await assertSucceeds(setDoc(doc(normal, "userAlertAcknowledgements/another__normalUser"), {
      notificationId: "another", userId: "normalUser", userName: "Synthetic",
      acknowledgedAt: serverTimestamp(), acknowledgedDateKey: "2026-11-19"
    }, { merge: true }));
    await assertFails(setDoc(doc(normal, "userAlertAcknowledgements/forged__otherUser"), {
      notificationId: "forged", userId: "otherUser",
      acknowledgedAt: serverTimestamp()
    }));
    await assertFails(setDoc(doc(normal, "userAlertAcknowledgements/notOwnName__normalUser"), {
      notificationId: "differentId", userId: "normalUser",
      acknowledgedAt: serverTimestamp()
    }));
    await assertFails(updateDoc(doc(normal, "userAlertAcknowledgements/alertA__normalUser"), {
      userName: "Forged rewrite"
    }));
    await assertFails(deleteDoc(doc(normal, "userAlertAcknowledgements/alertA__normalUser")));

    await assertSucceeds(getDoc(doc(normal, "commesse/syntheticJob")));
    await assertSucceeds(getDoc(doc(normal, "sharedStaticViews/currentTeam")));
    await assertSucceeds(getDoc(doc(normal, "commesse/syntheticJob/impianti/testPlant")));

    for (const docPath of [
      "commesse/syntheticJob",
      "commesse/syntheticJob/impianti/testPlant",
      "personale/staff",
      "utenti/staff",
      "oreReports/one",
      "sharedStaticViews/currentTeam",
      "platformUsers/reviewerUser"
    ]) {
      await assertFails(getDoc(doc(review, docPath)));
      await assertFails(setDoc(doc(review, docPath), { test: "attempt" }));
      await assertFails(updateDoc(doc(review, docPath), { test: "attempt" }));
      await assertFails(deleteDoc(doc(review, docPath)));
    }

    await assertFails(getDoc(doc(reviewAdminEmail, "appConfig/adminUsers")));
    await assertFails(setDoc(doc(reviewAdminEmail, "appConfig/adminUsers"), { emails: [] }));
    await assertFails(getDocs(collection(review, "commesse")));
    await assertFails(getDocs(collection(review, "personale")));
    console.log("Apple reviewer can only read synthetic demo data; live Firestore read/write denied and normal access unchanged.");
  } finally {
    await testEnv.cleanup();
  }
}
run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
