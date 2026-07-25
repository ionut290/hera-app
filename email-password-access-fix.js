(function installEmailPasswordAccessFix() {
  "use strict";

  function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
  }

  async function ensurePlatformProfileForAuthenticatedUser(user) {
    if (!user || !user.uid || !window.firebase || typeof firebase.firestore !== "function") return;

    const db = firebase.firestore();
    const currentRef = db.collection("platformUsers").doc(user.uid);
    const currentDoc = await currentRef.get();
    if (currentDoc.exists) return;

    const email = normalizeEmail(user.email);
    let existingProfile = null;

    if (email) {
      const snapshot = await db
        .collection("platformUsers")
        .where("email", "==", email)
        .limit(1)
        .get();
      if (!snapshot.empty) existingProfile = snapshot.docs[0].data() || null;
    }

    const safeProfile = existingProfile ? {
      displayName: existingProfile.displayName || user.displayName || user.email || "Utente",
      email: user.email || existingProfile.email || "",
      teamId: existingProfile.teamId || "",
      role: existingProfile.role || existingProfile.ruolo || "user",
      ruolo: existingProfile.ruolo || existingProfile.role || "user",
      isAdmin: Boolean(existingProfile.isAdmin),
      permissions: existingProfile.permissions || {},
      banned: Boolean(existingProfile.banned),
      bannedReason: existingProfile.bannedReason || null,
      bannedAt: existingProfile.bannedAt || null,
      bannedBy: existingProfile.bannedBy || null
    } : {
      displayName: user.displayName || user.email || "Utente",
      email: user.email || "",
      teamId: "",
      role: "user",
      ruolo: "user",
      isAdmin: false,
      permissions: {},
      banned: false
    };

    await currentRef.set({
      ...safeProfile,
      uid: user.uid,
      authProviders: Array.isArray(user.providerData)
        ? user.providerData.map((provider) => provider && provider.providerId).filter(Boolean)
        : [],
      createdAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      lastSeenAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
  }

  function install() {
    if (!window.firebase || typeof firebase.auth !== "function") return;
    firebase.auth().onAuthStateChanged((user) => {
      if (!user) return;
      void ensurePlatformProfileForAuthenticatedUser(user).catch((error) => {
        console.error("Errore associazione profilo login email/password:", error);
      });
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", install, { once: true });
  } else {
    install();
  }
})();
