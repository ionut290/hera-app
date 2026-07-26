(function installGoogleLoginFix() {
  "use strict";

  const LOGIN_BUTTON_IDS = new Set(["login-btn", "auth-gate-login-btn"]);
  let loginInProgress = false;

  function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
  }

  async function ensurePlatformProfileForAuthenticatedUser(user) {
    if (!user || !user.uid || !window.firebase || typeof firebase.firestore !== "function") return;

    const database = firebase.firestore();
    const currentRef = database.collection("platformUsers").doc(user.uid);
    const currentDoc = await currentRef.get();
    if (currentDoc.exists) return;

    const email = normalizeEmail(user.email);
    let existingProfile = null;

    if (email) {
      try {
        const exactSnapshot = await database
          .collection("platformUsers")
          .where("email", "==", email)
          .limit(1)
          .get();

        if (!exactSnapshot.empty) {
          existingProfile = exactSnapshot.docs[0].data() || null;
        } else {
          const originalEmailSnapshot = await database
            .collection("platformUsers")
            .where("email", "==", String(user.email || "").trim())
            .limit(1)
            .get();
          if (!originalEmailSnapshot.empty) {
            existingProfile = originalEmailSnapshot.docs[0].data() || null;
          }
        }
      } catch (lookupError) {
        console.warn(
          "Profilo precedente non leggibile: creo un profilo utente standard.",
          lookupError
        );
      }
    }

    const safeProfile = existingProfile ? {
      displayName: existingProfile.displayName || user.displayName || user.email || "Utente",
      email: user.email || existingProfile.email || "",
      teamId: existingProfile.teamId || "",
      role: existingProfile.role || existingProfile.ruolo || "user",
      ruolo: existingProfile.ruolo || existingProfile.role || "user",
      isAdmin: Boolean(existingProfile.isAdmin || existingProfile.admin),
      admin: Boolean(existingProfile.isAdmin || existingProfile.admin),
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
      admin: false,
      permissions: {}
    };

    await currentRef.set({
      ...safeProfile,
      uid: user.uid,
      authProviders: Array.isArray(user.providerData)
        ? user.providerData.map((provider) => provider && provider.providerId).filter(Boolean)
        : [],
      profileMigratedByEmail: Boolean(existingProfile),
      createdAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      lastSeenAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
  }

  function requiresEmailVerification(user) {
    return Boolean(user?.email && user.emailVerified === false);
  }

  function showEmailVerificationRequired() {
    const message = "Prima di accedere ai dati, apri l’email di verifica e conferma il tuo indirizzo.";
    const gateMessage = document.getElementById("auth-gate-message");
    const feedback = document.getElementById("auth-email-feedback");
    if (gateMessage) gateMessage.textContent = message;
    if (feedback) feedback.textContent = message;
  }

  function installProfileAccessGuard() {
    if (!window.firebase || typeof firebase.auth !== "function") return;
    const authInstance = firebase.auth();
    if (!authInstance || authInstance.__heraEmailAccessGuardInstalled) return;

    const originalOnAuthStateChanged = authInstance.onAuthStateChanged.bind(authInstance);
    authInstance.onAuthStateChanged = function onAuthStateChangedWithProfile(nextOrObserver, error, completed) {
      const wrapCallback = (callback) => async (user) => {
        const emailVerificationRequired = requiresEmailVerification(user);
        if (user) {
          try {
            await ensurePlatformProfileForAuthenticatedUser(user);
          } catch (profileError) {
            console.error("Errore preparazione profilo utente autenticato:", profileError);
          }
        }

        const effectiveUser = emailVerificationRequired ? null : user;
        if (emailVerificationRequired) window.__heraEmailVerificationRequired = true;
        else if (user) window.__heraEmailVerificationRequired = false;

        const result = typeof callback === "function"
          ? await callback(effectiveUser)
          : undefined;
        if (emailVerificationRequired) showEmailVerificationRequired();
        return result;
      };

      if (typeof nextOrObserver === "function") {
        return originalOnAuthStateChanged(wrapCallback(nextOrObserver), error, completed);
      }

      const observer = nextOrObserver || {};
      return originalOnAuthStateChanged({
        next: wrapCallback(observer.next),
        error: observer.error,
        complete: observer.complete
      });
    };

    Object.defineProperty(authInstance, "__heraEmailAccessGuardInstalled", {
      value: true,
      configurable: false,
      enumerable: false
    });
  }

  function isNativeAndroid() {
    return Boolean(
      window.Capacitor &&
      typeof window.Capacitor.isNativePlatform === "function" &&
      window.Capacitor.isNativePlatform() &&
      typeof window.Capacitor.getPlatform === "function" &&
      window.Capacitor.getPlatform() === "android"
    );
  }

  function configurePlatformLoginOptions() {
    const nativeAndroid = isNativeAndroid();
    const webGoogleEnabled = !nativeAndroid;

    document.documentElement.classList.toggle("android-email-password-only", nativeAndroid);

    const message = document.getElementById("auth-gate-message");
    if (message) {
      message.textContent = nativeAndroid
        ? "Accedi con la tua email e password."
        : "Accedi con Google oppure con la tua email e password.";
    }

    const emailLoginButton = document.getElementById("auth-email-login-btn");
    if (emailLoginButton) {
      emailLoginButton.textContent = nativeAndroid ? "Accedi" : "Entra";
    }

    const googleLoginButton = document.getElementById("auth-gate-login-btn");
    if (googleLoginButton) {
      googleLoginButton.hidden = !webGoogleEnabled;
      googleLoginButton.tabIndex = webGoogleEnabled ? 0 : -1;
      if (webGoogleEnabled) googleLoginButton.removeAttribute("aria-hidden");
      else googleLoginButton.setAttribute("aria-hidden", "true");
    }

    const divider = document.querySelector(".auth-gate-divider");
    if (divider) {
      divider.hidden = !webGoogleEnabled;
      if (webGoogleEnabled) divider.removeAttribute("aria-hidden");
      else divider.setAttribute("aria-hidden", "true");
    }
  }

  function getNativeFirebaseAuthentication() {
    if (!window.Capacitor) return null;
    if (window.Capacitor.Plugins && window.Capacitor.Plugins.FirebaseAuthentication) {
      return window.Capacitor.Plugins.FirebaseAuthentication;
    }
    if (typeof window.Capacitor.registerPlugin === "function") {
      return window.Capacitor.registerPlugin("FirebaseAuthentication");
    }
    return null;
  }

  function formatError(error) {
    const code = String(error && error.code ? error.code : "");
    if (code === "auth/popup-closed-by-user") return "Accesso Google annullato.";
    if (code === "12501" || code === "16" || code === "auth/cancelled-popup-request") {
      return "Accesso Google annullato.";
    }
    if (code === "10" || code.includes("DEVELOPER_ERROR")) {
      return "Login Google Android non configurato correttamente. Verifica SHA-1 e google-services.json.";
    }
    if (code === "auth/popup-blocked" || code === "auth/cancelled-popup-request") {
      return "Il browser ha bloccato la finestra Google. Consenti i popup per Varga Cantieri e riprova.";
    }
    return String(error && error.message ? error.message : "Accesso Google non riuscito.");
  }

  async function signInWithNativeGoogle() {
    if (isNativeAndroid()) {
      throw new Error("Nell'app Android è disponibile solo l'accesso con email e password.");
    }

    const nativeAuth = getNativeFirebaseAuthentication();
    if (!nativeAuth || typeof nativeAuth.signInWithGoogle !== "function") {
      throw new Error("Plugin Firebase Authentication non disponibile nell'app Android.");
    }

    const result = await nativeAuth.signInWithGoogle({ skipNativeAuth: true });
    const idToken = result && result.credential && result.credential.idToken;
    if (!idToken) {
      throw new Error("Google non ha restituito un token di accesso valido.");
    }

    const credential = firebase.auth.GoogleAuthProvider.credential(idToken);
    return firebase.auth().signInWithCredential(credential);
  }

  window.HeraNativeGoogleLogin = signInWithNativeGoogle;

  queueMicrotask(() => {
    window.loginWithGoogle = function loginWithGoogleFixed() {
      if (isNativeAndroid()) {
        return Promise.reject(
          new Error("Nell'app Android è disponibile solo l'accesso con email e password.")
        );
      }
      return signInWithWebGoogle();
    };
  });

  function signInWithWebGoogle() {
    const provider = new firebase.auth.GoogleAuthProvider();
    provider.addScope("https://www.googleapis.com/auth/userinfo.email");
    provider.setCustomParameters({ prompt: "select_account" });
    return firebase.auth().signInWithPopup(provider);
  }

  async function handleGoogleLoginClick(event) {
    const button = event.target && event.target.closest
      ? event.target.closest("button")
      : null;
    if (!button || !LOGIN_BUTTON_IDS.has(button.id)) return;

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();

    if (isNativeAndroid()) {
      configurePlatformLoginOptions();
      const emailInput = document.getElementById("auth-email-input");
      if (emailInput) emailInput.focus();
      return;
    }

    if (loginInProgress) return;

    if (!window.firebase || !firebase.auth || !firebase.auth.GoogleAuthProvider) {
      alert("Login Google non disponibile: configurazione Firebase non caricata.");
      return;
    }

    loginInProgress = true;
    button.disabled = true;
    const previousText = button.textContent;
    button.textContent = "Accesso Google...";

    try {
      await signInWithWebGoogle();
    } catch (error) {
      console.error("Login Google fallito:", error);
      alert(formatError(error));
    } finally {
      loginInProgress = false;
      button.disabled = false;
      button.textContent = previousText;
    }
  }

  installProfileAccessGuard();

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", configurePlatformLoginOptions, { once: true });
  } else {
    configurePlatformLoginOptions();
  }

  document.addEventListener("click", handleGoogleLoginClick, true);
})();