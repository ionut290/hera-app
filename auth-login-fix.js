(function installAuthStartupController() {
  "use strict";

  if (window.__heraAuthStartupControllerInstalled) return;
  window.__heraAuthStartupControllerInstalled = true;

  // L'auto-login legacy viene caricato tardi dal runtime menu. Il controllo
  // sessione ora vive qui, prima di app.js, quindi impediamo un secondo
  // controller concorrente di modificare l'auth gate dopo l'avvio.
  window.__heraSavedCredentialsAutoLoginInstalled = true;

  const AUTH_GATE_ID = "auth-gate";
  const RESOLVE_TIMEOUT_MS = 10000;
  let authResolved = false;
  let authenticatedUser = null;
  let tokenUnsubscribe = null;
  let gateObserver = null;
  let applyingGateState = false;

  function getAuth() {
    try {
      return window.firebase && typeof firebase.auth === "function" ? firebase.auth() : null;
    } catch (_) {
      return null;
    }
  }

  function isUsableAuthenticatedUser(user) {
    if (!user) return false;
    if (user.email && user.emailVerified === false) return false;
    return true;
  }

  function getGate() {
    return document.getElementById(AUTH_GATE_ID);
  }

  function applyGateHidden(hidden) {
    const gate = getGate();
    if (!gate || applyingGateState) return;
    const currentlyHidden = gate.hidden === true
      || gate.classList.contains("hidden")
      || gate.style.getPropertyValue("display") === "none";
    if (currentlyHidden === hidden) return;

    applyingGateState = true;
    try {
      if (hidden) {
        gate.hidden = true;
        gate.classList.add("hidden");
        gate.style.setProperty("display", "none", "important");
        gate.setAttribute("aria-hidden", "true");
      } else {
        gate.hidden = false;
        gate.classList.remove("hidden");
        gate.style.removeProperty("display");
        gate.removeAttribute("aria-hidden");
      }
    } finally {
      applyingGateState = false;
    }
  }

  function reconcileAuthGate() {
    const auth = getAuth();
    const rawUser = authenticatedUser || (auth && auth.currentUser) || null;
    const usableUser = isUsableAuthenticatedUser(rawUser) ? rawUser : null;

    if (usableUser) {
      authenticatedUser = usableUser;
      applyGateHidden(true);
      return true;
    }

    // Durante il ripristino Firebase nessun altro modulo deve poter mostrare
    // prematuramente "Login richiesto".
    if (!authResolved) {
      applyGateHidden(true);
      return false;
    }

    applyGateHidden(false);
    return false;
  }

  function installGateMutationGuard() {
    const gate = getGate();
    if (!gate || gateObserver || typeof MutationObserver !== "function") return;
    gateObserver = new MutationObserver(() => {
      if (applyingGateState) return;
      queueMicrotask(reconcileAuthGate);
    });
    gateObserver.observe(gate, {
      attributes: true,
      attributeFilter: ["class", "style", "hidden", "aria-hidden"]
    });
  }

  async function ensureLocalPersistence(auth) {
    if (!auth || typeof auth.setPersistence !== "function") return;
    const persistence = window.firebase?.auth?.Auth?.Persistence?.LOCAL;
    if (!persistence) return;
    try {
      await auth.setPersistence(persistence);
    } catch (error) {
      console.warn("Persistenza Firebase LOCAL non disponibile:", error);
    }
  }

  async function startAuthStartupController() {
    installGateMutationGuard();
    applyGateHidden(true);

    const auth = getAuth();
    if (!auth) {
      authResolved = true;
      reconcileAuthGate();
      return;
    }

    await ensureLocalPersistence(auth);

    if (isUsableAuthenticatedUser(auth.currentUser)) {
      authenticatedUser = auth.currentUser;
      authResolved = true;
      reconcileAuthGate();
    }

    if (typeof auth.onIdTokenChanged === "function" && !tokenUnsubscribe) {
      tokenUnsubscribe = auth.onIdTokenChanged((user) => {
        authenticatedUser = isUsableAuthenticatedUser(user) ? user : null;
        authResolved = true;
        reconcileAuthGate();
      }, (error) => {
        console.warn("Ripristino sessione Firebase non riuscito:", error);
        authenticatedUser = isUsableAuthenticatedUser(auth.currentUser) ? auth.currentUser : null;
        authResolved = true;
        reconcileAuthGate();
      });
    }

    window.setTimeout(() => {
      if (authResolved) return;
      authenticatedUser = isUsableAuthenticatedUser(auth.currentUser) ? auth.currentUser : null;
      authResolved = true;
      reconcileAuthGate();
    }, RESOLVE_TIMEOUT_MS);
  }

  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) reconcileAuthGate();
  });
  window.addEventListener("pageshow", () => queueMicrotask(reconcileAuthGate));

  window.HeraAuthStartupController = {
    getState: () => ({
      resolved: authResolved,
      authenticated: reconcileAuthGate()
    }),
    reconcile: reconcileAuthGate
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", startAuthStartupController, { once: true });
  } else {
    void startAuthStartupController();
  }
})();

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
    window.HeraPlatformProfileBootstrap = {
      uid: user.uid,
      exists: currentDoc.exists === true,
      data: currentDoc.exists ? (currentDoc.data() || {}) : null,
      loadedAt: Date.now()
    };
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

    const existingStatus = String(
      existingProfile?.statoAccount || existingProfile?.accountStatus || ""
    ).trim().toLowerCase();
    const initialStatus = existingProfile
      ? (existingProfile.banned === true ? "bloccato" : (existingStatus || "attivo"))
      : "in_attesa";

    await currentRef.set({
      ...safeProfile,
      uid: user.uid,
      statoAccount: initialStatus,
      accountStatus: initialStatus,
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

    document.documentElement.classList.remove("android-email-password-only");

    const message = document.getElementById("auth-gate-message");
    if (message) {
      message.textContent = nativeAndroid
        ? "Accedi con email/username e password. In alternativa puoi usare Google."
        : "Accedi con Google oppure con la tua email e password.";
    }

    const emailLoginButton = document.getElementById("auth-email-login-btn");
    if (emailLoginButton) {
      emailLoginButton.textContent = nativeAndroid ? "Accedi" : "Entra";
    }

    const googleLoginButton = document.getElementById("auth-gate-login-btn");
    if (googleLoginButton) {
      googleLoginButton.hidden = false;
      googleLoginButton.tabIndex = 0;
      googleLoginButton.removeAttribute("aria-hidden");
      googleLoginButton.textContent = nativeAndroid ? "Accedi con Google" : "Login con Google";
    }

    const divider = document.querySelector(".auth-gate-divider");
    if (divider) {
      divider.hidden = false;
      divider.removeAttribute("aria-hidden");
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
    const nativeAuth = getNativeFirebaseAuthentication();
    if (!nativeAuth || typeof nativeAuth.signInWithGoogle !== "function") {
      throw new Error("Plugin Firebase Authentication non disponibile nell'app Android.");
    }

    await firebase.auth().setPersistence(firebase.auth.Auth.Persistence.LOCAL);
    const result = await nativeAuth.signInWithGoogle({ skipNativeAuth: true });
    const idToken = result && result.credential && result.credential.idToken;
    if (!idToken) {
      throw new Error("Google non ha restituito un token di accesso valido.");
    }

    const credential = firebase.auth.GoogleAuthProvider.credential(idToken);
    const signedIn = await firebase.auth().signInWithCredential(credential);
    if (!signedIn?.user?.uid) {
      throw new Error("Accesso Google completato senza un utente valido.");
    }
    return signedIn;
  }

  window.HeraNativeGoogleLogin = signInWithNativeGoogle;

  queueMicrotask(() => {
    window.loginWithGoogle = function loginWithGoogleFixed() {
      return isNativeAndroid() ? signInWithNativeGoogle() : signInWithWebGoogle();
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
      await (isNativeAndroid() ? signInWithNativeGoogle() : signInWithWebGoogle());
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

(function installLoginPasswordVisibility() {
  "use strict";

  function install() {
    const input = document.getElementById("auth-password-input");
    if (!input || document.getElementById("auth-password-toggle")) return;

    const wrapper = document.createElement("div");
    wrapper.className = "auth-password-wrapper";
    wrapper.style.position = "relative";
    wrapper.style.display = "flex";
    wrapper.style.alignItems = "center";
    wrapper.style.width = "100%";

    input.parentNode.insertBefore(wrapper, input);
    wrapper.appendChild(input);
    input.style.paddingRight = "48px";
    input.style.width = "100%";

    const button = document.createElement("button");
    button.id = "auth-password-toggle";
    button.type = "button";
    button.textContent = "👁";
    button.setAttribute("aria-label", "Mostra password");
    button.setAttribute("aria-pressed", "false");
    button.title = "Mostra password";
    button.style.position = "absolute";
    button.style.right = "8px";
    button.style.top = "50%";
    button.style.transform = "translateY(-50%)";
    button.style.border = "0";
    button.style.background = "transparent";
    button.style.cursor = "pointer";
    button.style.fontSize = "20px";
    button.style.padding = "6px";
    button.style.lineHeight = "1";
    button.style.zIndex = "2";

    button.addEventListener("click", () => {
      const visible = input.type === "text";
      input.type = visible ? "password" : "text";
      button.textContent = visible ? "👁" : "🙈";
      button.setAttribute("aria-label", visible ? "Mostra password" : "Nascondi password");
      button.setAttribute("aria-pressed", visible ? "false" : "true");
      button.title = visible ? "Mostra password" : "Nascondi password";
      input.focus({ preventScroll: true });
      try {
        const end = input.value.length;
        input.setSelectionRange(end, end);
      } catch (_) {}
    });

    wrapper.appendChild(button);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", install, { once: true });
  } else {
    install();
  }
})();