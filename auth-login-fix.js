(function installGoogleLoginFix() {
  "use strict";

  const LOGIN_BUTTON_IDS = new Set(["login-btn", "auth-gate-login-btn"]);
  let loginInProgress = false;

  function isNativeAndroid() {
    return Boolean(
      window.Capacitor &&
      typeof window.Capacitor.isNativePlatform === "function" &&
      window.Capacitor.isNativePlatform() &&
      typeof window.Capacitor.getPlatform === "function" &&
      window.Capacitor.getPlatform() === "android"
    );
  }

  function configureAndroidEmailPasswordOnly() {
    if (!isNativeAndroid()) return;

    document.documentElement.classList.add("android-email-password-only");

    const message = document.getElementById("auth-gate-message");
    if (message) {
      message.textContent = "Inserisci nome, ruolo, email e password per accedere.";
    }

    const profileFields = document.getElementById("auth-profile-fields");
    if (profileFields) profileFields.classList.remove("hidden");

    const nameInput = document.getElementById("auth-name-input");
    const roleInput = document.getElementById("auth-role-input");
    if (nameInput) nameInput.required = true;
    if (roleInput) roleInput.required = true;

    const emailLoginButton = document.getElementById("auth-email-login-btn");
    if (emailLoginButton) emailLoginButton.textContent = "Accedi";

    const googleLoginButton = document.getElementById("auth-gate-login-btn");
    if (googleLoginButton) {
      googleLoginButton.hidden = true;
      googleLoginButton.setAttribute("aria-hidden", "true");
      googleLoginButton.tabIndex = -1;
    }

    const divider = document.querySelector(".auth-gate-divider");
    if (divider) {
      divider.hidden = true;
      divider.setAttribute("aria-hidden", "true");
    }
  }

  async function saveAndroidProfile(user) {
    if (!isNativeAndroid() || !user) return;

    const nameInput = document.getElementById("auth-name-input");
    const roleInput = document.getElementById("auth-role-input");
    const name = String(nameInput && nameInput.value ? nameInput.value : "").trim();
    const role = String(roleInput && roleInput.value ? roleInput.value : "").trim();

    if (!name || !role) {
      throw new Error("Inserisci nome e cognome e seleziona il ruolo.");
    }

    if (user.displayName !== name && typeof user.updateProfile === "function") {
      await user.updateProfile({ displayName: name });
    }

    const profile = {
      uid: user.uid,
      email: user.email || "",
      name,
      displayName: name,
      role,
      accessLevel: "full",
      canAccessAllData: true,
      updatedAt: new Date().toISOString()
    };

    localStorage.setItem("varga-user-profile", JSON.stringify(profile));
    localStorage.setItem("operatorName", name);
    localStorage.setItem("operatorRole", role);

    try {
      if (window.firebase && firebase.firestore) {
        await firebase.firestore().collection("users").doc(user.uid).set(
          {
            ...profile,
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
          },
          { merge: true }
        );
      }
    } catch (error) {
      console.warn("Profilo salvato solo localmente:", error);
    }
  }

  function installAndroidEmailLoginProfileHook() {
    const form = document.getElementById("auth-email-form");
    if (!form || form.dataset.androidProfileHook === "1") return;
    form.dataset.androidProfileHook = "1";

    form.addEventListener("submit", async () => {
      if (!isNativeAndroid()) return;

      const name = String(document.getElementById("auth-name-input")?.value || "").trim();
      const role = String(document.getElementById("auth-role-input")?.value || "").trim();
      if (!name || !role) return;

      const stopAt = Date.now() + 10000;
      while (Date.now() < stopAt) {
        const user = window.firebase && firebase.auth ? firebase.auth().currentUser : null;
        if (user) {
          try {
            await saveAndroidProfile(user);
          } catch (error) {
            console.error("Salvataggio profilo fallito:", error);
          }
          return;
        }
        await new Promise((resolve) => setTimeout(resolve, 200));
      }
    }, true);
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
      configureAndroidEmailPasswordOnly();
      const nameInput = document.getElementById("auth-name-input");
      if (nameInput) nameInput.focus();
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

  function initializeAndroidLogin() {
    configureAndroidEmailPasswordOnly();
    installAndroidEmailLoginProfileHook();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initializeAndroidLogin, { once: true });
  } else {
    initializeAndroidLogin();
  }

  document.addEventListener("click", handleGoogleLoginClick, true);
})();
