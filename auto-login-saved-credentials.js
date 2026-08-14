(() => {
  'use strict';

  if (window.__heraSavedCredentialsAutoLoginInstalled) return;
  window.__heraSavedCredentialsAutoLoginInstalled = true;

  const EMAIL_ID = 'auth-email-input';
  const PASSWORD_ID = 'auth-password-input';
  const LOGIN_BUTTON_ID = 'auth-email-login-btn';
  const MAX_WAIT_MS = 10000;
  const STABLE_DELAY_MS = 500;
  const CHECK_INTERVAL_MS = 200;
  const AUTH_RESOLVE_TIMEOUT_MS = 6000;
  const AUTH_GATE_ID = 'auth-gate';

  let attempted = false;
  let stableSince = 0;
  let timer = null;
  let authResolved = false;
  let authenticatedUser = null;
  const startedAt = Date.now();

  const normalizeEmail = (value) => String(value || '').trim().toLowerCase();

  function stop() {
    if (timer !== null) window.clearInterval(timer);
    timer = null;
  }

  function getAuth() {
    try {
      return window.firebase?.auth?.() || null;
    } catch (_) {
      return null;
    }
  }

  function getAuthGate() {
    return document.getElementById(AUTH_GATE_ID);
  }

  function setAuthGatePending(pending) {
    const gate = getAuthGate();
    if (!gate) return;
    if (pending) {
      gate.dataset.authStartupPending = '1';
      gate.style.setProperty('display', 'none', 'important');
      gate.setAttribute('aria-hidden', 'true');
    } else {
      delete gate.dataset.authStartupPending;
      gate.style.removeProperty('display');
      gate.removeAttribute('aria-hidden');
    }
  }

  function keepGateHiddenForAuthenticatedUser() {
    const gate = getAuthGate();
    if (!gate) return;
    gate.hidden = true;
    gate.style.setProperty('display', 'none', 'important');
    gate.setAttribute('aria-hidden', 'true');
  }

  function revealGateForSignedOutUser() {
    const gate = getAuthGate();
    if (!gate) return;
    gate.hidden = false;
    gate.style.removeProperty('display');
    gate.removeAttribute('aria-hidden');
  }

  async function ensureLocalPersistence(auth) {
    if (!auth || typeof auth.setPersistence !== 'function') return;
    const localPersistence = window.firebase?.auth?.Auth?.Persistence?.LOCAL;
    if (!localPersistence) return;
    try {
      await auth.setPersistence(localPersistence);
    } catch (error) {
      console.warn('Persistenza login locale non configurata:', error);
    }
  }

  async function preparePersistentSession() {
    const auth = getAuth();
    if (!auth) {
      authResolved = true;
      setAuthGatePending(false);
      return null;
    }

    setAuthGatePending(true);

    if (auth.currentUser) {
      authenticatedUser = auth.currentUser;
      await ensureLocalPersistence(auth);
      authResolved = true;
      keepGateHiddenForAuthenticatedUser();
      return authenticatedUser;
    }

    return new Promise((resolve) => {
      let settled = false;
      let unsubscribe = () => {};

      const finish = async (user) => {
        if (settled) return;
        settled = true;
        authenticatedUser = user || null;
        if (authenticatedUser) await ensureLocalPersistence(auth);
        authResolved = true;
        try { unsubscribe(); } catch (_) {}
        if (authenticatedUser) keepGateHiddenForAuthenticatedUser();
        else revealGateForSignedOutUser();
        resolve(authenticatedUser);
      };

      const timeout = window.setTimeout(() => {
        void finish(auth.currentUser || null);
      }, AUTH_RESOLVE_TIMEOUT_MS);

      unsubscribe = auth.onAuthStateChanged(
        (user) => {
          window.clearTimeout(timeout);
          void finish(user);
        },
        () => {
          window.clearTimeout(timeout);
          void finish(auth.currentUser || null);
        }
      );
    });
  }

  function credentialsReady() {
    const emailInput = document.getElementById(EMAIL_ID);
    const passwordInput = document.getElementById(PASSWORD_ID);
    const button = document.getElementById(LOGIN_BUTTON_ID);
    if (!emailInput || !passwordInput || !button) return null;

    const email = normalizeEmail(emailInput.value);
    const password = String(passwordInput.value || '');
    if (!email.includes('@') || password.length < 6 || button.disabled) return null;
    return { emailInput, passwordInput, button };
  }

  function tryAutoLogin() {
    if (attempted) return stop();
    if (!authResolved) return;
    if (Date.now() - startedAt > MAX_WAIT_MS) return stop();
    if (document.visibilityState === 'hidden') return;

    const auth = getAuth();
    if (authenticatedUser || auth?.currentUser) {
      attempted = true;
      keepGateHiddenForAuthenticatedUser();
      return stop();
    }

    const ready = credentialsReady();
    if (!ready) {
      stableSince = 0;
      return;
    }

    if (!stableSince) {
      stableSince = Date.now();
      return;
    }
    if (Date.now() - stableSince < STABLE_DELAY_MS) return;

    attempted = true;
    stop();
    ready.emailInput.setAttribute('aria-busy', 'true');
    ready.passwordInput.setAttribute('aria-busy', 'true');
    ready.button.click();
    window.setTimeout(() => {
      ready.emailInput.removeAttribute('aria-busy');
      ready.passwordInput.removeAttribute('aria-busy');
    }, 1500);
  }

  async function start() {
    if (timer !== null || attempted) return;
    await preparePersistentSession();

    const auth = getAuth();
    if (authenticatedUser || auth?.currentUser) {
      attempted = true;
      keepGateHiddenForAuthenticatedUser();
      stop();
      return;
    }

    await ensureLocalPersistence(auth);
    revealGateForSignedOutUser();
    timer = window.setInterval(tryAutoLogin, CHECK_INTERVAL_MS);
    tryAutoLogin();
  }

  document.addEventListener('click', (event) => {
    if (event.target?.closest?.(`#${LOGIN_BUTTON_ID}`)) {
      attempted = true;
      stop();
    }
  }, true);

  document.addEventListener('visibilitychange', () => {
    if (!document.hidden && !attempted) start();
  });

  window.HeraAutoLoginSession = {
    installed: true,
    getState: () => ({
      authResolved,
      authenticated: Boolean(authenticatedUser || getAuth()?.currentUser),
      fallbackAttempted: attempted
    })
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start, { once: true });
  } else {
    start();
  }
})();
