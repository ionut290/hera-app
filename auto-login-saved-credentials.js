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

  async function preparePersistentSession() {
    const auth = getAuth();
    if (!auth) {
      authResolved = true;
      return null;
    }

    try {
      if (window.firebase?.auth?.Auth?.Persistence?.LOCAL && typeof auth.setPersistence === 'function') {
        await auth.setPersistence(window.firebase.auth.Auth.Persistence.LOCAL);
      }
    } catch (error) {
      console.warn('Persistenza login locale non configurata:', error);
    }

    if (auth.currentUser) {
      authenticatedUser = auth.currentUser;
      authResolved = true;
      return authenticatedUser;
    }

    return new Promise((resolve) => {
      let settled = false;
      let unsubscribe = () => {};
      const finish = (user) => {
        if (settled) return;
        settled = true;
        authenticatedUser = user || null;
        authResolved = true;
        try { unsubscribe(); } catch (_) {}
        resolve(authenticatedUser);
      };

      const timeout = window.setTimeout(() => finish(auth.currentUser || null), AUTH_RESOLVE_TIMEOUT_MS);
      unsubscribe = auth.onAuthStateChanged(
        (user) => {
          window.clearTimeout(timeout);
          finish(user);
        },
        () => {
          window.clearTimeout(timeout);
          finish(auth.currentUser || null);
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

    if (authenticatedUser || getAuth()?.currentUser) {
      attempted = true;
      stop();
      return;
    }

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
