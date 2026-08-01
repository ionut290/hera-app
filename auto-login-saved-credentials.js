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

  let attempted = false;
  let stableSince = 0;
  let timer = null;
  const startedAt = Date.now();

  const normalizeEmail = (value) => String(value || '').trim().toLowerCase();

  function stop() {
    if (timer !== null) window.clearInterval(timer);
    timer = null;
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
    if (Date.now() - startedAt > MAX_WAIT_MS) return stop();
    if (document.visibilityState === 'hidden') return;
    if (window.firebase?.auth?.()?.currentUser) return stop();

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

  function start() {
    if (timer !== null || attempted) return;
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
    if (!document.hidden) start();
  });

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start, { once: true });
  } else {
    start();
  }
})();
