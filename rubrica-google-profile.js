(() => {
  'use strict';

  if (window.__vargaRubricaGoogleProfileLoaded) return;
  window.__vargaRubricaGoogleProfileLoaded = true;

  const CACHE_KEY = 'varga_rubrica_google_profiles_v1';
  const SCOPE = 'https://www.googleapis.com/auth/user.phonenumbers.read';
  const text = (value) => String(value ?? '').trim();
  const emailKey = (value) => text(value).toLowerCase();

  function readCache() {
    try {
      const value = JSON.parse(localStorage.getItem(CACHE_KEY) || '{}');
      return value && typeof value === 'object' ? value : {};
    } catch (_) {
      return {};
    }
  }

  function writeCache(value) {
    localStorage.setItem(CACHE_KEY, JSON.stringify(value));
  }

  function mergeIntoPlatformUsers(profile) {
    const users = Array.isArray(window.platformUsers) ? window.platformUsers : [];
    const email = emailKey(profile.email);
    const user = users.find((item) => email && emailKey(item?.email) === email)
      || users.find((item) => text(item?.id || item?.uid) === text(profile.uid));
    if (!user) return;
    if (!text(user.phone || user.telefono || user.cellulare) && profile.phone) user.phone = profile.phone;
    if (!text(user.photoURL || user.photoUrl || user.fotoUrl) && profile.photoURL) user.photoURL = profile.photoURL;
  }

  function applyCachedProfiles() {
    const cache = readCache();
    Object.values(cache).forEach(mergeIntoPlatformUsers);
  }

  async function authorizeAndReadProfile() {
    if (!window.firebase?.auth) throw new Error('Firebase Authentication non disponibile.');
    const auth = window.firebase.auth();
    const current = auth.currentUser;
    if (!current) throw new Error('Accedi prima con Google.');

    const provider = new window.firebase.auth.GoogleAuthProvider();
    provider.addScope(SCOPE);
    provider.setCustomParameters({ prompt: 'consent' });

    const result = await current.reauthenticateWithPopup(provider);
    const credential = result?.credential || window.firebase.auth.GoogleAuthProvider.credentialFromResult?.(result);
    const accessToken = credential?.accessToken;
    if (!accessToken) throw new Error('Autorizzazione Google non disponibile.');

    const response = await fetch('https://people.googleapis.com/v1/people/me?personFields=names,emailAddresses,phoneNumbers,photos', {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    if (!response.ok) throw new Error(`Google People API: errore ${response.status}.`);
    const data = await response.json();

    const phone = text((data.phoneNumbers || []).find((item) => item?.value)?.value);
    const photoURL = text((data.photos || []).find((item) => item?.url)?.url || current.photoURL);
    const email = emailKey((data.emailAddresses || []).find((item) => item?.value)?.value || current.email);
    const name = text((data.names || []).find((item) => item?.displayName)?.displayName || current.displayName);
    const profile = { uid: current.uid, email, name, phone, photoURL, updatedAt: new Date().toISOString() };

    const cache = readCache();
    if (email) cache[email] = profile;
    writeCache(cache);
    mergeIntoPlatformUsers(profile);
    window.dispatchEvent(new CustomEvent('rubrica-google-profile-updated', { detail: profile }));
    return profile;
  }

  function installButton() {
    const page = document.getElementById('rubrica-feature-v2-page');
    const head = page?.querySelector('.rubrica-v2-head');
    if (!head || head.querySelector('[data-rubrica-google-profile]')) return false;
    const button = document.createElement('button');
    button.type = 'button';
    button.dataset.rubricaGoogleProfile = '1';
    button.textContent = '🔄 COMPLETA IL MIO PROFILO GOOGLE';
    button.addEventListener('click', async () => {
      button.disabled = true;
      const original = button.textContent;
      button.textContent = 'Autorizzazione Google…';
      try {
        const profile = await authorizeAndReadProfile();
        const details = [profile.phone ? 'telefono acquisito' : 'telefono non presente', profile.photoURL ? 'foto acquisita' : 'foto non presente'].join(' · ');
        alert(`Profilo Google aggiornato: ${details}.`);
        document.getElementById('rubrica-feature-v2-page')?.remove();
        document.getElementById('today-alerts-btn')?.click();
      } catch (error) {
        console.warn('Rubrica: profilo Google non acquisito.', error);
        alert(error?.message || 'Impossibile leggere il profilo Google.');
      } finally {
        button.disabled = false;
        button.textContent = original;
      }
    });
    head.appendChild(button);
    return true;
  }

  function init() {
    applyCachedProfiles();
    const observer = new MutationObserver(() => installButton());
    observer.observe(document.body, { childList: true, subtree: true });
    installButton();
    window.addEventListener('rubrica-google-profile-updated', applyCachedProfiles);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();
