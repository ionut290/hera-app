(() => {
  'use strict';

  if (window.__rubricaUserEnrichmentLoaded) return;
  window.__rubricaUserEnrichmentLoaded = true;

  const text = (value) => String(value ?? '').trim();
  const emailKey = (value) => text(value).toLowerCase();
  const PHONE_FIELDS = ['telefono','phone','cellulare','mobile','numeroTelefono','telefonoPersonale','phoneNumber'];
  const PHOTO_FIELDS = ['photoURL','fotoUrl','fotoURL','profilePhoto','profilePhotoURL','avatarUrl','avatarURL'];
  const EMAIL_FIELDS = ['email','emailAccessoApp','linkedUserEmail','mail'];
  const first = (object, fields) => fields.map((field) => text(object?.[field])).find(Boolean) || '';

  function users() {
    return Array.isArray(window.platformUsers) ? window.platformUsers : [];
  }

  function personnel() {
    return Array.isArray(window.personaleRecords) ? window.personaleRecords : [];
  }

  function findUser(person) {
    const uid = text(person?.linkedUserId || person?.userId || person?.uid);
    const email = emailKey(first(person, EMAIL_FIELDS));
    return users().find((user) => uid && text(user?.id || user?.uid) === uid)
      || users().find((user) => email && emailKey(user?.email) === email)
      || null;
  }

  function enrichPersonnel() {
    let changed = false;
    personnel().forEach((person) => {
      const email = emailKey(first(person, EMAIL_FIELDS));
      if (!email) return;
      const user = findUser(person);
      if (!user) return;

      const userPhone = first(user, PHONE_FIELDS);
      const userPhoto = first(user, PHOTO_FIELDS);
      if (!first(person, PHONE_FIELDS) && userPhone) {
        person.telefono = userPhone;
        changed = true;
      }
      if (!first(person, PHOTO_FIELDS) && userPhoto) {
        person.photoURL = userPhoto;
        changed = true;
      }
    });
    return changed;
  }

  function contactPhotoFromCard(card) {
    const content = text(card?.textContent).toLowerCase();
    if (!content) return '';
    const matchingPerson = personnel().find((person) => {
      const name = text(person?.displayName || person?.nomeCompleto || [person?.nome, person?.cognome].map(text).filter(Boolean).join(' ')).toLowerCase();
      const email = emailKey(first(person, EMAIL_FIELDS));
      return (email && content.includes(email)) || (name && content.includes(name));
    });
    if (!matchingPerson) return '';
    const user = findUser(matchingPerson);
    return first(matchingPerson, PHOTO_FIELDS) || first(user, PHOTO_FIELDS);
  }

  function decorateCards(root = document) {
    root.querySelectorAll?.('.rubrica-v2-card').forEach((card) => {
      const avatar = card.querySelector('.rubrica-v2-avatar');
      if (!avatar || avatar.querySelector('img')) return;
      const photo = contactPhotoFromCard(card);
      if (!photo) return;
      const image = document.createElement('img');
      image.src = photo;
      image.alt = '';
      image.loading = 'lazy';
      image.referrerPolicy = 'no-referrer';
      image.style.cssText = 'width:100%;height:100%;object-fit:cover;border-radius:50%;display:block';
      image.addEventListener('error', () => image.remove(), { once: true });
      avatar.textContent = '';
      avatar.appendChild(image);
    });
  }

  function refreshOpenRubrica() {
    const page = document.getElementById('rubrica-feature-v2-page');
    if (!page) return;
    decorateCards(page);
  }

  function init() {
    enrichPersonnel();
    refreshOpenRubrica();

    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      const changed = enrichPersonnel();
      if (changed) refreshOpenRubrica();
      if ((personnel().length && users().length) || attempts >= 30) window.clearInterval(timer);
    }, 1000);

    const observer = new MutationObserver((mutations) => {
      for (const mutation of mutations) {
        for (const node of mutation.addedNodes) {
          if (!(node instanceof Element)) continue;
          if (node.matches?.('#rubrica-feature-v2-page,.rubrica-v2-card') || node.querySelector?.('.rubrica-v2-card')) {
            enrichPersonnel();
            decorateCards(node.matches?.('#rubrica-feature-v2-page') ? node : document);
          }
        }
      }
    });
    observer.observe(document.body, { childList: true, subtree: true });
    window.addEventListener('pagehide', () => {
      window.clearInterval(timer);
      observer.disconnect();
    }, { once: true });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();
