(() => {
  'use strict';

  const state = { operator: null };

  function normalizeEmail(value) {
    return String(value || '').trim().toLowerCase();
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function getDb() {
    return window.firebase?.firestore?.() || window.db || null;
  }

  async function findOperatorByEmail(email) {
    const db = getDb();
    const normalized = normalizeEmail(email);
    if (!db || !normalized) return null;

    const collections = ['operators', 'operatori', 'personale', 'users', 'utenti'];
    for (const collectionName of collections) {
      try {
        const direct = await db.collection(collectionName).where('emailLower', '==', normalized).limit(1).get();
        if (!direct.empty) return { id: direct.docs[0].id, ...direct.docs[0].data() };
      } catch (_) {}
      try {
        const exact = await db.collection(collectionName).where('email', '==', normalized).limit(1).get();
        if (!exact.empty) return { id: exact.docs[0].id, ...exact.docs[0].data() };
      } catch (_) {}
    }
    return null;
  }

  function getText(node) {
    return String(node?.textContent || '').replace(/\s+/g, ' ').trim();
  }

  function extractMemberName(node) {
    return getText(node).replace(/^[-–—,;\s]+|[-–—,;\s]+$/g, '');
  }

  function readMemberEmail(node) {
    return normalizeEmail(
      node.dataset.operatorEmail ||
      node.dataset.email ||
      node.closest('[data-operator-email]')?.dataset.operatorEmail ||
      node.closest('[data-email]')?.dataset.email || ''
    );
  }

  function isAdmin() {
    return Boolean(window.currentUserIsAdmin || window.isAdmin || document.body.classList.contains('is-admin'));
  }

  function field(label, icon, value, actionHtml = '', clickable = false) {
    const safeValue = value ? escapeHtml(value) : 'Non disponibile';
    return `<div class="operator-profile-row ${clickable ? 'is-clickable' : ''}">
      <span class="operator-profile-icon" aria-hidden="true">${icon}</span>
      <span class="operator-profile-label">${escapeHtml(label)}</span>
      <span class="operator-profile-value">${safeValue}</span>
      ${actionHtml}
    </div>`;
  }

  function commesseHtml(items) {
    const values = Array.isArray(items) ? items : String(items || '').split(',').map(v => v.trim()).filter(Boolean);
    if (!values.length) return field('Commesse abilitate', '📁', 'Non disponibile');
    const chips = values.map(name => `<button type="button" class="operator-commessa-chip" data-commessa-name="${escapeHtml(name)}">${escapeHtml(name)}</button>`).join('');
    return `<div class="operator-profile-row is-clickable">
      <span class="operator-profile-icon" aria-hidden="true">📁</span>
      <span class="operator-profile-label">Commesse abilitate</span>
      <span class="operator-profile-value operator-commesse-list">${chips}</span>
    </div>`;
  }

  function renderProfile(operator, fallbackName) {
    state.operator = operator || {};
    const fullName = operator?.displayName || operator?.nomeCompleto || [operator?.nome, operator?.cognome].filter(Boolean).join(' ') || fallbackName || 'Operatore';
    const photo = operator?.photoURL || operator?.fotoProfilo || operator?.foto || '';
    const email = normalizeEmail(operator?.email || operator?.mail || '');
    const phone = operator?.telefono || operator?.phone || operator?.phoneNumber || '';
    const birthday = operator?.compleanno || operator?.dataNascita || operator?.birthday || '';
    const matricola = operator?.matricola || '';
    const azienda = operator?.azienda || operator?.company || '';
    const ruolo = operator?.ruolo || operator?.role || 'Operatore';
    const status = operator?.attivo === false || operator?.disabled === true ? 'Disattivato' : 'Attivo';

    const dialog = document.getElementById('operator-profile-dialog');
    dialog.querySelector('.operator-profile-photo').innerHTML = photo
      ? `<button type="button" class="operator-photo-button" aria-label="Ingrandisci foto"><img src="${escapeHtml(photo)}" alt="Foto di ${escapeHtml(fullName)}"></button>`
      : '<div class="operator-photo-placeholder" aria-hidden="true">👤</div>';
    dialog.querySelector('.operator-profile-name').textContent = fullName;
    dialog.querySelector('.operator-profile-role').textContent = ruolo;
    dialog.querySelector('.operator-profile-status').textContent = status;
    dialog.querySelector('.operator-profile-status').classList.toggle('is-disabled', status !== 'Attivo');

    const emailAction = email ? `<a class="operator-profile-inline-action" href="mailto:${escapeHtml(email)}" aria-label="Scrivi una email">✉️</a>` : '';
    const phoneAction = phone ? `<a class="operator-profile-inline-action" href="tel:${escapeHtml(phone)}" aria-label="Chiama">📞</a>` : '';
    const birthdayAction = birthday ? '<button type="button" class="operator-profile-inline-action" data-operator-action="birthday" aria-label="Aggiungi compleanno al calendario">📅</button>' : '';
    const matricolaAction = matricola ? '<button type="button" class="operator-profile-inline-action" data-operator-action="copy-matricola" aria-label="Copia matricola">📋</button>' : '';
    const aziendaAction = azienda ? '<button type="button" class="operator-profile-inline-action" data-operator-action="azienda" aria-label="Apri scheda azienda">›</button>' : '';

    dialog.querySelector('.operator-profile-fields').innerHTML = [
      field('E-mail', '✉️', email, emailAction, Boolean(email)),
      field('Telefono', '📞', phone, phoneAction, Boolean(phone)),
      field('Compleanno', '🎂', birthday, birthdayAction, Boolean(birthday)),
      field('Matricola', '🪪', matricola, matricolaAction, Boolean(matricola)),
      field('Azienda', '🏢', azienda, aziendaAction, Boolean(azienda)),
      commesseHtml(operator?.commesseAbilitate || operator?.commesse || operator?.projects)
    ].join('');

    const callBtn = dialog.querySelector('[data-profile-call]');
    const mailBtn = dialog.querySelector('[data-profile-email]');
    callBtn.hidden = !phone;
    mailBtn.hidden = !email;
    if (phone) callBtn.href = `tel:${phone}`;
    if (email) mailBtn.href = `mailto:${email}`;
    dialog.querySelector('[data-profile-edit]').hidden = !isAdmin();

    if (typeof dialog.showModal === 'function') dialog.showModal();
    else dialog.setAttribute('open', '');
  }

  async function openProfileFromMember(node) {
    const fallbackName = extractMemberName(node);
    const email = readMemberEmail(node);
    const embedded = node.dataset.operatorJson;
    let operator = null;
    if (embedded) {
      try { operator = JSON.parse(embedded); } catch (_) {}
    }
    if (!operator && email) operator = await findOperatorByEmail(email);
    if (!operator) operator = { displayName: fallbackName, email };
    renderProfile(operator, fallbackName);
  }

  function markMemberNodes(root = document) {
    const selectors = [
      '[data-operator-email]',
      '[data-operator-id]',
      '.squadra-member',
      '.team-member',
      '.operator-name',
      '.membro-squadra'
    ];
    root.querySelectorAll(selectors.join(',')).forEach(node => {
      if (node.dataset.operatorProfileReady === '1') return;
      node.dataset.operatorProfileReady = '1';
      node.classList.add('operator-member-link');
      node.setAttribute('role', 'button');
      node.setAttribute('tabindex', '0');
    });
  }

  function installDialog() {
    if (document.getElementById('operator-profile-dialog')) return;
    document.body.insertAdjacentHTML('beforeend', `
      <dialog id="operator-profile-dialog" class="operator-profile-dialog">
        <div class="operator-profile-shell">
          <header class="operator-profile-header">
            <button type="button" class="operator-profile-close" aria-label="Chiudi">‹</button>
            <h2>SCHEDA OPERATORE</h2>
          </header>
          <section class="operator-profile-summary">
            <div class="operator-profile-photo"></div>
            <div>
              <button type="button" class="operator-profile-name"></button>
              <p class="operator-profile-role"></p>
              <span class="operator-profile-status"></span>
            </div>
          </section>
          <div class="operator-profile-fields"></div>
          <div class="operator-profile-actions">
            <a class="btn btn-primary" data-profile-call>CHIAMA</a>
            <a class="btn btn-primary" data-profile-email>INVIA E-MAIL</a>
            <button type="button" class="btn operator-profile-edit" data-profile-edit>MODIFICA SCHEDA</button>
            <button type="button" class="btn operator-profile-close-bottom">CHIUDI</button>
          </div>
        </div>
      </dialog>
      <dialog id="operator-photo-dialog" class="operator-photo-dialog"><button type="button" aria-label="Chiudi foto">×</button><img alt="Foto operatore ingrandita"></dialog>
    `);
  }

  function addBirthdayToCalendar(value, name) {
    const match = String(value || '').match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
    if (!match) return;
    const [, dd, mm, yyyy] = match;
    const start = `${yyyy}${mm.padStart(2, '0')}${dd.padStart(2, '0')}`;
    const url = `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent('Compleanno ' + name)}&dates=${start}/${start}&recur=RRULE:FREQ=YEARLY`;
    window.open(url, '_blank', 'noopener');
  }

  document.addEventListener('click', event => {
    const member = event.target.closest('.operator-member-link,[data-operator-email],[data-operator-id]');
    if (member && !event.target.closest('button,a,input,select,textarea')) {
      event.preventDefault();
      openProfileFromMember(member);
      return;
    }

    const dialog = document.getElementById('operator-profile-dialog');
    if (!dialog) return;
    if (event.target.closest('.operator-profile-close,.operator-profile-close-bottom')) dialog.close();
    if (event.target.closest('.operator-profile-name')) return;
    if (event.target.closest('[data-operator-action="copy-matricola"]')) {
      navigator.clipboard?.writeText(String(state.operator?.matricola || ''));
    }
    if (event.target.closest('[data-operator-action="birthday"]')) {
      addBirthdayToCalendar(state.operator?.compleanno || state.operator?.dataNascita || state.operator?.birthday, dialog.querySelector('.operator-profile-name').textContent);
    }
    const commessa = event.target.closest('[data-commessa-name]');
    if (commessa) {
      const name = commessa.dataset.commessaName;
      window.dispatchEvent(new CustomEvent('open-commessa-by-name', { detail: { name } }));
      if (typeof window.openCommessaByName === 'function') window.openCommessaByName(name);
      dialog.close();
    }
    const photoButton = event.target.closest('.operator-photo-button');
    if (photoButton) {
      const photoDialog = document.getElementById('operator-photo-dialog');
      photoDialog.querySelector('img').src = photoButton.querySelector('img').src;
      photoDialog.showModal();
    }
    if (event.target.closest('#operator-photo-dialog button')) document.getElementById('operator-photo-dialog').close();
  });

  document.addEventListener('keydown', event => {
    const member = event.target.closest?.('.operator-member-link');
    if (member && (event.key === 'Enter' || event.key === ' ')) {
      event.preventDefault();
      openProfileFromMember(member);
    }
  });

  function init() {
    installDialog();
    markMemberNodes();
    const observer = new MutationObserver(records => records.forEach(record => record.addedNodes.forEach(node => {
      if (node.nodeType === 1) markMemberNodes(node);
    })));
    observer.observe(document.body, { childList: true, subtree: true });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();

  window.openOperatorProfile = renderProfile;
})();
