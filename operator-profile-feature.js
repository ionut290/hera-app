(() => {
  'use strict';

  const state = { operator: null, operatorCache: new Map(), pendingLookups: new Map() };
  const CACHE_MISS = Symbol('operator-cache-miss');

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
    try {
      return window.firebase?.firestore?.() || window.db || null;
    } catch (_) {
      return null;
    }
  }

  function getMemoryCollections() {
    const sources = [
      window.personaleRecords,
      window.platformUsers,
      window.operatori,
      window.operators,
      window.utenti,
      window.users
    ];
    const records = [];
    sources.forEach(source => {
      if (Array.isArray(source)) records.push(...source);
      else if (source instanceof Map) records.push(...source.values());
    });
    return records;
  }

  function findOperatorInMemory(email) {
    const normalized = normalizeEmail(email);
    if (!normalized) return null;
    return getMemoryCollections().find(record => normalizeEmail(record?.emailLower || record?.email || record?.mail) === normalized) || null;
  }

  async function queryOperatorByEmail(email) {
    const db = getDb();
    const normalized = normalizeEmail(email);
    if (!db || !normalized) return null;

    const collections = ['operators', 'operatori', 'personale', 'users', 'utenti'];
    for (const collectionName of collections) {
      for (const field of ['emailLower', 'email']) {
        try {
          const result = await db.collection(collectionName).where(field, '==', normalized).limit(1).get();
          if (!result.empty) return { id: result.docs[0].id, ...result.docs[0].data() };
        } catch (_) {}
      }
    }
    return null;
  }

  async function findOperatorByEmail(email) {
    const normalized = normalizeEmail(email);
    if (!normalized) return null;

    const memoryRecord = findOperatorInMemory(normalized);
    if (memoryRecord) {
      state.operatorCache.set(normalized, memoryRecord);
      return memoryRecord;
    }

    if (state.operatorCache.has(normalized)) {
      const cached = state.operatorCache.get(normalized);
      return cached === CACHE_MISS ? null : cached;
    }

    if (state.pendingLookups.has(normalized)) return state.pendingLookups.get(normalized);

    const lookup = queryOperatorByEmail(normalized)
      .then(operator => {
        state.operatorCache.set(normalized, operator || CACHE_MISS);
        return operator;
      })
      .finally(() => state.pendingLookups.delete(normalized));
    state.pendingLookups.set(normalized, lookup);
    return lookup;
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

  function createDialog() {
    if (document.getElementById('operator-profile-dialog')) return;
    const wrapper = document.createElement('div');
    wrapper.innerHTML = `
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
      <dialog id="operator-photo-dialog" class="operator-photo-dialog"><button type="button" aria-label="Chiudi foto">×</button><img alt="Foto operatore ingrandita"></dialog>`;
    document.body.append(...wrapper.children);
  }

  function field(label, icon, value, actionHtml = '') {
    const available = Boolean(value);
    return `<div class="operator-profile-row ${available ? 'is-clickable' : ''}">
      <span class="operator-profile-icon" aria-hidden="true">${icon}</span>
      <span class="operator-profile-label">${escapeHtml(label)}</span>
      <span class="operator-profile-value">${available ? escapeHtml(value) : 'Non disponibile'}</span>
      ${actionHtml}
    </div>`;
  }

  function commesseHtml(items) {
    const values = Array.isArray(items) ? items : String(items || '').split(',').map(v => v.trim()).filter(Boolean);
    if (!values.length) return field('Commesse abilitate', '📁', '');
    const chips = values.map(name => `<button type="button" class="operator-commessa-chip" data-commessa-name="${escapeHtml(name)}">${escapeHtml(name)}</button>`).join('');
    return `<div class="operator-profile-row is-clickable"><span class="operator-profile-icon">📁</span><span class="operator-profile-label">Commesse abilitate</span><span class="operator-profile-value operator-commesse-list">${chips}</span></div>`;
  }

  function safeShow(dialog) {
    try {
      if (typeof dialog.showModal === 'function') dialog.showModal();
      else dialog.setAttribute('open', '');
    } catch (_) {
      dialog.setAttribute('open', '');
    }
  }

  function renderProfile(operator, fallbackName) {
    createDialog();
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
      ? `<button type="button" class="operator-photo-button"><img src="${escapeHtml(photo)}" alt="Foto di ${escapeHtml(fullName)}"></button>`
      : '<div class="operator-photo-placeholder">👤</div>';
    dialog.querySelector('.operator-profile-name').textContent = fullName;
    dialog.querySelector('.operator-profile-role').textContent = ruolo;
    const statusNode = dialog.querySelector('.operator-profile-status');
    statusNode.textContent = status;
    statusNode.classList.toggle('is-disabled', status !== 'Attivo');

    const emailAction = email ? `<a class="operator-profile-inline-action" href="mailto:${escapeHtml(email)}">✉️</a>` : '';
    const phoneAction = phone ? `<a class="operator-profile-inline-action" href="tel:${escapeHtml(phone)}">📞</a>` : '';
    const birthdayAction = birthday ? '<button type="button" class="operator-profile-inline-action" data-operator-action="birthday">📅</button>' : '';
    const matricolaAction = matricola ? '<button type="button" class="operator-profile-inline-action" data-operator-action="copy-matricola">📋</button>' : '';
    const aziendaAction = azienda ? '<button type="button" class="operator-profile-inline-action" data-operator-action="azienda">›</button>' : '';

    dialog.querySelector('.operator-profile-fields').innerHTML = [
      field('E-mail', '✉️', email, emailAction),
      field('Telefono', '📞', phone, phoneAction),
      field('Compleanno', '🎂', birthday, birthdayAction),
      field('Matricola', '🪪', matricola, matricolaAction),
      field('Azienda', '🏢', azienda, aziendaAction),
      commesseHtml(operator?.commesseAbilitate || operator?.commesse || operator?.projects)
    ].join('');

    const callBtn = dialog.querySelector('[data-profile-call]');
    const mailBtn = dialog.querySelector('[data-profile-email]');
    callBtn.hidden = !phone;
    mailBtn.hidden = !email;
    if (phone) callBtn.href = `tel:${phone}`;
    if (email) mailBtn.href = `mailto:${email}`;
    dialog.querySelector('[data-profile-edit]').hidden = !isAdmin();
    safeShow(dialog);
  }

  async function openProfileFromMember(node) {
    try {
      const fallbackName = extractMemberName(node);
      const email = readMemberEmail(node);
      const embedded = node.dataset.operatorJson;
      let operator = null;
      if (embedded) {
        try { operator = JSON.parse(embedded); } catch (_) {}
      }
      if (!operator && email) operator = await findOperatorByEmail(email);
      renderProfile(operator || { displayName: fallbackName, email }, fallbackName);
    } catch (error) {
      console.warn('Scheda operatore non disponibile:', error);
    }
  }

  function markMemberNodes(root = document) {
    const selectors = ['[data-operator-email]','[data-operator-id]','.squadra-member','.team-member','.operator-name','.membro-squadra'];
    root.querySelectorAll?.(selectors.join(',')).forEach(node => {
      if (node.dataset.operatorProfileReady === '1') return;
      node.dataset.operatorProfileReady = '1';
      node.classList.add('operator-member-link');
      node.setAttribute('role', 'button');
      node.setAttribute('tabindex', '0');
    });
  }

  function addBirthdayToCalendar(value, name) {
    const match = String(value || '').match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
    if (!match) return;
    const [, dd, mm, yyyy] = match;
    const date = `${yyyy}${mm.padStart(2, '0')}${dd.padStart(2, '0')}`;
    window.open(`https://calendar.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent('Compleanno ' + name)}&dates=${date}/${date}&recur=RRULE:FREQ=YEARLY`, '_blank', 'noopener');
  }

  document.addEventListener('click', event => {
    const member = event.target.closest?.('.operator-member-link,[data-operator-email],[data-operator-id]');
    if (member && !event.target.closest('button,a,input,select,textarea')) {
      event.preventDefault();
      openProfileFromMember(member);
      return;
    }
    const dialog = document.getElementById('operator-profile-dialog');
    if (!dialog) return;
    if (event.target.closest('.operator-profile-close,.operator-profile-close-bottom')) dialog.close?.();
    if (event.target.closest('[data-operator-action="copy-matricola"]')) navigator.clipboard?.writeText(String(state.operator?.matricola || ''));
    if (event.target.closest('[data-operator-action="birthday"]')) addBirthdayToCalendar(state.operator?.compleanno || state.operator?.dataNascita || state.operator?.birthday, dialog.querySelector('.operator-profile-name').textContent);
    const commessa = event.target.closest('[data-commessa-name]');
    if (commessa) {
      const name = commessa.dataset.commessaName;
      window.dispatchEvent(new CustomEvent('open-commessa-by-name', { detail: { name } }));
      if (typeof window.openCommessaByName === 'function') window.openCommessaByName(name);
      dialog.close?.();
    }
    const photoButton = event.target.closest('.operator-photo-button');
    if (photoButton) {
      const photoDialog = document.getElementById('operator-photo-dialog');
      photoDialog.querySelector('img').src = photoButton.querySelector('img').src;
      safeShow(photoDialog);
    }
    if (event.target.closest('#operator-photo-dialog button')) document.getElementById('operator-photo-dialog').close?.();
  });

  document.addEventListener('keydown', event => {
    const member = event.target.closest?.('.operator-member-link');
    if (member && (event.key === 'Enter' || event.key === ' ')) {
      event.preventDefault();
      openProfileFromMember(member);
    }
  });

  function init() {
    try {
      createDialog();
      markMemberNodes();
      const observer = new MutationObserver(records => records.forEach(record => record.addedNodes.forEach(node => {
        if (node.nodeType === 1) markMemberNodes(node);
      })));
      observer.observe(document.body, { childList: true, subtree: true });
    } catch (error) {
      console.warn('Scheda operatore: inizializzazione saltata', error);
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
  window.openOperatorProfile = renderProfile;
})();
