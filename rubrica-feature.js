(() => {
  'use strict';

  if (window.__vargaRubricaFeatureLoaded) return;
  window.__vargaRubricaFeatureLoaded = true;

  const STORAGE_KEY = 'varga_rubrica_contatti_manual_v1';
  const PHONE_FIELDS = ['telefono', 'phone', 'cellulare', 'mobile', 'numeroTelefono', 'telefonoPersonale'];
  const EMAIL_FIELDS = ['email', 'emailAccessoApp', 'linkedUserEmail', 'mail'];
  const text = (value) => String(value ?? '').trim();
  const esc = (value) => text(value).replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
  const normalizePhone = (value) => text(value).replace(/[^\d+]/g, '').replace(/^00/, '+');
  const normalizeHeader = (value) => text(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '').toUpperCase();
  const whatsappPhone = (value) => {
    let phone = normalizePhone(value).replace(/\D/g, '');
    if (phone.startsWith('0')) phone = `39${phone.slice(1)}`;
    if (phone && !phone.startsWith('39') && phone.length <= 10) phone = `39${phone}`;
    return phone;
  };
  const firstValue = (object, fields) => fields.map((field) => text(object?.[field])).find(Boolean) || '';

  function readManualContacts() {
    try {
      const value = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
      return Array.isArray(value) ? value : [];
    } catch (_) {
      return [];
    }
  }

  function writeManualContacts(rows) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
  }

  function appUsers() {
    return Array.isArray(window.platformUsers) ? window.platformUsers : [];
  }

  function personnel() {
    return Array.isArray(window.personaleRecords) ? window.personaleRecords : [];
  }

  function displayName(person) {
    if (typeof window.getPersonaleDisplayName === 'function') {
      const value = text(window.getPersonaleDisplayName(person));
      if (value) return value;
    }
    return text(person?.displayName || person?.nomeCompleto || `${text(person?.nome)} ${text(person?.cognome)}` || person?.name);
  }

  function linkedUser(person) {
    const users = appUsers();
    const uid = text(person?.linkedUserId || person?.userId || person?.uid);
    const email = firstValue(person, EMAIL_FIELDS).toLowerCase();
    return users.find((user) => uid && text(user?.id || user?.uid) === uid)
      || users.find((user) => email && text(user?.email).toLowerCase() === email)
      || null;
  }

  function buildContacts() {
    const byPhone = new Map();
    personnel().forEach((person) => {
      const phone = firstValue(person, PHONE_FIELDS);
      if (!phone) return;
      const user = linkedUser(person);
      const normalized = normalizePhone(phone);
      if (!normalized) return;
      const contact = {
        id: `operator:${text(person?.id || normalized)}`,
        source: 'operator',
        name: displayName(person) || text(user?.displayName) || 'Operatore',
        phone,
        email: firstValue(person, EMAIL_FIELDS) || text(user?.email),
        role: text(person?.mansione || person?.ruolo || person?.role),
        company: text(person?.azienda || person?.company),
        editable: false
      };
      const existing = byPhone.get(normalized);
      if (!existing || (!existing.email && contact.email)) byPhone.set(normalized, { ...existing, ...contact });
    });

    readManualContacts().forEach((contact) => {
      const normalized = normalizePhone(contact?.phone);
      if (!normalized) return;
      byPhone.set(normalized, { ...contact, id: contact.id || `manual:${Date.now()}`, source: 'manual', editable: true });
    });

    return [...byPhone.values()].sort((a, b) => text(a.name).localeCompare(text(b.name), 'it', { sensitivity: 'base' }));
  }

  function isManager() {
    try { return typeof window.canManageData === 'function' && window.canManageData(); } catch (_) { return false; }
  }

  function installStyle() {
    if (document.getElementById('rubrica-feature-style')) return;
    const style = document.createElement('style');
    style.id = 'rubrica-feature-style';
    style.textContent = `
      .rubrica-overlay{position:fixed;inset:0;z-index:12000;background:#eef4f2;overflow:auto;padding:env(safe-area-inset-top) 14px calc(20px + env(safe-area-inset-bottom))}
      .rubrica-shell{width:min(760px,100%);margin:0 auto;padding:14px 0 30px}
      .rubrica-head{position:sticky;top:0;z-index:2;background:#eef4f2;padding:8px 0 12px;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
      .rubrica-head h2{margin:0;flex:1;color:#12352f}.rubrica-back,.rubrica-add,.rubrica-import{min-height:44px;border:1px solid #bfd0cb;border-radius:12px;background:#fff;padding:8px 12px;font-weight:800}
      .rubrica-search{width:100%;min-height:48px;border:1px solid #bfd0cb;border-radius:14px;padding:0 14px;font-size:16px;background:#fff;margin-bottom:12px}
      .rubrica-list{display:grid;gap:10px}.rubrica-card{width:100%;border:1px solid #d3dedb;border-radius:18px;background:#fff;padding:14px;text-align:left;display:flex;gap:12px;align-items:center;box-shadow:0 5px 18px rgba(15,55,48,.06)}
      .rubrica-avatar{width:46px;height:46px;border-radius:50%;display:grid;place-items:center;background:#e5f2ee;font-weight:900;color:#164d42;flex:none}.rubrica-main{min-width:0;flex:1}.rubrica-main strong,.rubrica-main small{display:block}.rubrica-main strong{font-size:1rem;color:#15332d}.rubrica-main small{color:#64736f;margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.rubrica-empty{padding:24px;text-align:center;color:#667773}
      .rubrica-actions-overlay{position:fixed;inset:0;z-index:12010;background:rgba(6,18,16,.55);display:flex;align-items:flex-end;justify-content:center;padding:14px}.rubrica-actions{width:min(520px,100%);background:#fff;border-radius:20px;padding:18px;display:grid;gap:10px}.rubrica-actions h3{margin:0 0 2px}.rubrica-action{min-height:48px;border:0;border-radius:12px;font-weight:900;font-size:1rem;text-decoration:none;display:flex;align-items:center;justify-content:center;background:#edf4f2;color:#143b33}.rubrica-action.primary{background:#13795b;color:#fff}.rubrica-action.danger{background:#fff0f0;color:#9f1d1d}.rubrica-form{display:grid;gap:10px}.rubrica-form input{min-height:46px;border:1px solid #c7d6d2;border-radius:11px;padding:0 12px;font-size:16px}
      @media(max-width:620px){.rubrica-head h2{flex-basis:calc(100% - 130px)}.rubrica-import{order:3;flex:1}.rubrica-add{order:4;flex:1}}
      @media(min-width:700px){.rubrica-actions-overlay{align-items:center}}
    `;
    document.head.appendChild(style);
  }

  function closeActions() {
    document.querySelector('.rubrica-actions-overlay')?.remove();
  }

  function openContact(contact, rerender) {
    closeActions();
    const overlay = document.createElement('div');
    overlay.className = 'rubrica-actions-overlay';
    const wa = whatsappPhone(contact.phone);
    overlay.innerHTML = `<section class="rubrica-actions" role="dialog" aria-modal="true">
      <h3>${esc(contact.name)}</h3><p>${esc(contact.phone)}${contact.email ? ` · ${esc(contact.email)}` : ''}</p>
      <a class="rubrica-action primary" href="tel:${esc(normalizePhone(contact.phone))}">📞 CHIAMA</a>
      <a class="rubrica-action" href="https://wa.me/${esc(wa)}" target="_blank" rel="noopener">💬 WHATSAPP</a>
      ${contact.email ? `<a class="rubrica-action" href="mailto:${esc(contact.email)}">✉️ E-MAIL</a>` : ''}
      ${contact.editable && isManager() ? '<button class="rubrica-action" type="button" data-rubrica-edit>✏️ MODIFICA</button><button class="rubrica-action danger" type="button" data-rubrica-delete>🗑️ ELIMINA</button>' : ''}
      <button class="rubrica-action" type="button" data-rubrica-close>CHIUDI</button>
    </section>`;
    overlay.addEventListener('click', (event) => { if (event.target === overlay) closeActions(); });
    overlay.querySelector('[data-rubrica-close]')?.addEventListener('click', closeActions);
    overlay.querySelector('[data-rubrica-edit]')?.addEventListener('click', () => openEditor(contact, rerender));
    overlay.querySelector('[data-rubrica-delete]')?.addEventListener('click', () => {
      if (!confirm(`Eliminare ${contact.name} dalla rubrica manuale?`)) return;
      writeManualContacts(readManualContacts().filter((item) => item.id !== contact.id));
      closeActions();
      rerender();
    });
    document.body.appendChild(overlay);
  }

  function openEditor(contact, rerender) {
    closeActions();
    const editing = Boolean(contact?.id);
    const overlay = document.createElement('div');
    overlay.className = 'rubrica-actions-overlay';
    overlay.innerHTML = `<form class="rubrica-actions rubrica-form" data-rubrica-form>
      <h3>${editing ? 'Modifica contatto' : 'Nuovo contatto'}</h3>
      <input name="name" required maxlength="120" placeholder="Nome e cognome" value="${esc(contact?.name)}">
      <input name="phone" required inputmode="tel" maxlength="40" placeholder="Numero di telefono" value="${esc(contact?.phone)}">
      <input name="email" type="email" maxlength="160" placeholder="E-mail (facoltativa)" value="${esc(contact?.email)}">
      <input name="role" maxlength="100" placeholder="Ruolo (facoltativo)" value="${esc(contact?.role)}">
      <button class="rubrica-action primary" type="submit">SALVA</button><button class="rubrica-action" type="button" data-rubrica-close>ANNULLA</button>
    </form>`;
    overlay.querySelector('[data-rubrica-close]').addEventListener('click', closeActions);
    overlay.addEventListener('click', (event) => { if (event.target === overlay) closeActions(); });
    overlay.querySelector('[data-rubrica-form]').addEventListener('submit', (event) => {
      event.preventDefault();
      const data = new FormData(event.currentTarget);
      const row = { id: contact?.id || `manual:${Date.now()}`, name: text(data.get('name')), phone: text(data.get('phone')), email: text(data.get('email')), role: text(data.get('role')), source: 'manual', editable: true };
      if (!normalizePhone(row.phone)) return;
      const rows = readManualContacts();
      const index = rows.findIndex((item) => item.id === row.id);
      if (index >= 0) rows[index] = row; else rows.push(row);
      writeManualContacts(rows);
      closeActions();
      rerender();
    });
    document.body.appendChild(overlay);
    overlay.querySelector('input')?.focus();
  }

  function valueFromRow(row, aliases) {
    const normalized = Object.fromEntries(Object.entries(row || {}).map(([key, value]) => [normalizeHeader(key), value]));
    for (const alias of aliases) {
      const value = text(normalized[normalizeHeader(alias)]);
      if (value) return value;
    }
    return '';
  }

  function mapImportedRow(row, index) {
    const name = valueFromRow(row, ['NOME_COMPLETO', 'NOMINATIVO', 'OPERATORE', 'DIPENDENTE']);
    const firstName = valueFromRow(row, ['NOME', 'FIRST_NAME']);
    const lastName = valueFromRow(row, ['COGNOME', 'LAST_NAME']);
    const phone = valueFromRow(row, ['TELEFONO', 'CELLULARE', 'MOBILE', 'NUMERO_TELEFONO', 'TELEFONO_PERSONALE']);
    if (!normalizePhone(phone)) return null;
    return {
      id: `import:${normalizePhone(phone)}:${index}`,
      name: name || [firstName, lastName].filter(Boolean).join(' ') || `Contatto ${index + 1}`,
      phone,
      email: valueFromRow(row, ['EMAIL', 'E_MAIL', 'MAIL', 'EMAIL_ACCESSO_APP', 'LINKED_USER_EMAIL']),
      role: valueFromRow(row, ['MANSIONE', 'RUOLO', 'QUALIFICA']),
      company: valueFromRow(row, ['AZIENDA', 'DITTA', 'SOCIETA']),
      source: 'manual',
      imported: true,
      editable: true
    };
  }

  async function importPersonnelMatrix(file, rerender) {
    if (!file) return;
    if (!window.XLSX) {
      alert('Lettore Excel non disponibile. Aggiorna l’app e riprova.');
      return;
    }
    try {
      const buffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(buffer, { type: 'array', cellDates: false });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = window.XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
      const imported = rows.map(mapImportedRow).filter(Boolean);
      if (!imported.length) {
        alert('Nessun contatto importabile. La matrice deve contenere una colonna TELEFONO o CELLULARE.');
        return;
      }
      const existing = readManualContacts();
      const byPhone = new Map(existing.map((item) => [normalizePhone(item.phone), item]));
      let added = 0;
      let updated = 0;
      imported.forEach((item) => {
        const key = normalizePhone(item.phone);
        const previous = byPhone.get(key);
        if (previous) {
          byPhone.set(key, { ...previous, ...item, id: previous.id || item.id, source: 'manual', editable: true });
          updated += 1;
        } else {
          byPhone.set(key, item);
          added += 1;
        }
      });
      if (!confirm(`Importare ${imported.length} contatti con numero di telefono?\n\nNuovi: ${added}\nAggiornati: ${updated}\nRighe senza telefono ignorate: ${Math.max(0, rows.length - imported.length)}`)) return;
      writeManualContacts([...byPhone.values()]);
      rerender();
      alert(`Importazione completata.\nNuovi contatti: ${added}\nContatti aggiornati: ${updated}.`);
    } catch (error) {
      console.error('Rubrica: importazione matrice non riuscita.', error);
      alert(`Importazione non riuscita: ${error?.message || 'file non valido'}`);
    }
  }

  function openRubrica() {
    document.getElementById('rubrica-feature-page')?.remove();
    installStyle();
    const page = document.createElement('section');
    page.id = 'rubrica-feature-page';
    page.className = 'rubrica-overlay';
    page.innerHTML = `<div class="rubrica-shell"><header class="rubrica-head"><button class="rubrica-back" type="button">← INDIETRO</button><h2>📒 Rubrica</h2>${isManager() ? '<button class="rubrica-import" type="button">📥 IMPORTA MATRICE PERSONALE</button><button class="rubrica-add" type="button">+ CONTATTO</button><input class="rubrica-import-input" type="file" accept=".xlsx,.xls,.csv" hidden>' : ''}</header><input class="rubrica-search" type="search" placeholder="Cerca nome, numero, ruolo o azienda…" autocomplete="off"><div class="rubrica-list"></div></div>`;
    const list = page.querySelector('.rubrica-list');
    const search = page.querySelector('.rubrica-search');
    const render = () => {
      const query = text(search.value).toLocaleLowerCase('it');
      const contacts = buildContacts().filter((contact) => [contact.name, contact.phone, contact.email, contact.role, contact.company].join(' ').toLocaleLowerCase('it').includes(query));
      list.innerHTML = contacts.length ? contacts.map((contact, index) => {
        const initials = text(contact.name).split(/\s+/).slice(0, 2).map((part) => part[0] || '').join('').toUpperCase() || '👤';
        return `<button class="rubrica-card" type="button" data-contact-index="${index}"><span class="rubrica-avatar">${esc(initials)}</span><span class="rubrica-main"><strong>${esc(contact.name)}</strong><small>${esc(contact.phone)}</small>${contact.role || contact.company ? `<small>${esc([contact.role, contact.company].filter(Boolean).join(' · '))}</small>` : ''}</span><span aria-hidden="true">›</span></button>`;
      }).join('') : '<p class="rubrica-empty">Nessun contatto con numero di telefono.</p>';
      list.querySelectorAll('[data-contact-index]').forEach((button) => button.addEventListener('click', () => openContact(contacts[Number(button.dataset.contactIndex)], render)));
    };
    search.addEventListener('input', render);
    page.querySelector('.rubrica-back').addEventListener('click', () => page.remove());
    page.querySelector('.rubrica-add')?.addEventListener('click', () => openEditor(null, render));
    const fileInput = page.querySelector('.rubrica-import-input');
    page.querySelector('.rubrica-import')?.addEventListener('click', () => fileInput?.click());
    fileInput?.addEventListener('change', async () => {
      const file = fileInput.files?.[0];
      fileInput.value = '';
      await importPersonnelMatrix(file, render);
    });
    document.body.appendChild(page);
    render();
  }

  function bindButton() {
    const current = document.getElementById('today-alerts-btn');
    if (!current || current.dataset.rubricaBound === '1') return false;
    const replacement = current.cloneNode(true);
    replacement.dataset.rubricaBound = '1';
    replacement.removeAttribute('disabled');
    replacement.setAttribute('aria-label', 'Apri Rubrica contatti');
    const icon = replacement.querySelector('.today-summary-icon');
    if (icon) icon.textContent = '📒';
    const strong = replacement.querySelector('strong');
    if (strong) strong.textContent = 'RUBRICA';
    const label = replacement.querySelector('span:last-child');
    if (label && label !== icon) label.textContent = 'Contatti';
    replacement.addEventListener('click', openRubrica);
    current.replaceWith(replacement);
    try {
      if (window.ui?.todayAlertsBtn === current) window.ui.todayAlertsBtn = replacement;
    } catch (_) {}
    return true;
  }

  function init() {
    installStyle();
    if (bindButton()) return;
    const observer = new MutationObserver(() => {
      if (bindButton()) observer.disconnect();
    });
    observer.observe(document.body, { childList: true, subtree: true });
    window.setTimeout(() => observer.disconnect(), 30000);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();
