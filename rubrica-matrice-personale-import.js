(() => {
  'use strict';
  if (window.__vargaRubricaMatriceImport) return;
  window.__vargaRubricaMatriceImport = true;

  const text = (value) => String(value ?? '').trim();
  const normalizePhone = (value) => text(value).replace(/[^\d+]/g, '').replace(/^00/, '+');
  const normalizeEmail = (value) => text(value).toLowerCase();
  const normalizeHeader = (value) => text(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '').toUpperCase();

  function valueFromRow(row, aliases) {
    const normalized = Object.fromEntries(Object.entries(row || {}).map(([key, value]) => [normalizeHeader(key), value]));
    return aliases.map((alias) => text(normalized[normalizeHeader(alias)])).find(Boolean) || '';
  }

  function mapRow(row, index) {
    const phone = valueFromRow(row, ['TELEFONO', 'CELLULARE', 'MOBILE', 'NUMERO_TELEFONO', 'TELEFONO_PERSONALE']);
    const email = normalizeEmail(valueFromRow(row, ['EMAIL', 'E_MAIL', 'MAIL', 'EMAIL_ACCESSO_APP', 'LINKED_USER_EMAIL']));
    if (!normalizePhone(phone) && !email) return null;
    const fullName = valueFromRow(row, ['NOME_COMPLETO', 'NOMINATIVO', 'OPERATORE', 'DIPENDENTE']);
    const name = fullName || [valueFromRow(row, ['NOME', 'FIRST_NAME']), valueFromRow(row, ['COGNOME', 'LAST_NAME'])].filter(Boolean).join(' ') || email || phone || `Contatto ${index + 1}`;
    return {
      name,
      phone,
      email,
      role: valueFromRow(row, ['MANSIONE', 'RUOLO', 'QUALIFICA']),
      company: valueFromRow(row, ['AZIENDA', 'DITTA', 'SOCIETA']),
      source: 'matrice-personale',
      recordType: 'rubrica'
    };
  }

  function keyOf(row) {
    const phone = normalizePhone(row?.phone);
    if (phone) return `p:${phone}`;
    const email = normalizeEmail(row?.email);
    return email ? `e:${email}` : '';
  }

  async function importFile(file) {
    if (!file) return;
    if (!window.XLSX) throw new Error('Lettore Excel non disponibile. Aggiorna l’app e riprova.');
    const db = window.firebase?.firestore?.();
    const currentUser = window.firebase?.auth?.()?.currentUser;
    if (!db || !currentUser) throw new Error('Accedi prima di importare la matrice personale.');

    const workbook = window.XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const sourceRows = window.XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
    const imported = sourceRows.map(mapRow).filter(Boolean);
    if (!imported.length) throw new Error('Nessun contatto importabile. Serve almeno TELEFONO/CELLULARE oppure EMAIL.');

    const collection = db.collection('rubricaContacts');
    const snapshot = await collection.get();
    const existing = new Map();
    snapshot.forEach((doc) => {
      const row = { id: doc.id, ...doc.data() };
      const key = keyOf(row);
      if (key) existing.set(key, row);
    });

    let added = 0;
    let updated = 0;
    let ignored = Math.max(0, sourceRows.length - imported.length);
    const serverTimestamp = window.firebase?.firestore?.FieldValue?.serverTimestamp?.() || new Date().toISOString();

    if (!confirm(`Importare ${imported.length} contatti dalla matrice personale?\n\nLe righe senza telefono ed e-mail saranno ignorate.`)) return;

    for (const row of imported) {
      const key = keyOf(row);
      if (!key) { ignored += 1; continue; }
      const previous = existing.get(key);
      const payload = {
        ...previous,
        ...row,
        phone: row.phone || previous?.phone || '',
        email: row.email || previous?.email || '',
        role: row.role || previous?.role || '',
        company: row.company || previous?.company || '',
        updatedBy: currentUser.uid,
        updatedAt: serverTimestamp
      };
      if (previous?.id) {
        delete payload.id;
        await collection.doc(previous.id).set(payload, { merge: true });
        updated += 1;
      } else {
        await collection.add({
          ...payload,
          createdBy: currentUser.uid,
          createdAt: serverTimestamp
        });
        existing.set(key, payload);
        added += 1;
      }
    }

    alert(`Importazione completata.\nNuovi: ${added}\nAggiornati: ${updated}\nIgnorati: ${ignored}`);
  }

  function installButton() {
    const page = document.querySelector('.rcv3');
    const actions = page?.querySelector('.rcv3-actions');
    if (!actions || actions.querySelector('[data-import-matrice-personale]')) return false;

    const button = document.createElement('button');
    button.type = 'button';
    button.dataset.importMatricePersonale = '1';
    button.textContent = '📥 IMPORTA MATRICE PERSONALE';

    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls,.csv';
    input.hidden = true;

    button.addEventListener('click', () => input.click());
    input.addEventListener('change', async () => {
      const file = input.files?.[0];
      if (!file) return;
      button.disabled = true;
      const original = button.textContent;
      button.textContent = 'Importazione…';
      try {
        await importFile(file);
      } catch (error) {
        console.error('Importazione matrice personale non riuscita:', error);
        alert(error?.message || 'Importazione non riuscita.');
      } finally {
        input.value = '';
        button.disabled = false;
        button.textContent = original;
      }
    });

    actions.prepend(button, input);
    return true;
  }

  function init() {
    installButton();
    const observer = new MutationObserver(() => installButton());
    observer.observe(document.body, { childList: true, subtree: true });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();