(() => {
  'use strict';
  if (window.__vargaRubricaMatriceImport) return;
  window.__vargaRubricaMatriceImport = true;

  const text = (value) => String(value ?? '').trim();
  const normalizePhone = (value) => text(value).replace(/[^\d+]/g, '').replace(/^00/, '+');
  const normalizeEmail = (value) => text(value).toLowerCase();
  const normalizeHeader = (value) => text(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '').toUpperCase();
  const normalizeKey = (value) => text(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/gi, ' ').trim().replace(/\s+/g, ' ').toUpperCase();
  const yesNo = (value) => ['SI', 'SÌ', 'YES', 'TRUE', '1', 'X'].includes(normalizeKey(value));
  const isBlank = (value) => value == null || text(value) === '';
  const serverTime = () => window.firebase?.firestore?.FieldValue?.serverTimestamp?.() || new Date().toISOString();

  function valueFromRow(row, aliases) {
    const normalized = Object.fromEntries(Object.entries(row || {}).map(([key, value]) => [normalizeHeader(key), value]));
    for (const alias of aliases) {
      const value = normalized[normalizeHeader(alias)];
      if (!isBlank(value)) return text(value);
    }
    return '';
  }

  function mapRubricaRow(row, index) {
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

  function rubricaKey(row) {
    const phone = normalizePhone(row?.phone);
    if (phone) return `p:${phone}`;
    const email = normalizeEmail(row?.email);
    return email ? `e:${email}` : '';
  }

  async function importRubricaFile(file) {
    if (!file) return;
    if (!window.XLSX) throw new Error('Lettore Excel non disponibile. Aggiorna l’app e riprova.');
    const db = window.firebase?.firestore?.();
    const currentUser = window.firebase?.auth?.()?.currentUser;
    if (!db || !currentUser) throw new Error('Accedi prima di importare la matrice personale.');

    const workbook = window.XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const sourceRows = window.XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
    const imported = sourceRows.map(mapRubricaRow).filter(Boolean);
    if (!imported.length) throw new Error('Nessun contatto importabile. Serve almeno TELEFONO/CELLULARE oppure EMAIL.');

    const collection = db.collection('rubricaContacts');
    const snapshot = await collection.get();
    const existing = new Map();
    snapshot.forEach((doc) => {
      const row = { id: doc.id, ...doc.data() };
      const key = rubricaKey(row);
      if (key) existing.set(key, row);
    });

    let added = 0;
    let updated = 0;
    let ignored = Math.max(0, sourceRows.length - imported.length);
    if (!confirm(`Importare ${imported.length} contatti nella rubrica?\n\nLe righe senza telefono ed e-mail saranno ignorate.`)) return;

    for (const row of imported) {
      const key = rubricaKey(row);
      if (!key) { ignored += 1; continue; }
      const previous = existing.get(key);
      const payload = {
        ...row,
        phone: row.phone || previous?.phone || '',
        email: row.email || previous?.email || '',
        role: row.role || previous?.role || '',
        company: row.company || previous?.company || '',
        updatedBy: currentUser.uid,
        updatedAt: serverTime()
      };
      if (previous?.id) {
        await collection.doc(previous.id).set(payload, { merge: true });
        updated += 1;
      } else {
        await collection.add({ ...payload, createdBy: currentUser.uid, createdAt: serverTime() });
        existing.set(key, payload);
        added += 1;
      }
    }
    alert(`Importazione rubrica completata.\nNuovi: ${added}\nAggiornati: ${updated}\nIgnorati: ${ignored}`);
  }

  function personnelCollection(db) {
    const name = typeof window.getPersonaleCollectionName === 'function' ? window.getPersonaleCollectionName() : 'personale';
    return { name, ref: db.collection(name) };
  }

  function fullName(data = {}) {
    return text(data.nomeCompleto || data.fullName || data.displayName || data.nominativo || data.operatore || [data.nome || data.name || data.firstName, data.cognome || data.surname || data.lastName].filter(Boolean).join(' '));
  }

  function listText(value) {
    if (Array.isArray(value)) return value.map((item) => typeof item === 'object' ? text(item.nome || item.name || item.titolo || item.label) : text(item)).filter(Boolean).join('; ');
    return text(value);
  }

  function exportRow(doc) {
    const data = doc.data() || {};
    return {
      'ID FIRESTORE - NON MODIFICARE': doc.id,
      'ID OPERATORE - NON MODIFICARE': text(data.idOperatore || data.operatorId || ''),
      'NOME': text(data.nome || data.name || data.firstName || ''),
      'COGNOME': text(data.cognome || data.surname || data.lastName || ''),
      'NOME COMPLETO': fullName(data),
      'CODICE FISCALE': text(data.codiceFiscale || data.cf || data.fiscalCode || ''),
      'QUALIFICA': text(data.qualifica || data.mansione || data.role || ''),
      'TELEFONO': text(data.telefono || data.phone || data.cellulare || ''),
      'EMAIL': text(data.email || data.mail || data.linkedUserEmail || ''),
      'DATA DI NASCITA': text(data.dataNascita || data.birthDate || ''),
      'DATA ASSUNZIONE': text(data.dataAssunzione || data.hireDate || ''),
      'TIPO CONTRATTO': text(data.tipoContratto || data.contractType || ''),
      'LIVELLO': text(data.livello || data.level || ''),
      'CORSI': listText(data.corsi || data.formazione || data.training),
      'ABILITAZIONI': listText(data.abilitazioni || data.qualifiche || data.certifications),
      'SCADENZA CORSI': listText(data.scadenzaCorsi || data.scadenzeCorsi || data.trainingExpiries),
      'ACCESSO HERA': data.accessoHera === true ? 'SÌ' : 'NO',
      'TUTTE LE COMMESSE': data.abilitatoTutteCommesse === true ? 'SÌ' : 'NO',
      'NOTE': text(data.note || ''),
      'ATTIVO': data.attivo === false || data.active === false ? 'NO' : 'SÌ'
    };
  }

  async function exportPersonnel() {
    if (!window.XLSX) throw new Error('Libreria Excel non disponibile. Aggiorna l’app e riprova.');
    const db = window.firebase?.firestore?.();
    const user = window.firebase?.auth?.()?.currentUser;
    if (!db || !user) throw new Error('Devi accedere prima di esportare il personale.');
    if (typeof window.canManageData === 'function' && !window.canManageData()) throw new Error('Solo un amministratore può esportare il personale.');

    const { name, ref } = personnelCollection(db);
    const snapshot = typeof window.runFirestoreGetWithRetry === 'function'
      ? await window.runFirestoreGetWithRetry(ref, { label: 'ESPORTAZIONE PERSONALE', timeoutMs: 20000, retries: 2 })
      : await ref.get();
    const rows = snapshot.docs.map(exportRow).sort((a, b) => a['NOME COMPLETO'].localeCompare(b['NOME COMPLETO'], 'it'));
    if (!rows.length) throw new Error('Nessun operatore presente nell’archivio personale.');

    const sheet = window.XLSX.utils.json_to_sheet(rows);
    sheet['!cols'] = Object.keys(rows[0]).map((header) => ({ wch: Math.min(42, Math.max(14, header.length + 2)) }));
    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, sheet, 'PERSONALE');
    const guide = window.XLSX.utils.aoa_to_sheet([
      ['ISTRUZIONI IMPORTAZIONE PERSONALE'],
      ['1', 'Non modificare le due colonne ID.'],
      ['2', 'Compila solo i dati nuovi o corretti.'],
      ['3', 'Le celle vuote non cancellano i dati esistenti.'],
      ['4', 'Nessun operatore viene eliminato o creato automaticamente.'],
      ['5', 'CORSI e ABILITAZIONI: separare le voci con punto e virgola.'],
      ['6', 'ACCESSO HERA e TUTTE LE COMMESSE: usare SÌ oppure NO.'],
      ['Archivio origine', name],
      ['Data esportazione', new Date().toLocaleString('it-IT')]
    ]);
    guide['!cols'] = [{ wch: 24 }, { wch: 85 }];
    window.XLSX.utils.book_append_sheet(workbook, guide, 'ISTRUZIONI');
    const date = new Date().toISOString().slice(0, 10);
    window.XLSX.writeFile(workbook, `personale-varga-${date}.xlsx`);
  }

  function parseList(value) {
    return text(value).split(/[;\n]+/).map((item) => item.trim()).filter(Boolean);
  }

  function mapPersonnelPatch(row) {
    const patch = {};
    const assign = (field, aliases, transform = text) => {
      const value = valueFromRow(row, aliases);
      if (!isBlank(value)) patch[field] = transform(value);
    };
    assign('nome', ['NOME']);
    assign('cognome', ['COGNOME']);
    assign('nomeCompleto', ['NOME COMPLETO', 'NOMINATIVO']);
    assign('codiceFiscale', ['CODICE FISCALE', 'CF'], (value) => normalizeKey(value).replace(/\s/g, ''));
    assign('qualifica', ['QUALIFICA', 'MANSIONE', 'RUOLO']);
    assign('telefono', ['TELEFONO', 'CELLULARE', 'NUMERO DI TELEFONO']);
    assign('email', ['EMAIL', 'E MAIL', 'MAIL'], normalizeEmail);
    assign('dataNascita', ['DATA DI NASCITA']);
    assign('dataAssunzione', ['DATA ASSUNZIONE']);
    assign('tipoContratto', ['TIPO CONTRATTO']);
    assign('livello', ['LIVELLO']);
    assign('corsi', ['CORSI', 'FORMAZIONE'], parseList);
    assign('abilitazioni', ['ABILITAZIONI', 'QUALIFICHE'], parseList);
    assign('scadenzaCorsi', ['SCADENZA CORSI', 'SCADENZE CORSI'], parseList);
    assign('note', ['NOTE']);
    const accessoHera = valueFromRow(row, ['ACCESSO HERA']);
    if (!isBlank(accessoHera)) patch.accessoHera = yesNo(accessoHera);
    const tutte = valueFromRow(row, ['TUTTE LE COMMESSE', 'ABILITATO TUTTE COMMESSE']);
    if (!isBlank(tutte)) patch.abilitatoTutteCommesse = yesNo(tutte);
    const attivo = valueFromRow(row, ['ATTIVO', 'STATO ATTIVO']);
    if (!isBlank(attivo)) patch.attivo = yesNo(attivo);
    return patch;
  }

  function buildIndexes(snapshot) {
    const indexes = { byDocId: new Map(), byOperatorId: new Map(), byCf: new Map(), byEmail: new Map(), byName: new Map() };
    const add = (map, key, item) => {
      if (!key) return;
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(item);
    };
    snapshot.docs.forEach((doc) => {
      const data = doc.data() || {};
      const item = { doc, data };
      indexes.byDocId.set(doc.id, [item]);
      add(indexes.byOperatorId, normalizeKey(data.idOperatore || data.operatorId), item);
      add(indexes.byCf, normalizeKey(data.codiceFiscale || data.cf || data.fiscalCode).replace(/\s/g, ''), item);
      add(indexes.byEmail, normalizeEmail(data.email || data.mail || data.linkedUserEmail), item);
      add(indexes.byName, normalizeKey(fullName(data)), item);
    });
    return indexes;
  }

  function uniqueMatch(row, indexes) {
    const candidates = [
      ['ID FIRESTORE', valueFromRow(row, ['ID FIRESTORE - NON MODIFICARE', 'ID FIRESTORE', 'DOCUMENT ID']), indexes.byDocId],
      ['ID OPERATORE', valueFromRow(row, ['ID OPERATORE - NON MODIFICARE', 'ID OPERATORE']), indexes.byOperatorId],
      ['CODICE FISCALE', valueFromRow(row, ['CODICE FISCALE', 'CF']).replace(/\s/g, ''), indexes.byCf],
      ['EMAIL', normalizeEmail(valueFromRow(row, ['EMAIL', 'MAIL'])), indexes.byEmail],
      ['NOME COMPLETO', normalizeKey(valueFromRow(row, ['NOME COMPLETO', 'NOMINATIVO']) || [valueFromRow(row, ['NOME']), valueFromRow(row, ['COGNOME'])].filter(Boolean).join(' ')), indexes.byName]
    ];
    for (const [method, raw, index] of candidates) {
      const key = method === 'ID FIRESTORE' ? text(raw) : normalizeKey(raw);
      if (!key) continue;
      const matches = index.get(key) || [];
      if (matches.length === 1) return { status: 'matched', method, item: matches[0] };
      if (matches.length > 1) return { status: 'ambiguous', method, matches };
    }
    return { status: 'missing' };
  }

  function changedPatch(previous, patch) {
    const changed = {};
    Object.entries(patch).forEach(([key, value]) => {
      if (Array.isArray(value)) {
        const before = JSON.stringify(Array.isArray(previous[key]) ? previous[key] : parseList(previous[key]));
        if (before !== JSON.stringify(value)) changed[key] = value;
      } else if (String(previous[key] ?? '') !== String(value ?? '')) changed[key] = value;
    });
    return changed;
  }

  async function importPersonnel(file) {
    if (!file) return;
    if (!window.XLSX) throw new Error('Libreria Excel non disponibile. Aggiorna l’app e riprova.');
    const db = window.firebase?.firestore?.();
    const user = window.firebase?.auth?.()?.currentUser;
    if (!db || !user) throw new Error('Devi accedere prima di importare il personale.');
    if (typeof window.canManageData === 'function' && !window.canManageData()) throw new Error('Solo un amministratore può importare il personale.');

    const workbook = window.XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
    const sheet = workbook.Sheets.PERSONALE || workbook.Sheets[workbook.SheetNames[0]];
    const rows = window.XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
    if (!rows.length) throw new Error('Il foglio PERSONALE è vuoto.');

    const { name, ref } = personnelCollection(db);
    const snapshot = typeof window.runFirestoreGetWithRetry === 'function'
      ? await window.runFirestoreGetWithRetry(ref, { label: 'ANTEPRIMA IMPORT PERSONALE', timeoutMs: 20000, retries: 2 })
      : await ref.get();
    const indexes = buildIndexes(snapshot);
    const plan = [];
    let unchanged = 0;
    let missing = 0;
    let ambiguous = 0;
    let invalid = 0;

    rows.forEach((row, index) => {
      const patch = mapPersonnelPatch(row);
      if (!Object.keys(patch).length) { invalid += 1; return; }
      const match = uniqueMatch(row, indexes);
      if (match.status === 'missing') { missing += 1; return; }
      if (match.status === 'ambiguous') { ambiguous += 1; return; }
      const changes = changedPatch(match.item.data, patch);
      if (!Object.keys(changes).length) { unchanged += 1; return; }
      plan.push({ rowNumber: index + 2, doc: match.item.doc, method: match.method, changes });
    });

    const summary = `ANTEPRIMA IMPORTAZIONE PERSONALE\n\nArchivio: ${name}\nRighe nel file: ${rows.length}\nOperatori da aggiornare: ${plan.length}\nSenza modifiche: ${unchanged}\nNon trovati e NON creati: ${missing}\nCorrispondenze ambigue: ${ambiguous}\nRighe vuote/non valide: ${invalid}\n\nNessun operatore sarà eliminato o duplicato. Le celle vuote non cancellano dati.`;
    if (!plan.length) {
      alert(summary + '\n\nNon ci sono aggiornamenti da salvare.');
      return;
    }
    if (!confirm(summary + '\n\nConfermi l’aggiornamento?')) return;

    for (let start = 0; start < plan.length; start += 400) {
      const batch = db.batch();
      plan.slice(start, start + 400).forEach(({ doc, method, changes }) => {
        batch.set(doc.ref, {
          ...changes,
          ultimoImportPersonaleAt: serverTime(),
          ultimoImportPersonaleDa: user.uid,
          ultimoImportPersonaleEmail: user.email || '',
          ultimoImportPersonaleMetodo: method,
          importPersonaleVersione: 1
        }, { merge: true });
      });
      await batch.commit();
    }

    alert(`Personale aggiornato correttamente.\n\nOperatori aggiornati: ${plan.length}\nOperatori creati: 0\nOperatori eliminati: 0\nDuplicati creati: 0`);
    window.dispatchEvent(new CustomEvent('varga-personale-import-complete', { detail: { updated: plan.length, collection: name } }));
  }

  function makeButton(label, datasetName, handler) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'btn';
    button.dataset[datasetName] = '1';
    button.textContent = label;
    button.addEventListener('click', async () => {
      if (button.disabled) return;
      button.disabled = true;
      const original = button.textContent;
      try {
        button.textContent = 'Operazione in corso…';
        await handler();
      } catch (error) {
        console.error(label, error);
        alert(error?.message || 'Operazione non riuscita.');
      } finally {
        button.disabled = false;
        button.textContent = original;
      }
    });
    return button;
  }

  function installPersonnelButtons() {
    const panel = document.getElementById('panel-personale');
    if (!panel || panel.querySelector('[data-export-personale-excel]')) return false;
    const host = panel.querySelector('.item-actions, .management-actions, form, h3')?.parentElement || panel;
    const actions = document.createElement('div');
    actions.className = 'item-actions personale-excel-actions';

    const exportButton = makeButton('📤 ESPORTA PERSONALE EXCEL', 'exportPersonaleExcel', exportPersonnel);
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls';
    input.hidden = true;
    input.dataset.importPersonaleExcelInput = '1';
    const importButton = makeButton('📥 IMPORTA AGGIORNAMENTI PERSONALE', 'importPersonaleExcel', async () => input.click());
    input.addEventListener('change', async () => {
      const file = input.files?.[0];
      if (!file) return;
      importButton.disabled = true;
      const original = importButton.textContent;
      importButton.textContent = 'Analisi del file…';
      try { await importPersonnel(file); }
      catch (error) { console.error('Import personale non riuscito', error); alert(error?.message || 'Importazione non riuscita.'); }
      finally { input.value = ''; importButton.disabled = false; importButton.textContent = original; }
    });

    const note = document.createElement('p');
    note.className = 'muted';
    note.textContent = 'Esporta l’elenco, compila corsi, abilitazioni, telefono, email e accessi, poi ricarica lo stesso file. L’importazione aggiorna solo gli operatori esistenti e non cancella dati.';
    actions.append(exportButton, importButton, input);
    host.prepend(note);
    host.prepend(actions);
    return true;
  }

  function installRubricaButton() {
    const page = document.querySelector('.rcv3');
    const actions = page?.querySelector('.rcv3-actions');
    if (!actions || actions.querySelector('[data-import-matrice-personale]')) return false;
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls,.csv';
    input.hidden = true;
    const button = makeButton('📥 IMPORTA MATRICE PERSONALE', 'importMatricePersonale', async () => input.click());
    input.addEventListener('change', async () => {
      const file = input.files?.[0];
      if (!file) return;
      button.disabled = true;
      const original = button.textContent;
      button.textContent = 'Importazione…';
      try { await importRubricaFile(file); }
      catch (error) { console.error('Importazione matrice personale non riuscita:', error); alert(error?.message || 'Importazione non riuscita.'); }
      finally { input.value = ''; button.disabled = false; button.textContent = original; }
    });
    actions.prepend(button, input);
    return true;
  }

  function init() {
    installPersonnelButtons();
    installRubricaButton();
    const observer = new MutationObserver(() => {
      installPersonnelButtons();
      installRubricaButton();
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  window.HeraPersonnelExcel = { exportPersonnel, importPersonnel };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();
