(() => {
  'use strict';
  if (window.HeraOccasionalDocumentsRuntime?.installed) return;

  const MAX_FILE_SIZE = 15 * 1024 * 1024;
  const SOURCE = 'documentazione-cantiere';
  let listObserver = null;
  let observedList = null;
  let overlay = null;

  const text = (value) => String(value ?? '').trim();
  const upper = (value) => text(value).replace(/\s+/g, ' ').toLocaleUpperCase('it-IT');
  const esc = (value) => text(value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');

  function isAdmin() {
    try { if (typeof canManageData === 'function' && canManageData()) return true; } catch (_) {}
    try { if (typeof window.canManageData === 'function' && window.canManageData()) return true; } catch (_) {}
    return false;
  }

  function currentCommessaId() {
    try { if (typeof selectedCommessaId !== 'undefined' && text(selectedCommessaId)) return text(selectedCommessaId); } catch (_) {}
    return text(window.selectedCommessaId || window.currentCommessaId);
  }

  function currentCommessa() {
    const id = currentCommessaId();
    try {
      if (typeof commesseById !== 'undefined' && commesseById instanceof Map && commesseById.has(id)) return commesseById.get(id);
    } catch (_) {}
    try {
      if (window.commesseById instanceof Map && window.commesseById.has(id)) return window.commesseById.get(id);
    } catch (_) {}
    return null;
  }

  function isOccasionalCommessa() {
    const id = upper(currentCommessaId());
    if (id === 'LAVORI-OCCASIONALI') return true;
    const commessa = currentCommessa();
    const meta = upper([commessa?.nome, commessa?.name, commessa?.codice, commessa?.code].filter(Boolean).join(' '));
    const header = upper(document.getElementById('commessa-focus-label')?.textContent);
    return meta.includes('OCCASIONAL') || header.includes('OCCASIONAL');
  }

  function getPlants() {
    try { if (Array.isArray(currentImpianti)) return currentImpianti; } catch (_) {}
    return Array.isArray(window.currentImpianti) ? window.currentImpianti : [];
  }

  function plantId(plant) {
    return text(plant?.id || plant?.docId || plant?.impiantoId || plant?.physicalPlantId || plant?.idSap || plant?.idSAP || plant?.['ID SAP']);
  }

  function plantName(plant) {
    return text(plant?.denominazione || plant?.nome || plant?.impianto || plant?.['Denominazione Impianto'] || 'Cantiere');
  }

  function plantCommessaId(plant) {
    return text(plant?.commessaId || plant?.parentCommessaId || currentCommessaId());
  }

  function plantKey(plant) {
    const cid = plantCommessaId(plant);
    const pid = plantId(plant);
    return cid && pid ? `${cid}::${pid}` : '';
  }

  function isOccasionalPlant(plant) {
    return plant?.lavoroOccasionale === true || isOccasionalCommessa();
  }

  function findPlantForCard(card) {
    if (!card) return null;
    const body = upper(card.textContent);
    return getPlants().filter(isOccasionalPlant).find((plant) => {
      const name = upper(plantName(plant));
      const id = upper(plantId(plant));
      return (name && body.includes(name)) || (id && body.includes(id));
    }) || null;
  }

  function cardFromStack(stack) {
    let node = stack?.parentElement || null;
    const list = document.getElementById('impianti-lista');
    while (node && node !== list && node !== document.body) {
      if (node.querySelector?.('.impianto-primary-actions [data-action-key="navigate"]')) return node;
      node = node.parentElement;
    }
    return null;
  }

  function ensureStyles() {
    if (document.getElementById('hera-occasional-doc-runtime-style')) return;
    const style = document.createElement('style');
    style.id = 'hera-occasional-doc-runtime-style';
    style.textContent = `
      .occasional-doc-runtime-btn{margin-top:6px;width:100%;border:1px solid #2563eb;background:#eff6ff;color:#174ea6;border-radius:10px;padding:9px 11px;font-weight:850}
      .occasional-doc-overlay{position:fixed;inset:0;z-index:2147482500;background:rgba(15,23,42,.68);display:flex;align-items:flex-end;justify-content:center}
      .occasional-doc-panel{width:min(720px,100%);max-height:92dvh;background:#f8fafc;border-radius:20px 20px 0 0;display:flex;flex-direction:column;overflow:hidden}
      .occasional-doc-head{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;padding:14px 16px;background:#fff;border-bottom:1px solid #e2e8f0}
      .occasional-doc-head h2{margin:0;font-size:19px}.occasional-doc-head p{margin:3px 0 0;color:#64748b;font-size:12px}
      .occasional-doc-close{border:0;border-radius:999px;width:36px;height:36px;font-size:18px}.occasional-doc-body{padding:14px;overflow:auto}
      .occasional-doc-form{display:grid;gap:9px;background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:12px;margin-bottom:12px}
      .occasional-doc-form input,.occasional-doc-form select,.occasional-doc-form textarea{width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:10px;padding:9px;font:inherit}
      .occasional-doc-save{border:0;background:#16a34a;color:#fff;border-radius:10px;padding:10px 12px;font-weight:850}.occasional-doc-status{font-size:12px;color:#64748b}
      .occasional-doc-list{display:grid;gap:9px}.occasional-doc-item{background:#fff;border:1px solid #e2e8f0;border-radius:13px;padding:10px;display:flex;justify-content:space-between;gap:10px;align-items:center}
      .occasional-doc-item strong{display:block}.occasional-doc-meta{font-size:11px;color:#64748b;margin-top:3px}.occasional-doc-open{border:0;background:#e8f1ff;color:#165bb6;border-radius:9px;padding:8px 10px;font-weight:800}
    `;
    document.head.appendChild(style);
  }

  function closeOverlay() {
    overlay?.remove();
    overlay = null;
  }

  function userInfo() {
    let user = null;
    try { user = firebase.auth?.().currentUser || null; } catch (_) {}
    return {
      uid: text(user?.uid),
      email: text(user?.email),
      name: text(user?.displayName || user?.email || 'Operatore')
    };
  }

  async function loadDocs(plant) {
    const key = plantKey(plant);
    if (!key || typeof db === 'undefined') return [];
    try {
      const snap = await db.collection('documents').where('impiantoKey', '==', key).get();
      return snap.docs.map((doc) => ({ id: doc.id, ...doc.data() }))
        .filter((item) => item.source === SOURCE && text(item.fileUrl));
    } catch (error) {
      console.warn('[DOC OCCASIONALI] lettura fallita', error);
      return [];
    }
  }

  function renderDocs(host, items) {
    if (!items.length) {
      host.innerHTML = '<div class="occasional-doc-status">Nessuna documentazione allegata.</div>';
      return;
    }
    host.innerHTML = `<div class="occasional-doc-list">${items.map((item) => `
      <article class="occasional-doc-item">
        <div><strong>${esc(item.title || item.fileName || 'Documento')}</strong><div class="occasional-doc-meta">${esc(item.category || 'Documento')}</div></div>
        <button type="button" class="occasional-doc-open" data-url="${esc(item.fileUrl)}">APRI</button>
      </article>`).join('')}</div>`;
    host.querySelectorAll('.occasional-doc-open').forEach((button) => button.addEventListener('click', () => {
      const url = button.dataset.url;
      if (url) window.open(url, '_blank', 'noopener,noreferrer');
    }));
  }

  async function refreshList(plant, host) {
    host.innerHTML = '<div class="occasional-doc-status">Caricamento documenti…</div>';
    renderDocs(host, await loadDocs(plant));
  }

  async function uploadFile(plant, file, values) {
    if (!isAdmin()) throw new Error('Operazione riservata all’amministratore.');
    if (!(file instanceof Blob) || !file.size) throw new Error('Seleziona un file.');
    if (file.size > MAX_FILE_SIZE) throw new Error('Il file supera 15 MB.');
    const user = userInfo();
    if (!user.uid) throw new Error('Utente non autenticato.');
    if (typeof firebase?.storage !== 'function') throw new Error('Firebase Storage non disponibile.');
    const docRef = db.collection('documents').doc();
    const safe = text(file.name || `documento-${Date.now()}`).replace(/[\\/:*?"<>|#%{}\[\]]+/g, '-').slice(-120);
    const path = `documents/${user.uid}/${docRef.id}/${safe}`;
    const storageRef = firebase.storage().ref(path);
    await storageRef.put(file, { contentType: text(file.type) || 'application/octet-stream' });
    const fileUrl = await storageRef.getDownloadURL();
    const cid = plantCommessaId(plant);
    await docRef.set({
      id: docRef.id,
      source: SOURCE,
      ownerUserId: user.uid,
      createdBy: user.uid,
      createdByEmail: user.email,
      createdByName: user.name,
      visibility: 'global',
      sharedToAll: true,
      commessaIds: [cid],
      commessaId: cid,
      impiantoId: plantId(plant),
      impiantoKey: plantKey(plant),
      impiantoName: plantName(plant),
      title: text(values.title || file.name || safe),
      category: text(values.category || 'ALTRO'),
      note: text(values.note),
      importantBeforeNavigation: values.importantBeforeNavigation === true,
      showBeforeNavigation: true,
      fileName: safe,
      fileUrl,
      storagePath: path,
      mimeType: text(file.type) || 'application/octet-stream',
      fileSize: Number(file.size),
      uploadStatus: 'completed',
      createdAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    });
  }

  async function openPanel(plant) {
    if (!plant) return;
    ensureStyles();
    closeOverlay();
    overlay = document.createElement('section');
    overlay.className = 'occasional-doc-overlay';
    overlay.innerHTML = `<div class="occasional-doc-panel"><header class="occasional-doc-head"><div><h2>📁 Documentazione cantiere</h2><p>${esc(plantName(plant))}</p></div><button class="occasional-doc-close" type="button">✕</button></header><main class="occasional-doc-body"><form class="occasional-doc-form"><input name="title" type="text" placeholder="Titolo documento"><select name="category"><option value="PREVENTIVO">Preventivo</option><option value="FOTO_PRIMA">Foto prima lavori</option><option value="FOTO_DOPO">Foto dopo lavori</option><option value="PLANIMETRIA">Planimetria</option><option value="SICUREZZA">POS / Sicurezza</option><option value="ORDINE_LAVORO">Ordine di lavoro</option><option value="AUTORIZZAZIONE">Autorizzazione</option><option value="VERBALE">Verbale</option><option value="CONSUNTIVO">Consuntivo / Fattura</option><option value="ALTRO" selected>Altro</option></select><textarea name="note" rows="2" placeholder="Nota facoltativa"></textarea><label><input name="important" type="checkbox"> Da leggere prima di NAVIGA</label><input name="file" type="file" accept=".pdf,image/*,.doc,.docx,.xls,.xlsx,.txt" required><button class="occasional-doc-save" type="submit">CARICA DOCUMENTO</button><div class="occasional-doc-status" data-status></div></form><div data-list></div></main></div>`;
    document.body.appendChild(overlay);
    overlay.querySelector('.occasional-doc-close')?.addEventListener('click', closeOverlay);
    overlay.addEventListener('click', (event) => { if (event.target === overlay) closeOverlay(); });
    const form = overlay.querySelector('.occasional-doc-form');
    const list = overlay.querySelector('[data-list]');
    const status = overlay.querySelector('[data-status]');
    await refreshList(plant, list);
    form.addEventListener('submit', async (event) => {
      event.preventDefault();
      const file = form.elements.file.files?.[0];
      status.textContent = 'Caricamento…';
      try {
        await uploadFile(plant, file, {
          title: form.elements.title.value,
          category: form.elements.category.value,
          note: form.elements.note.value,
          importantBeforeNavigation: form.elements.important.checked
        });
        form.reset();
        status.textContent = 'Documento caricato.';
        await refreshList(plant, list);
      } catch (error) {
        console.error('[DOC OCCASIONALI] upload fallito', error);
        status.textContent = `Errore: ${error?.message || 'caricamento non riuscito'}`;
      }
    });
  }

  function addButton(stack) {
    if (!(stack instanceof HTMLElement) || !isAdmin() || !isOccasionalCommessa()) return;
    const managementActions = stack.querySelector('.item-actions-gestione');
    if (!managementActions || managementActions.querySelector('[data-occasional-doc-runtime]')) return;
    const plant = findPlantForCard(cardFromStack(stack));
    if (!plant) return;
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'btn occasional-doc-runtime-btn';
    button.dataset.occasionalDocRuntime = '1';
    button.textContent = '📁 ALLEGA DOCUMENTAZIONE';
    button.addEventListener('click', (event) => {
      event.preventDefault();
      event.stopPropagation();
      if (plantCommessaId(plant) === 'lavori-occasionali' && window.HeraCantiereDocuments?.open) {
        window.HeraCantiereDocuments.open(plant);
      } else {
        void openPanel(plant);
      }
    });
    managementActions.appendChild(button);
  }

  function decorate(list = document.getElementById('impianti-lista')) {
    if (!list) return;
    list.querySelectorAll('.impianto-management-stack').forEach(addButton);
  }

  function bindList() {
    const list = document.getElementById('impianti-lista');
    if (!list) return false;
    if (list === observedList) { decorate(list); return true; }
    listObserver?.disconnect();
    observedList = list;
    decorate(list);
    listObserver = new MutationObserver((mutations) => {
      for (const mutation of mutations) {
        for (const node of mutation.addedNodes) {
          if (!(node instanceof HTMLElement)) continue;
          if (node.matches?.('.impianto-management-stack')) addButton(node);
          node.querySelectorAll?.('.impianto-management-stack').forEach(addButton);
        }
      }
    });
    listObserver.observe(list, { childList: true, subtree: true });
    return true;
  }

  function install() {
    ensureStyles();
    bindList();
    let attempts = 0;
    const timer = setInterval(() => {
      attempts += 1;
      if (bindList() || attempts >= 20) clearInterval(timer);
    }, 250);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install, { once: true }); else install();
  window.HeraOccasionalDocumentsRuntime = { installed: true, refresh: bindList };
})();