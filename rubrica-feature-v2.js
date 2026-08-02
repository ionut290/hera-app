(() => {
  'use strict';

  if (window.__vargaRubricaFeatureV2Loaded) return;
  window.__vargaRubricaFeatureV2Loaded = true;

  const STORAGE_KEY = 'varga_rubrica_contatti_manual_v1';
  const text = (value) => String(value ?? '').trim();
  const esc = (value) => text(value).replace(/[&<>"']/g, (char) => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[char]));
  const normalizePhone = (value) => text(value).replace(/[^\d+]/g, '').replace(/^00/, '+');
  const normalizeEmail = (value) => text(value).toLowerCase();
  const normalizeHeader = (value) => text(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '').toUpperCase();
  const phoneFields = ['telefono','phone','cellulare','mobile','numeroTelefono','telefonoPersonale'];
  const emailFields = ['email','emailAccessoApp','linkedUserEmail','mail'];
  const firstValue = (object, fields) => fields.map((field) => text(object?.[field])).find(Boolean) || '';

  function readManual() {
    try { const rows = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]'); return Array.isArray(rows) ? rows : []; }
    catch (_) { return []; }
  }
  function writeManual(rows) { localStorage.setItem(STORAGE_KEY, JSON.stringify(rows)); }
  function personnel() { return Array.isArray(window.personaleRecords) ? window.personaleRecords : []; }
  function appUsers() { return Array.isArray(window.platformUsers) ? window.platformUsers : []; }
  function isManager() { try { return typeof window.canManageData === 'function' && window.canManageData(); } catch (_) { return false; } }

  function displayName(person) {
    try {
      const resolved = typeof window.getPersonaleDisplayName === 'function' ? text(window.getPersonaleDisplayName(person)) : '';
      if (resolved) return resolved;
    } catch (_) {}
    return text(person?.displayName || person?.nomeCompleto || [person?.nome, person?.cognome].map(text).filter(Boolean).join(' ') || person?.name);
  }

  function linkedUser(person) {
    const uid = text(person?.linkedUserId || person?.userId || person?.uid);
    const email = normalizeEmail(firstValue(person, emailFields));
    return appUsers().find((user) => uid && text(user?.id || user?.uid) === uid)
      || appUsers().find((user) => email && normalizeEmail(user?.email) === email)
      || null;
  }

  function contactKey(contact) {
    const phone = normalizePhone(contact?.phone);
    if (phone) return `p:${phone}`;
    const email = normalizeEmail(contact?.email);
    return email ? `e:${email}` : '';
  }

  function buildContacts() {
    const contacts = new Map();
    personnel().forEach((person) => {
      const user = linkedUser(person);
      const phone = firstValue(person, phoneFields);
      const email = firstValue(person, emailFields) || text(user?.email);
      if (!normalizePhone(phone) && !normalizeEmail(email)) return;
      const row = {
        id: `operator:${text(person?.id || contactKey({ phone, email }))}`,
        source: 'operator', editable: false,
        name: displayName(person) || text(user?.displayName) || email || phone || 'Contatto',
        phone, email,
        role: text(person?.mansione || person?.ruolo || person?.role),
        company: text(person?.azienda || person?.company)
      };
      const key = contactKey(row);
      const previous = contacts.get(key);
      contacts.set(key, previous ? { ...previous, ...row, phone: row.phone || previous.phone, email: row.email || previous.email } : row);
    });
    readManual().forEach((row) => {
      if (!normalizePhone(row?.phone) && !normalizeEmail(row?.email)) return;
      const normalized = { ...row, id: row.id || `manual:${Date.now()}`, source:'manual', editable:true };
      contacts.set(contactKey(normalized), normalized);
    });
    return [...contacts.values()].sort((a,b) => text(a.name).localeCompare(text(b.name), 'it', { sensitivity:'base' }));
  }

  function installStyle() {
    if (document.getElementById('rubrica-feature-v2-style')) return;
    const style = document.createElement('style');
    style.id = 'rubrica-feature-v2-style';
    style.textContent = `
      .rubrica-v2-page{position:fixed;inset:0;z-index:13000;background:#eef4f2;overflow:auto;padding:env(safe-area-inset-top) 14px calc(20px + env(safe-area-inset-bottom))}
      .rubrica-v2-shell{width:min(760px,100%);margin:auto;padding:12px 0 28px}.rubrica-v2-head{position:sticky;top:0;z-index:2;background:#eef4f2;padding:8px 0 12px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}.rubrica-v2-head h2{margin:0;flex:1;color:#12352f}.rubrica-v2-head button{min-height:44px;border:1px solid #bfd0cb;border-radius:12px;background:#fff;padding:8px 11px;font-weight:800}
      .rubrica-v2-search{width:100%;min-height:48px;border:1px solid #bfd0cb;border-radius:14px;padding:0 14px;font-size:16px;background:#fff;margin-bottom:12px}.rubrica-v2-list{display:grid;gap:10px}.rubrica-v2-card{width:100%;border:1px solid #d3dedb;border-radius:18px;background:#fff;padding:14px;text-align:left;display:flex;gap:12px;align-items:center}.rubrica-v2-avatar{width:46px;height:46px;border-radius:50%;display:grid;place-items:center;background:#e5f2ee;font-weight:900;color:#164d42;flex:none}.rubrica-v2-main{min-width:0;flex:1}.rubrica-v2-main strong,.rubrica-v2-main small{display:block}.rubrica-v2-main small{color:#64736f;margin-top:3px;overflow:hidden;text-overflow:ellipsis}.rubrica-v2-empty{text-align:center;padding:24px;color:#667773}
      .rubrica-v2-modal{position:fixed;inset:0;z-index:13010;background:rgba(6,18,16,.55);display:flex;align-items:flex-end;justify-content:center;padding:14px}.rubrica-v2-dialog{width:min(520px,100%);background:#fff;border-radius:20px;padding:18px;display:grid;gap:10px}.rubrica-v2-dialog h3{margin:0}.rubrica-v2-action{min-height:48px;border:0;border-radius:12px;font-weight:900;font-size:1rem;text-decoration:none;display:flex;align-items:center;justify-content:center;background:#edf4f2;color:#143b33}.rubrica-v2-primary{background:#13795b;color:#fff}.rubrica-v2-danger{background:#fff0f0;color:#9f1d1d}.rubrica-v2-dialog input{min-height:46px;border:1px solid #c7d6d2;border-radius:11px;padding:0 12px;font-size:16px}
      @media(max-width:620px){.rubrica-v2-head h2{flex-basis:calc(100% - 120px)}.rubrica-v2-import,.rubrica-v2-add{flex:1}}@media(min-width:700px){.rubrica-v2-modal{align-items:center}}
    `;
    document.head.appendChild(style);
  }

  function closeModal() { document.querySelector('.rubrica-v2-modal')?.remove(); }
  function whatsappNumber(phone) {
    let value = normalizePhone(phone).replace(/\D/g, '');
    if (value.startsWith('0')) value = `39${value.slice(1)}`;
    if (value && !value.startsWith('39') && value.length <= 10) value = `39${value}`;
    return value;
  }

  function openContact(contact, rerender) {
    closeModal();
    const modal = document.createElement('div');
    modal.className = 'rubrica-v2-modal';
    const phone = normalizePhone(contact.phone);
    const email = normalizeEmail(contact.email);
    modal.innerHTML = `<section class="rubrica-v2-dialog" role="dialog" aria-modal="true"><h3>${esc(contact.name)}</h3><p>${phone ? esc(contact.phone) : ''}${phone && email ? ' · ' : ''}${email ? esc(contact.email) : ''}</p>${phone ? `<a class="rubrica-v2-action rubrica-v2-primary" href="tel:${esc(phone)}">📞 CHIAMA</a><a class="rubrica-v2-action" href="https://wa.me/${esc(whatsappNumber(phone))}" target="_blank" rel="noopener">💬 WHATSAPP</a>` : ''}${email ? `<a class="rubrica-v2-action${phone ? '' : ' rubrica-v2-primary'}" href="mailto:${esc(email)}">✉️ E-MAIL</a>` : ''}${contact.editable && isManager() ? '<button class="rubrica-v2-action" data-edit>✏️ MODIFICA</button><button class="rubrica-v2-action rubrica-v2-danger" data-delete>🗑️ ELIMINA</button>' : ''}<button class="rubrica-v2-action" data-close>CHIUDI</button></section>`;
    modal.addEventListener('click', (event) => { if (event.target === modal) closeModal(); });
    modal.querySelector('[data-close]').addEventListener('click', closeModal);
    modal.querySelector('[data-edit]')?.addEventListener('click', () => openEditor(contact, rerender));
    modal.querySelector('[data-delete]')?.addEventListener('click', () => {
      if (!confirm(`Eliminare ${contact.name} dalla rubrica?`)) return;
      writeManual(readManual().filter((row) => row.id !== contact.id)); closeModal(); rerender();
    });
    document.body.appendChild(modal);
  }

  function openEditor(contact, rerender) {
    closeModal();
    const modal = document.createElement('div');
    modal.className = 'rubrica-v2-modal';
    modal.innerHTML = `<form class="rubrica-v2-dialog"><h3>${contact?.id ? 'Modifica contatto' : 'Nuovo contatto'}</h3><input name="name" required maxlength="120" placeholder="Nome e cognome" value="${esc(contact?.name)}"><input name="phone" inputmode="tel" maxlength="40" placeholder="Telefono (facoltativo)" value="${esc(contact?.phone)}"><input name="email" type="email" maxlength="160" placeholder="E-mail (facoltativa)" value="${esc(contact?.email)}"><input name="role" maxlength="100" placeholder="Ruolo (facoltativo)" value="${esc(contact?.role)}"><button class="rubrica-v2-action rubrica-v2-primary" type="submit">SALVA</button><button class="rubrica-v2-action" type="button" data-close>ANNULLA</button></form>`;
    modal.querySelector('[data-close]').addEventListener('click', closeModal);
    modal.addEventListener('click', (event) => { if (event.target === modal) closeModal(); });
    modal.querySelector('form').addEventListener('submit', (event) => {
      event.preventDefault(); const data = new FormData(event.currentTarget);
      const row = { id: contact?.id || `manual:${Date.now()}`, name:text(data.get('name')), phone:text(data.get('phone')), email:normalizeEmail(data.get('email')), role:text(data.get('role')), source:'manual', editable:true };
      if (!normalizePhone(row.phone) && !row.email) return alert('Inserisci almeno un numero di telefono oppure un’e-mail.');
      const rows = readManual(); const index = rows.findIndex((item) => item.id === row.id); if (index >= 0) rows[index] = row; else rows.push(row);
      writeManual(rows); closeModal(); rerender();
    });
    document.body.appendChild(modal); modal.querySelector('input')?.focus();
  }

  function valueFromRow(row, aliases) {
    const normalized = Object.fromEntries(Object.entries(row || {}).map(([key,value]) => [normalizeHeader(key), value]));
    return aliases.map((alias) => text(normalized[normalizeHeader(alias)])).find(Boolean) || '';
  }
  function mapImportedRow(row, index) {
    const phone = valueFromRow(row, ['TELEFONO','CELLULARE','MOBILE','NUMERO_TELEFONO','TELEFONO_PERSONALE']);
    const email = normalizeEmail(valueFromRow(row, ['EMAIL','E_MAIL','MAIL','EMAIL_ACCESSO_APP','LINKED_USER_EMAIL']));
    if (!normalizePhone(phone) && !email) return null;
    const fullName = valueFromRow(row, ['NOME_COMPLETO','NOMINATIVO','OPERATORE','DIPENDENTE']);
    const name = fullName || [valueFromRow(row,['NOME','FIRST_NAME']), valueFromRow(row,['COGNOME','LAST_NAME'])].filter(Boolean).join(' ') || email || phone || `Contatto ${index + 1}`;
    return { id:`import:${Date.now()}:${index}`, name, phone, email, role:valueFromRow(row,['MANSIONE','RUOLO','QUALIFICA']), company:valueFromRow(row,['AZIENDA','DITTA','SOCIETA']), source:'manual', imported:true, editable:true };
  }

  async function importMatrix(file, rerender) {
    if (!file) return;
    if (!window.XLSX) return alert('Lettore Excel non disponibile. Aggiorna l’app e riprova.');
    try {
      const workbook = window.XLSX.read(await file.arrayBuffer(), { type:'array', cellDates:false });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const sourceRows = window.XLSX.utils.sheet_to_json(sheet, { defval:'', raw:false });
      const imported = sourceRows.map(mapImportedRow).filter(Boolean);
      if (!imported.length) return alert('Nessun contatto importabile. Serve almeno TELEFONO/CELLULARE oppure EMAIL.');
      const existing = readManual(); const map = new Map(existing.map((row) => [contactKey(row), row]).filter(([key]) => key)); let added=0, updated=0;
      imported.forEach((row) => { const key=contactKey(row); const previous=map.get(key); if (previous) { map.set(key,{...previous,...row,id:previous.id,source:'manual',editable:true}); updated++; } else { map.set(key,row); added++; } });
      const ignored = Math.max(0, sourceRows.length - imported.length);
      if (!confirm(`Importare ${imported.length} contatti con telefono oppure e-mail?\n\nNuovi: ${added}\nAggiornati: ${updated}\nRighe senza telefono ed e-mail ignorate: ${ignored}`)) return;
      writeManual([...map.values()]); rerender(); alert(`Importazione completata.\nNuovi: ${added}\nAggiornati: ${updated}`);
    } catch (error) { console.error('Rubrica V2: importazione non riuscita.', error); alert(`Importazione non riuscita: ${error?.message || 'file non valido'}`); }
  }

  function openRubrica() {
    document.getElementById('rubrica-feature-page')?.remove(); document.getElementById('rubrica-feature-v2-page')?.remove(); installStyle();
    const page = document.createElement('section'); page.id='rubrica-feature-v2-page'; page.className='rubrica-v2-page';
    page.innerHTML = `<div class="rubrica-v2-shell"><header class="rubrica-v2-head"><button class="rubrica-v2-back">← INDIETRO</button><h2>📒 Rubrica</h2>${isManager() ? '<button class="rubrica-v2-import">📥 IMPORTA MATRICE PERSONALE</button><button class="rubrica-v2-add">+ CONTATTO</button><input class="rubrica-v2-file" type="file" accept=".xlsx,.xls,.csv" hidden>' : ''}</header><input class="rubrica-v2-search" type="search" placeholder="Cerca nome, telefono, e-mail, ruolo o azienda…"><div class="rubrica-v2-list"></div></div>`;
    const list=page.querySelector('.rubrica-v2-list'); const search=page.querySelector('.rubrica-v2-search');
    const render=()=>{ const query=text(search.value).toLocaleLowerCase('it'); const rows=buildContacts().filter((row)=>[row.name,row.phone,row.email,row.role,row.company].join(' ').toLocaleLowerCase('it').includes(query)); list.innerHTML=rows.length?rows.map((row,index)=>{const initials=text(row.name).split(/\s+/).slice(0,2).map((part)=>part[0]||'').join('').toUpperCase()||'👤';const details=[row.phone,row.email].filter(Boolean).join(' · ');return `<button class="rubrica-v2-card" data-index="${index}"><span class="rubrica-v2-avatar">${esc(initials)}</span><span class="rubrica-v2-main"><strong>${esc(row.name)}</strong><small>${esc(details)}</small>${row.role||row.company?`<small>${esc([row.role,row.company].filter(Boolean).join(' · '))}</small>`:''}</span><span>›</span></button>`;}).join(''):'<p class="rubrica-v2-empty">Nessun contatto con telefono oppure e-mail.</p>';list.querySelectorAll('[data-index]').forEach((button)=>button.addEventListener('click',()=>openContact(rows[Number(button.dataset.index)],render)));};
    search.addEventListener('input',render); page.querySelector('.rubrica-v2-back').addEventListener('click',()=>page.remove()); page.querySelector('.rubrica-v2-add')?.addEventListener('click',()=>openEditor(null,render));
    const file=page.querySelector('.rubrica-v2-file'); page.querySelector('.rubrica-v2-import')?.addEventListener('click',()=>file?.click()); file?.addEventListener('change',async()=>{const selected=file.files?.[0];file.value='';await importMatrix(selected,render);});
    document.body.appendChild(page); render();
  }

  function bindButton() {
    const current=document.getElementById('today-alerts-btn'); if(!current) return false;
    const replacement=current.cloneNode(true); replacement.dataset.rubricaV2Bound='1'; replacement.removeAttribute('disabled'); replacement.setAttribute('aria-label','Apri Rubrica contatti');
    replacement.querySelector('.today-summary-icon') && (replacement.querySelector('.today-summary-icon').textContent='📒'); const strong=replacement.querySelector('strong'); if(strong) strong.textContent='RUBRICA'; const spans=replacement.querySelectorAll('span'); if(spans.length) spans[spans.length-1].textContent='Contatti';
    replacement.addEventListener('click',openRubrica); current.replaceWith(replacement); try { if(window.ui?.todayAlertsBtn===current) window.ui.todayAlertsBtn=replacement; } catch(_){} return true;
  }

  function init(){ installStyle(); if(bindButton()) return; const observer=new MutationObserver(()=>{if(bindButton()) observer.disconnect();}); observer.observe(document.body,{childList:true,subtree:true}); setTimeout(()=>observer.disconnect(),30000); }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',init,{once:true}); else init();
})();