(() => {
  'use strict';
  if (window.__vargaRubricaCloudV3) return;
  window.__vargaRubricaCloudV3 = true;

  const LOCAL_KEY = 'varga_rubrica_contatti_manual_v1';
  const CLOUD_CACHE_KEY = 'varga_rubrica_cloud_cache_v1';
  const COLLECTION = 'rubricaContacts';
  let cloudRows = [];
  let unsubscribe = null;

  const t = v => String(v ?? '').trim();
  const esc = v => t(v).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const phone = v => t(v).replace(/[^\d+]/g, '').replace(/^00/, '+');
  const email = v => t(v).toLowerCase();
  const db = () => window.firebase?.firestore?.();
  const user = () => window.firebase?.auth?.()?.currentUser || null;
  const manager = () => { try { return typeof window.canManageData === 'function' && window.canManageData(); } catch { return false; } };
  const readLocal = key => { try { const x = JSON.parse(localStorage.getItem(key) || '[]'); return Array.isArray(x) ? x : []; } catch { return []; } };
  const writeLocal = (key, rows) => localStorage.setItem(key, JSON.stringify(rows));
  const keyOf = r => phone(r.phone) ? `p:${phone(r.phone)}` : email(r.email) ? `e:${email(r.email)}` : '';
  const uid = () => user()?.uid || 'anon';

  function operatorRows() {
    const people = Array.isArray(window.personaleRecords) ? window.personaleRecords : [];
    const users = Array.isArray(window.platformUsers) ? window.platformUsers : [];
    return people.map(p => {
      const e = email(p.email || p.emailAccessoApp || p.linkedUserEmail);
      const u = users.find(x => e && email(x.email) === e) || {};
      const ph = t(p.telefono || p.phone || p.cellulare || p.mobile || p.telefonoPersonale);
      const em = e || email(u.email);
      if (!phone(ph) && !em) return null;
      const name = t(p.displayName || p.nomeCompleto || `${t(p.nome)} ${t(p.cognome)}` || u.displayName || em || ph);
      return { id:`operator:${t(p.id || keyOf({phone:ph,email:em}))}`, name, phone:ph, email:em, role:t(p.mansione || p.ruolo || p.role), company:t(p.azienda || p.company), photoUrl:t(p.photoUrl || p.fotoUrl || u.photoURL), editable:false, source:'operator' };
    }).filter(Boolean);
  }

  function allRows() {
    const map = new Map();
    [...operatorRows(), ...readLocal(CLOUD_CACHE_KEY), ...cloudRows].forEach(r => {
      const k = keyOf(r) || `id:${r.id}`;
      const prev = map.get(k) || {};
      map.set(k, { ...prev, ...r, phone:r.phone || prev.phone, email:r.email || prev.email, photoUrl:r.photoUrl || prev.photoUrl });
    });
    return [...map.values()].sort((a,b)=>t(a.name).localeCompare(t(b.name),'it',{sensitivity:'base'}));
  }

  async function saveCloud(row) {
    const firestore = db();
    const clean = { name:t(row.name), phone:t(row.phone), email:email(row.email), role:t(row.role), company:t(row.company), photoUrl:t(row.photoUrl), source:t(row.source || 'manual'), updatedBy:uid(), updatedAt:window.firebase?.firestore?.FieldValue?.serverTimestamp?.() || new Date().toISOString() };
    if (!firestore || !user()) throw new Error('Accedi prima di salvare nella Rubrica condivisa.');
    if (row.id && !String(row.id).startsWith('local:')) {
      await firestore.collection(COLLECTION).doc(row.id).set(clean, { merge:true });
      return row.id;
    }
    const ref = await firestore.collection(COLLECTION).add({ ...clean, createdBy:uid(), createdAt:window.firebase?.firestore?.FieldValue?.serverTimestamp?.() || new Date().toISOString() });
    return ref.id;
  }

  async function deleteCloud(row) {
    if (db() && row.id && !String(row.id).startsWith('local:')) await db().collection(COLLECTION).doc(row.id).delete();
    cloudRows = cloudRows.filter(x => x.id !== row.id);
    writeLocal(CLOUD_CACHE_KEY, cloudRows);
  }

  async function migrateLocal() {
    const legacy = readLocal(LOCAL_KEY);
    if (!legacy.length || !db() || !user()) return;
    const existing = new Set(allRows().map(keyOf).filter(Boolean));
    for (const row of legacy) {
      const k = keyOf(row);
      if (!k || existing.has(k)) continue;
      try { await saveCloud({ ...row, id:null, source:'legacy' }); existing.add(k); } catch (e) { console.warn('Migrazione Rubrica non completata:', e); return; }
    }
    localStorage.setItem(`${LOCAL_KEY}_migrated`, new Date().toISOString());
  }

  function startCloud(render) {
    if (unsubscribe) return;
    const firestore = db();
    if (!firestore || !user()) { cloudRows = readLocal(CLOUD_CACHE_KEY); render(); return; }
    unsubscribe = firestore.collection(COLLECTION).onSnapshot(snap => {
      cloudRows = snap.docs.map(d => ({ id:d.id, ...d.data(), editable:true }));
      writeLocal(CLOUD_CACHE_KEY, cloudRows);
      render();
      migrateLocal().catch(console.warn);
    }, err => { console.warn('Rubrica cloud non disponibile:', err); cloudRows = readLocal(CLOUD_CACHE_KEY); render(); });
  }

  async function imageToDataUrl(file) {
    if (!file) return '';
    if (!file.type.startsWith('image/')) throw new Error('Seleziona un’immagine valida.');
    const img = await new Promise((resolve,reject)=>{ const i=new Image(); i.onload=()=>resolve(i); i.onerror=reject; i.src=URL.createObjectURL(file); });
    const max=320, scale=Math.min(1,max/Math.max(img.width,img.height));
    const c=document.createElement('canvas'); c.width=Math.max(1,Math.round(img.width*scale)); c.height=Math.max(1,Math.round(img.height*scale));
    c.getContext('2d').drawImage(img,0,0,c.width,c.height);
    return c.toDataURL('image/jpeg',0.72);
  }

  function initials(name) { return t(name).split(/\s+/).slice(0,2).map(x=>x[0]||'').join('').toUpperCase() || '👤'; }
  function avatar(r, cls='rcv3-avatar') { return r.photoUrl ? `<img class="${cls}" src="${esc(r.photoUrl)}" alt="Foto ${esc(r.name)}">` : `<span class="${cls}">${esc(initials(r.name))}</span>`; }

  function parseVcf(text) {
    return text.split(/END:VCARD/i).map((block,i)=>{
      const fn=(block.match(/\nFN(?:;[^:]*)?:(.*)/i)||[])[1];
      const tel=(block.match(/\nTEL(?:;[^:]*)?:(.*)/i)||[])[1];
      const mail=(block.match(/\nEMAIL(?:;[^:]*)?:(.*)/i)||[])[1];
      if (!fn || (!tel && !mail)) return null;
      return { name:t(fn), phone:t(tel), email:email(mail), source:'phone-vcf' };
    }).filter(Boolean);
  }

  async function importRows(rows, render) {
    let ok=0, skip=0;
    const existing=new Set(allRows().map(keyOf).filter(Boolean));
    for (const row of rows) {
      const k=keyOf(row); if (!k || existing.has(k)) { skip++; continue; }
      try { await saveCloud(row); existing.add(k); ok++; } catch (e) { alert(e.message); break; }
    }
    render(); alert(`Importazione completata.\nAggiunti: ${ok}\nDuplicati o non validi: ${skip}`);
  }

  function style() {
    if (document.getElementById('rubrica-cloud-v3-style')) return;
    const s=document.createElement('style'); s.id='rubrica-cloud-v3-style'; s.textContent=`
    .rcv3{position:fixed;inset:0;z-index:16000;background:#eef4f2;overflow:auto;padding:calc(12px + env(safe-area-inset-top)) 14px 30px}.rcv3-shell{max-width:760px;margin:auto}.rcv3-head{display:flex;gap:10px;align-items:center;flex-wrap:wrap}.rcv3 button{min-height:46px;border:1px solid #bfd0cb;border-radius:14px;background:#fff;padding:8px 12px;font-weight:850;color:#075fae}.rcv3 h2{flex:1;color:#12352f}.rcv3-actions{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:12px 0}.rcv3-search{width:100%;min-height:54px;border:1px solid #bfd0cb;border-radius:16px;padding:0 14px;font-size:17px}.rcv3-list{display:grid;gap:10px;margin-top:14px}.rcv3-card{display:flex;align-items:center;gap:12px;text-align:left;color:#173c35}.rcv3-avatar{width:52px;height:52px;border-radius:50%;object-fit:cover;background:#dcebe7;display:grid;place-items:center;flex:none;font-weight:900}.rcv3-main{min-width:0;flex:1}.rcv3-main strong,.rcv3-main small{display:block;overflow:hidden;text-overflow:ellipsis}.rcv3-modal{position:fixed;inset:0;z-index:16010;background:#0008;display:flex;align-items:flex-end;justify-content:center;padding:14px}.rcv3-dialog{width:min(520px,100%);background:#fff;border-radius:22px;padding:18px;display:grid;gap:10px}.rcv3-dialog input{min-height:46px;border:1px solid #ccd8d5;border-radius:12px;padding:0 12px;font-size:16px}.rcv3-photo{width:100px;height:100px;border-radius:50%;object-fit:cover;background:#dcebe7;display:grid;place-items:center;margin:auto;font-size:26px}.rcv3-primary{background:#0d7a58!important;color:#fff!important}.rcv3-danger{color:#a11919!important}.rcv3-empty{text-align:center;color:#687a75;padding:36px 10px}@media(max-width:520px){.rcv3-actions{grid-template-columns:1fr}}
    `; document.head.appendChild(s);
  }

  function editor(row, render) {
    const m=document.createElement('div'); m.className='rcv3-modal';
    m.innerHTML=`<form class="rcv3-dialog"><h3>${row?.id?'Modifica contatto':'Nuovo contatto'}</h3><div data-photo>${avatar(row||{},'rcv3-photo')}</div><input type="file" name="photo" accept="image/*"><input name="name" required placeholder="Nome e cognome" value="${esc(row?.name)}"><input name="phone" inputmode="tel" placeholder="Telefono" value="${esc(row?.phone)}"><input name="email" type="email" placeholder="E-mail" value="${esc(row?.email)}"><input name="role" placeholder="Ruolo" value="${esc(row?.role)}"><input name="company" placeholder="Azienda" value="${esc(row?.company)}"><button class="rcv3-primary" type="submit">SALVA</button><button type="button" data-close>ANNULLA</button></form>`;
    let photoUrl=t(row?.photoUrl); const f=m.querySelector('form');
    f.photo.addEventListener('change', async()=>{ try { photoUrl=await imageToDataUrl(f.photo.files[0]); m.querySelector('[data-photo]').innerHTML=`<img class="rcv3-photo" src="${photoUrl}">`; } catch(e){ alert(e.message); } });
    f.addEventListener('submit',async e=>{ e.preventDefault(); const d=new FormData(f); const next={...row,name:t(d.get('name')),phone:t(d.get('phone')),email:email(d.get('email')),role:t(d.get('role')),company:t(d.get('company')),photoUrl,source:row?.source||'manual'}; if(!phone(next.phone)&&!next.email)return alert('Inserisci telefono oppure e-mail.'); try{await saveCloud(next);m.remove();render();}catch(err){alert(`Salvataggio non riuscito: ${err.message}`);} });
    m.querySelector('[data-close]').onclick=()=>m.remove(); document.body.appendChild(m);
  }

  function details(row, render) {
    const m=document.createElement('div'); m.className='rcv3-modal';
    m.innerHTML=`<div class="rcv3-dialog">${avatar(row,'rcv3-photo')}<h3>${esc(row.name)}</h3><p>${esc(row.phone||'')}${row.phone&&row.email?' · ':''}${esc(row.email||'')}</p>${row.phone?`<a href="tel:${esc(phone(row.phone))}"><button>📞 CHIAMA</button></a>`:''}${row.email?`<a href="mailto:${esc(row.email)}"><button>✉️ E-MAIL</button></a>`:''}${row.editable&&manager()?'<button data-edit>✏️ MODIFICA</button><button class="rcv3-danger" data-delete>🗑️ ELIMINA</button>':''}<button data-close>CHIUDI</button></div>`;
    m.querySelector('[data-close]').onclick=()=>m.remove();
    m.querySelector('[data-edit]')?.addEventListener('click',()=>{m.remove();editor(row,render);});
    m.querySelector('[data-delete]')?.addEventListener('click',async()=>{if(!confirm(`Eliminare ${row.name}?`))return;try{await deleteCloud(row);m.remove();render();}catch(e){alert(e.message);}});
    document.body.appendChild(m);
  }

  function open() {
    document.querySelector('.rcv3')?.remove(); style();
    const page=document.createElement('section'); page.className='rcv3';
    page.innerHTML=`<div class="rcv3-shell"><div class="rcv3-head"><button data-back>← INDIETRO</button><h2>📒 Rubrica condivisa</h2></div>${manager()?`<div class="rcv3-actions"><button data-add>+ CONTATTO</button><button data-phone>📱 AGGIUNGI DAL TELEFONO</button><button data-vcf>📄 IMPORTA FILE VCF</button><input data-vcf-file type="file" accept=".vcf,text/vcard" hidden></div>`:''}<input class="rcv3-search" type="search" placeholder="Cerca nome, telefono, e-mail, ruolo o azienda…"><div class="rcv3-list"></div></div>`;
    const list=page.querySelector('.rcv3-list'), search=page.querySelector('.rcv3-search');
    const render=()=>{ const q=t(search.value).toLowerCase(); const rows=allRows().filter(r=>[r.name,r.phone,r.email,r.role,r.company].some(v=>t(v).toLowerCase().includes(q))); list.innerHTML=rows.length?rows.map((r,i)=>`<button class="rcv3-card" data-i="${i}">${avatar(r)}<span class="rcv3-main"><strong>${esc(r.name)}</strong><small>${esc([r.phone,r.email,r.role,r.company].filter(Boolean).join(' · '))}</small></span></button>`).join(''):'<div class="rcv3-empty">Nessun contatto con telefono oppure e-mail.</div>'; list.querySelectorAll('[data-i]').forEach(b=>b.onclick=()=>details(rows[Number(b.dataset.i)],render)); };
    search.oninput=render; page.querySelector('[data-back]').onclick=()=>{page.remove();}; page.querySelector('[data-add]')?.addEventListener('click',()=>editor(null,render));
    page.querySelector('[data-phone]')?.addEventListener('click',async()=>{ if(navigator.contacts?.select){ try{const rows=await navigator.contacts.select(['name','tel','email','icon'],{multiple:true}); const mapped=await Promise.all(rows.map(async r=>({name:t(r.name?.[0]),phone:t(r.tel?.[0]),email:email(r.email?.[0]),photoUrl:r.icon?.[0]?await imageToDataUrl(r.icon[0]):'',source:'phone'}))); await importRows(mapped,render);}catch(e){if(e.name!=='AbortError')alert(`Importazione non disponibile: ${e.message}`);} } else { alert('Il selettore diretto non è disponibile su questo dispositivo. Usa IMPORTA FILE VCF.'); } });
    const file=page.querySelector('[data-vcf-file]'); page.querySelector('[data-vcf]')?.addEventListener('click',()=>file.click()); file?.addEventListener('change',async()=>{const f=file.files[0];if(f)await importRows(parseVcf(await f.text()),render);file.value='';});
    document.body.appendChild(page); startCloud(render); render();
  }

  window.openRubricaCloudV3 = open;
  document.addEventListener('click', e => {
    const el=e.target.closest('button,a,[role="button"]'); if(!el || el.closest('.rcv3')) return;
    const label=t(el.textContent).toUpperCase();
    if(label==='RUBRICA' || label.includes('RUBRICA CONTATTI')) { e.preventDefault(); e.stopImmediatePropagation(); open(); }
  }, true);
})();