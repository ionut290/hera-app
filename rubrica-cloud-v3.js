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
    if (row.id && !String(row.id).startsWith('local:')) { await firestore.collection(COLLECTION).doc(row.id).set(clean, { merge:true }); return row.id; }
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
      const k = keyOf(row); if (!k || existing.has(k)) continue;
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
      writeLocal(CLOUD_CACHE_KEY, cloudRows); render(); migrateLocal().catch(console.warn);
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

  function initials(name) { return t(name).split(/\s+/).slice(0,2).map(x=>x[0]||'').join('').toUpperCase() || '—'; }
  function avatar(r, cls='rcv3-avatar') { return r.photoUrl ? `<img class="${cls}" src="${esc(r.photoUrl)}" alt="Foto ${esc(r.name)}">` : `<span class="${cls}" aria-hidden="true">${esc(initials(r.name))}</span>`; }
  function visibleEmail(value) { const e=email(value); return e.endsWith('@operatori.vargacantieri.app') ? '' : e; }
  function category(r) { if (r.source === 'operator') return 'operator'; if (/responsabile|caposquadra|amministratore|manager/i.test(`${r.role} ${r.company}`)) return 'manager'; if (r.company) return 'company'; return 'other'; }
  function whatsappNumber(value) { let n=phone(value).replace(/\D/g,''); if(n.startsWith('0'))n=`39${n.slice(1)}`; if(n&&!n.startsWith('39')&&n.length<=10)n=`39${n}`; return n; }

  function parseVcf(text) {
    return text.split(/END:VCARD/i).map(block=>{
      const fn=(block.match(/\nFN(?:;[^:]*)?:(.*)/i)||[])[1], tel=(block.match(/\nTEL(?:;[^:]*)?:(.*)/i)||[])[1], mail=(block.match(/\nEMAIL(?:;[^:]*)?:(.*)/i)||[])[1];
      return fn && (tel || mail) ? { name:t(fn), phone:t(tel), email:email(mail), source:'phone-vcf' } : null;
    }).filter(Boolean);
  }

  async function importRows(rows, render) {
    let ok=0, skip=0; const existing=new Set(allRows().map(keyOf).filter(Boolean));
    for (const row of rows) { const k=keyOf(row); if(!k||existing.has(k)){skip++;continue;} try{await saveCloud(row);existing.add(k);ok++;}catch(e){alert(e.message);break;} }
    render(); alert(`Importazione completata.\nAggiunti: ${ok}\nDuplicati o non validi: ${skip}`);
  }

  function style() {
    if (document.getElementById('rubrica-cloud-v3-style')) return;
    const s=document.createElement('style'); s.id='rubrica-cloud-v3-style'; s.textContent=`
    .rcv3{--navy:#123d36;--blue:#075fae;--green:#087a58;--line:#d7e2df;--soft:#eef5f3;position:fixed;inset:0;z-index:16000;background:#f5f8f7;color:var(--navy);overflow:auto;padding:calc(10px + env(safe-area-inset-top)) 14px calc(28px + env(safe-area-inset-bottom));font-family:inherit}.rcv3 *{box-sizing:border-box}.rcv3-shell{max-width:760px;margin:auto}.rcv3-head{position:sticky;top:calc(-10px - env(safe-area-inset-top));z-index:5;display:grid;grid-template-columns:48px 1fr 48px;align-items:center;padding:12px 0;background:#f5f8f7eF;backdrop-filter:blur(12px)}.rcv3-head h2{margin:0;text-align:center;font-size:21px;color:var(--navy)}.rcv3-head p{grid-column:2;margin:2px 0 0;text-align:center;color:#6d7d79;font-size:12px}.rcv3 button{font:inherit}.rcv3-icon-btn{width:44px;height:44px;border:1px solid var(--line);border-radius:14px;background:#fff;color:var(--navy);font-size:23px;display:grid;place-items:center}.rcv3-manager{display:grid;grid-template-columns:1fr auto;gap:10px;margin:10px 0 14px}.rcv3-add{min-height:52px;border:0;border-radius:15px;background:var(--green);color:#fff;font-weight:850;font-size:16px;padding:0 18px;box-shadow:0 7px 18px #087a5824}.rcv3-import-toggle{min-width:52px;border:1px solid var(--line);border-radius:15px;background:#fff;color:var(--blue);font-weight:850;padding:0 14px}.rcv3-actions{grid-column:1/-1;display:none;grid-template-columns:repeat(3,1fr);gap:8px;padding:10px;border:1px solid var(--line);border-radius:16px;background:#fff}.rcv3-actions.is-open{display:grid}.rcv3-actions button{min-height:44px;border:0;border-radius:11px;background:var(--soft);color:var(--blue);padding:7px;font-weight:800;font-size:12px}.rcv3-tools{position:sticky;top:72px;z-index:4;background:#f5f8f7;padding:2px 0 10px}.rcv3-search-wrap{position:relative}.rcv3-search-icon{position:absolute;left:15px;top:50%;transform:translateY(-50%);color:#788a85;font-size:18px;pointer-events:none}.rcv3-search{width:100%;min-height:50px;border:1px solid var(--line);border-radius:15px;padding:0 42px;font-size:16px;background:#fff;color:var(--navy);outline:none}.rcv3-search:focus{border-color:#54a48d;box-shadow:0 0 0 3px #54a48d20}.rcv3-clear{position:absolute;right:7px;top:6px;width:38px;height:38px;border:0;background:transparent;color:#7b8c87;font-size:20px}.rcv3-filters{display:flex;gap:7px;overflow:auto;padding:9px 1px 0;scrollbar-width:none}.rcv3-filter{white-space:nowrap;min-height:34px;border:1px solid var(--line);border-radius:999px;background:#fff;color:#536864;padding:0 13px;font-weight:750}.rcv3-filter.is-active{background:var(--navy);border-color:var(--navy);color:#fff}.rcv3-count{margin:7px 2px 0;color:#778783;font-size:12px}.rcv3-list{display:grid;gap:8px}.rcv3-letter{position:sticky;top:166px;z-index:2;margin:8px 0 0;padding:3px 4px;color:var(--green);font-size:13px;font-weight:900;background:#f5f8f7}.rcv3-card{width:100%;min-height:76px;border:1px solid var(--line);border-radius:17px;background:#fff;padding:11px 12px;text-align:left;display:flex;align-items:center;gap:12px;color:var(--navy);box-shadow:0 2px 7px #123d3608}.rcv3-card:active{transform:scale(.992)}.rcv3-avatar{width:50px;height:50px;border-radius:15px;object-fit:cover;background:#dfeeea;color:var(--green);display:grid;place-items:center;flex:none;font-weight:900}.rcv3-main{min-width:0;flex:1}.rcv3-name{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:15px;letter-spacing:.01em}.rcv3-meta{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#687b76;font-size:13px;margin-top:3px;font-weight:550}.rcv3-role{display:inline-block;margin-top:5px;padding:2px 8px;border-radius:999px;background:var(--soft);color:#55706a;font-size:11px;font-weight:750}.rcv3-chevron{font-size:24px;color:#a7b5b1}.rcv3-modal{position:fixed;inset:0;z-index:16010;background:#102c2780;display:flex;align-items:flex-end;justify-content:center;padding:14px;backdrop-filter:blur(3px)}.rcv3-dialog{width:min(520px,100%);max-height:92vh;overflow:auto;background:#fff;border-radius:24px;padding:20px;display:grid;gap:12px;box-shadow:0 24px 70px #0b241e40}.rcv3-dialog h3{margin:0;color:var(--navy);font-size:22px}.rcv3-dialog .rcv3-subtitle{margin:-7px 0 4px;color:#6b7e79}.rcv3-dialog input{min-height:48px;border:1px solid var(--line);border-radius:13px;padding:0 13px;font-size:16px}.rcv3-photo{width:92px;height:92px;border-radius:26px;object-fit:cover;background:#dfeeea;color:var(--green);display:grid;place-items:center;margin:auto;font-size:27px;font-weight:900}.rcv3-detail-head{text-align:center}.rcv3-detail-head h3{margin-top:12px}.rcv3-detail-lines{display:grid;gap:8px;margin:2px 0}.rcv3-detail-line{padding:11px 12px;border-radius:12px;background:var(--soft);color:#35544d;overflow-wrap:anywhere}.rcv3-quick{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}.rcv3-quick a{text-decoration:none}.rcv3-quick button,.rcv3-dialog>.rcv3-wide,.rcv3-dialog form button{width:100%;min-height:46px;border:0;border-radius:13px;background:var(--soft);color:var(--navy);font-weight:800}.rcv3-quick button{display:grid;gap:3px;place-items:center;font-size:12px}.rcv3-quick button span{font-size:21px}.rcv3-primary{background:var(--green)!important;color:#fff!important}.rcv3-secondary{background:#fff!important;border:1px solid var(--line)!important}.rcv3-danger{background:#fff1f1!important;color:#a11919!important}.rcv3-empty{text-align:center;color:#687a75;padding:54px 10px}.rcv3-empty strong{display:block;color:var(--navy);margin-bottom:6px}@media(max-width:560px){.rcv3-actions{grid-template-columns:1fr}.rcv3-actions button{font-size:13px}.rcv3-head h2{font-size:19px}}
    `; document.head.appendChild(s);
  }

  function editor(row, render) {
    const m=document.createElement('div'); m.className='rcv3-modal';
    m.innerHTML=`<form class="rcv3-dialog"><h3>${row?.id?'Modifica contatto':'Nuovo contatto'}</h3><p class="rcv3-subtitle">Inserisci i dati utili per la squadra.</p><div data-photo>${avatar(row||{},'rcv3-photo')}</div><input type="file" name="photo" accept="image/*"><input name="name" required placeholder="Nome e cognome" value="${esc(row?.name)}"><input name="phone" inputmode="tel" placeholder="Telefono" value="${esc(row?.phone)}"><input name="email" type="email" placeholder="E-mail" value="${esc(row?.email)}"><input name="role" placeholder="Ruolo" value="${esc(row?.role)}"><input name="company" placeholder="Azienda" value="${esc(row?.company)}"><button class="rcv3-primary" type="submit">Salva contatto</button><button class="rcv3-secondary" type="button" data-close>Annulla</button></form>`;
    let photoUrl=t(row?.photoUrl); const f=m.querySelector('form');
    f.photo.addEventListener('change', async()=>{ try { photoUrl=await imageToDataUrl(f.photo.files[0]); m.querySelector('[data-photo]').innerHTML=`<img class="rcv3-photo" src="${photoUrl}" alt="Anteprima foto">`; } catch(e){ alert(e.message); } });
    f.addEventListener('submit',async e=>{ e.preventDefault(); const d=new FormData(f); const next={...row,name:t(d.get('name')),phone:t(d.get('phone')),email:email(d.get('email')),role:t(d.get('role')),company:t(d.get('company')),photoUrl,source:row?.source||'manual'}; if(!phone(next.phone)&&!next.email)return alert('Inserisci telefono oppure e-mail.'); try{await saveCloud(next);m.remove();render();}catch(err){alert(`Salvataggio non riuscito: ${err.message}`);} });
    m.querySelector('[data-close]').onclick=()=>m.remove(); m.onclick=e=>{if(e.target===m)m.remove();}; document.body.appendChild(m);
  }

  function details(row, render) {
    const mail=visibleEmail(row.email), wa=whatsappNumber(row.phone);
    const m=document.createElement('div'); m.className='rcv3-modal';
    m.innerHTML=`<div class="rcv3-dialog"><div class="rcv3-detail-head">${avatar(row,'rcv3-photo')}<h3>${esc(row.name)}</h3><p class="rcv3-subtitle">${esc([row.role,row.company].filter(Boolean).join(' · ') || 'Contatto della rubrica')}</p></div><div class="rcv3-detail-lines">${row.phone?`<div class="rcv3-detail-line">${esc(row.phone)}</div>`:''}${mail?`<div class="rcv3-detail-line">${esc(mail)}</div>`:''}</div><div class="rcv3-quick">${row.phone?`<a href="tel:${esc(phone(row.phone))}"><button><span>☎</span>Chiama</button></a>`:''}${wa?`<a href="https://wa.me/${esc(wa)}" target="_blank" rel="noopener"><button><span>◉</span>WhatsApp</button></a>`:''}${mail?`<a href="mailto:${esc(mail)}"><button><span>✉</span>E-mail</button></a>`:''}</div>${row.editable&&manager()?'<button class="rcv3-wide rcv3-secondary" data-edit>Modifica contatto</button><button class="rcv3-wide rcv3-danger" data-delete>Elimina contatto</button>':''}<button class="rcv3-wide rcv3-primary" data-close>Chiudi</button></div>`;
    m.querySelector('[data-close]').onclick=()=>m.remove(); m.onclick=e=>{if(e.target===m)m.remove();};
    m.querySelector('[data-edit]')?.addEventListener('click',()=>{m.remove();editor(row,render);});
    m.querySelector('[data-delete]')?.addEventListener('click',async()=>{if(!confirm(`Eliminare definitivamente ${row.name}?`))return;try{await deleteCloud(row);m.remove();render();}catch(e){alert(e.message);}});
    document.body.appendChild(m);
  }

  function open() {
    document.querySelector('.rcv3')?.remove(); style();
    const page=document.createElement('section'); page.className='rcv3';
    page.innerHTML=`<div class="rcv3-shell"><header class="rcv3-head"><button class="rcv3-icon-btn" data-back aria-label="Indietro">‹</button><h2>Rubrica aziendale</h2><span></span><p>Contatti condivisi della squadra</p></header>${manager()?`<div class="rcv3-manager"><button class="rcv3-add" data-add>＋ Nuovo contatto</button><button class="rcv3-import-toggle" data-import-toggle aria-expanded="false">Importa</button><div class="rcv3-actions"><button data-phone>Dal telefono</button><button data-vcf>File VCF</button><input data-vcf-file type="file" accept=".vcf,text/vcard" hidden></div></div>`:''}<div class="rcv3-tools"><div class="rcv3-search-wrap"><span class="rcv3-search-icon">⌕</span><input class="rcv3-search" type="search" placeholder="Cerca un contatto…" aria-label="Cerca contatti"><button class="rcv3-clear" data-clear aria-label="Cancella ricerca" hidden>×</button></div><div class="rcv3-filters"><button class="rcv3-filter is-active" data-filter="all">Tutti</button><button class="rcv3-filter" data-filter="operator">Operatori</button><button class="rcv3-filter" data-filter="manager">Responsabili</button><button class="rcv3-filter" data-filter="company">Aziende</button></div><div class="rcv3-count"></div></div><div class="rcv3-list"></div></div>`;
    const list=page.querySelector('.rcv3-list'), search=page.querySelector('.rcv3-search'), count=page.querySelector('.rcv3-count'), clear=page.querySelector('[data-clear]'); let activeFilter='all';
    const render=()=>{
      const q=t(search.value).toLowerCase(); let rows=allRows().filter(r=>[r.name,r.phone,r.email,r.role,r.company].some(v=>t(v).toLowerCase().includes(q)));
      if(activeFilter!=='all')rows=rows.filter(r=>category(r)===activeFilter);
      count.textContent=`${rows.length} ${rows.length===1?'contatto':'contatti'}`; clear.hidden=!q;
      let letter=''; list.innerHTML=rows.length?rows.map((r,i)=>{const next=(t(r.name)[0]||'#').toUpperCase(), head=next!==letter?`<div class="rcv3-letter">${esc(next)}</div>`:'';letter=next;const mail=visibleEmail(r.email), meta=[r.phone,mail].filter(Boolean).join(' · '), role=[r.role,r.company].filter(Boolean).join(' · ');return `${head}<button class="rcv3-card" data-i="${i}">${avatar(r)}<span class="rcv3-main"><strong class="rcv3-name">${esc(r.name)}</strong>${meta?`<small class="rcv3-meta">${esc(meta)}</small>`:''}${role?`<small class="rcv3-role">${esc(role)}</small>`:''}</span><span class="rcv3-chevron">›</span></button>`;}).join(''):'<div class="rcv3-empty"><strong>Nessun contatto trovato</strong>Prova a modificare la ricerca o il filtro.</div>';
      list.querySelectorAll('[data-i]').forEach(b=>b.onclick=()=>details(rows[Number(b.dataset.i)],render));
    };
    search.oninput=render; clear.onclick=()=>{search.value='';search.focus();render();};
    page.querySelectorAll('[data-filter]').forEach(b=>b.onclick=()=>{activeFilter=b.dataset.filter;page.querySelectorAll('[data-filter]').forEach(x=>x.classList.toggle('is-active',x===b));render();});
    page.querySelector('[data-back]').onclick=()=>page.remove(); page.querySelector('[data-add]')?.addEventListener('click',()=>editor(null,render));
    page.querySelector('[data-import-toggle]')?.addEventListener('click',e=>{const box=page.querySelector('.rcv3-actions'),open=box.classList.toggle('is-open');e.currentTarget.setAttribute('aria-expanded',String(open));e.currentTarget.textContent=open?'Chiudi':'Importa';});
    page.querySelector('[data-phone]')?.addEventListener('click',async()=>{ if(navigator.contacts?.select){ try{const rows=await navigator.contacts.select(['name','tel','email','icon'],{multiple:true}); const mapped=await Promise.all(rows.map(async r=>({name:t(r.name?.[0]),phone:t(r.tel?.[0]),email:email(r.email?.[0]),photoUrl:r.icon?.[0]?await imageToDataUrl(r.icon[0]):'',source:'phone'}))); await importRows(mapped,render);}catch(e){if(e.name!=='AbortError')alert(`Importazione non disponibile: ${e.message}`);} } else alert('Il selettore diretto non è disponibile su questo dispositivo. Usa File VCF.'); });
    const file=page.querySelector('[data-vcf-file]'); page.querySelector('[data-vcf]')?.addEventListener('click',()=>file.click()); file?.addEventListener('change',async()=>{const f=file.files[0];if(f)await importRows(parseVcf(await f.text()),render);file.value='';});
    document.body.appendChild(page); startCloud(render); render();
  }

  window.openRubricaCloudV3 = open;
  document.addEventListener('click', e => { const el=e.target.closest('button,a,[role="button"]'); if(!el||el.closest('.rcv3'))return; const label=t(el.textContent).toUpperCase(); if(label==='RUBRICA'||label.includes('RUBRICA CONTATTI')){e.preventDefault();e.stopImmediatePropagation();open();} }, true);
})();
