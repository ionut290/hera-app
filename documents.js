(() => {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const page = $("private-docs-page");
  const list = $("private-docs-list");
  const dialog = $("document-dialog");
  const form = $("document-form");
  if (!page || !list || !dialog || !form || !window.firebase?.firestore) return;

  const db = firebase.firestore();
  const auth = firebase.auth();
  const serverTime = () => firebase.firestore.FieldValue.serverTimestamp();
  const arrayUnion = (...values) => firebase.firestore.FieldValue.arrayUnion(...values);
  const state = { user: null, tab: "personal", filter: "all", query: "", documents: [], legacy: [], commesse: [], users: [], unsubs: [], commessaOnly: "" };
  const filters = [["all","Tutti"],["pdf","PDF"],["photo","Foto"],["word","Word"],["excel","Excel"],["expiration","Con scadenza"],["expired","Scaduti"],["soon","In scadenza"],["favorite","Preferiti"],["commessa","Collegati a commessa"]];
  const esc = (value) => String(value ?? "").replace(/[&<>'"]/g, (c) => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));
  const date = (value) => { const d = value?.toDate?.() || (value ? new Date(`${value}T12:00:00`) : null); return d && !Number.isNaN(d.valueOf()) ? d.toLocaleDateString("it-IT") : ""; };
  const uid = () => state.user?.uid || "";
  const visibility = (doc) => doc.visibility || "personal";
  const canEdit = (doc) => Boolean(uid() && (doc.ownerUserId === uid() || doc.createdBy === uid()));
  const authorized = (doc) => visibility(doc) === "personal" ? doc.ownerUserId === uid() : visibility(doc) === "global" || doc.sharedToAll || (doc.authorizedUserIds || doc.sharedUserIds || []).includes(uid());
  const commessaName = (id) => state.commesse.find((x) => x.id === id)?.nome || state.commesse.find((x) => x.id === id)?.name || "";
  const eventRef = (docId) => db.collection("calendarEvents").doc(`document_${docId}`);

  function fileKind(doc) {
    const value = `${doc.fileType || ""} ${doc.fileName || ""}`.toLowerCase();
    if (/pdf/.test(value)) return ["pdf","📕"];
    if (/image|jpg|jpeg|png|webp|heic/.test(value)) return ["photo","🖼️"];
    if (/word|docx?/.test(value)) return ["word","📘"];
    if (/excel|sheet|xlsx?/.test(value)) return ["excel","📗"];
    return ["other","📄"];
  }

  function expirationState(doc) {
    if (!doc.expirationDate) return ["",""];
    const today = new Date(); today.setHours(0,0,0,0);
    const expiry = new Date(`${doc.expirationDate}T00:00:00`);
    const days = Math.ceil((expiry - today) / 86400000);
    if (days < 0) return ["expired","Scaduto"];
    if (days <= 7) return ["soon", days === 0 ? "Scade oggi" : `Scade tra ${days} giorni`];
    return ["valid","In corso"];
  }

  function renderFilters() {
    $("documents-filters").innerHTML = filters.map(([id,label]) => `<button type="button" class="documents-filter ${state.filter === id ? "active" : ""}" data-doc-filter="${id}">${label}</button>`).join("");
  }

  function normalizeLegacy(item) {
    return { ...item, legacy: true, visibility: "personal", ownerUserId: uid(), category: item.category || (/pin/i.test(item.name || "") ? "PIN carburante" : /tessera/i.test(item.name || "") ? "Tessera" : "Documento"), description: item.description || item.note || "", fileUrl: item.fileUrl || item.driveWebViewLink || item.fileDataUrl || "", favoriteBy: item.favoriteBy || [] };
  }

  function visibleDocuments() {
    const all = [...state.documents, ...state.legacy.map(normalizeLegacy)].filter(authorized);
    const query = state.query.trim().toLocaleLowerCase("it");
    return all.filter((doc) => visibility(doc) === state.tab && (!state.commessaOnly || doc.commessaId === state.commessaOnly)).filter((doc) => {
      const [kind] = fileKind(doc); const [expiry] = expirationState(doc);
      if (state.filter === "pdf" || state.filter === "photo" || state.filter === "word" || state.filter === "excel") if (kind !== state.filter) return false;
      if (state.filter === "expiration" && !doc.expirationDate) return false;
      if ((state.filter === "expired" || state.filter === "soon") && expiry !== state.filter) return false;
      if (state.filter === "favorite" && !(doc.favoriteBy || []).includes(uid())) return false;
      if (state.filter === "commessa" && !doc.commessaId) return false;
      return !query || [doc.name,doc.category,doc.description,doc.code,doc.fileName,commessaName(doc.commessaId)].some((x) => String(x || "").toLocaleLowerCase("it").includes(query));
    }).sort((a,b) => Number((b.favoriteBy || []).includes(uid())) - Number((a.favoriteBy || []).includes(uid())) || String(b.createdAt?.seconds || "").localeCompare(String(a.createdAt?.seconds || "")));
  }

  function render() {
    renderFilters();
    document.querySelectorAll("[data-doc-tab]").forEach((button) => { button.classList.toggle("active", button.dataset.docTab === state.tab); button.setAttribute("aria-selected", String(button.dataset.docTab === state.tab)); });
    const rows = visibleDocuments();
    $("private-docs-feedback").textContent = state.commessaOnly ? `${rows.length} documenti autorizzati collegati a ${commessaName(state.commessaOnly) || "questa commessa"}.` : `${rows.length} ${rows.length === 1 ? "documento" : "documenti"}`;
    showDueReminder();
    if (!rows.length) { list.innerHTML = `<div class="documents-empty"><span>📂</span><strong>${state.commessaOnly ? "Nessun documento collegato a questa commessa." : "Nessun documento in questa scheda."}</strong>${state.user ? '<button class="btn btn-primary" data-empty-new>+ Aggiungi documento</button>' : ""}</div>`; list.querySelector("[data-empty-new]")?.addEventListener("click", () => openForm()); return; }
    list.innerHTML = rows.map((doc) => { const [kind,icon] = fileKind(doc); const [status,statusText] = expirationState(doc); const favorite = (doc.favoriteBy || []).includes(uid()); return `<article class="document-card" data-document-id="${esc(doc.id)}" data-legacy="${doc.legacy ? "1" : "0"}"><div class="document-format document-format-${kind}" aria-hidden="true">${icon}</div><div class="document-card-body"><div class="document-title-row"><h3>${esc(doc.name || "Documento")}</h3><button class="document-star ${favorite ? "active" : ""}" data-doc-action="favorite" aria-label="${favorite ? "Rimuovi dai" : "Aggiungi ai"} preferiti">★</button><button class="document-menu-btn" data-doc-action="menu" aria-label="Azioni documento">⋮</button></div><p>${esc(doc.category || "Senza categoria")}${doc.commessaId ? ` <span>• ${esc(commessaName(doc.commessaId) || "Commessa")}</span>` : ""}</p><div class="document-meta"><span>Caricato ${esc(date(doc.createdAt) || "-")}</span>${doc.expirationDate ? `<span class="expiration-${status}">Scadenza ${esc(date(doc.expirationDate))} • ${statusText}</span>` : ""}</div></div><div class="document-actions hidden"><button data-doc-action="open">Apri</button><button data-doc-action="download">Scarica</button>${canEdit(doc) && !doc.legacy ? '<button data-doc-action="edit">Modifica</button><button data-doc-action="replace">Sostituisci file</button><button data-doc-action="unlink">Rimuovi da commessa</button><button class="danger" data-doc-action="delete">Elimina</button>' : ""}</div></article>`; }).join("");
  }

  async function showDueReminder() {
    if (!state.user || document.querySelector(".document-reminder")) return;
    const today = new Date().toISOString().slice(0,10);
    const doc = state.documents.find((item) => authorized(item) && item.reminderEnabled && item.reminderDate && item.reminderDate <= today && item.expirationDate >= today && sessionStorage.getItem(`document-remind-later:${item.id}`) !== "hidden");
    if (!doc) return;
    const ackId = `${doc.id}_${uid()}`;
    const acknowledged = await db.collection("documentReminderAcknowledgements").doc(ackId).get().then((snap) => snap.exists).catch(() => false);
    if (acknowledged || document.querySelector(".document-reminder")) return;
    const reminder = document.createElement("aside"); reminder.className = "document-reminder"; reminder.setAttribute("role","alertdialog");
    reminder.innerHTML = `<strong>Documento in scadenza</strong><p>Il documento ${esc(doc.name)} scadrà il ${esc(date(doc.expirationDate))}.${doc.commessaId ? `<br>Commessa: ${esc(commessaName(doc.commessaId))}.` : ""}</p><div><button class="btn btn-primary" data-reminder-open>Apri documento</button><button class="btn" data-reminder-calendar>Apri calendario</button><button class="btn" data-reminder-ack>Ho capito</button><button class="btn" data-reminder-later>Ricordamelo</button></div>`;
    document.body.appendChild(reminder);
    reminder.querySelector("[data-reminder-open]").onclick = () => { state.tab=visibility(doc); state.query=doc.name; $("documents-search").value=doc.name; $("open-private-docs-btn")?.click(); render(); reminder.remove(); };
    reminder.querySelector("[data-reminder-calendar]").onclick = () => { location.hash=`calendar=${doc.expirationDate}`; reminder.remove(); };
    reminder.querySelector("[data-reminder-ack]").onclick = async () => { await db.collection("documentReminderAcknowledgements").doc(ackId).set({documentId:doc.id,userId:uid(),acknowledgedAt:serverTime()}); reminder.remove(); };
    reminder.querySelector("[data-reminder-later]").onclick = () => { sessionStorage.setItem(`document-remind-later:${doc.id}`,"hidden"); reminder.remove(); };
  }

  async function loadReferenceData() {
    const [commesse, users] = await Promise.all([db.collection("commesse").get().catch(() => null), db.collection("utenti").get().catch(() => null)]);
    state.commesse = commesse?.docs.map((d) => ({id:d.id,...d.data()})) || [];
    state.users = users?.docs.map((d) => ({id:d.id,...d.data()})).filter((u) => u.id !== uid()) || [];
    $("document-commessa").innerHTML = '<option value="">Nessuna commessa</option>' + state.commesse.map((c) => `<option value="${esc(c.id)}">${esc(c.nome || c.name || "Commessa")}</option>`).join("");
    $("document-users").innerHTML = state.users.map((u) => `<label><input type="checkbox" value="${esc(u.id)}"> ${esc(u.nome || u.displayName || u.email || "Utente")}</label>`).join("") || '<span class="muted">Nessun utente disponibile.</span>';
  }

  function stop() { state.unsubs.splice(0).forEach((fn) => fn()); state.documents = []; state.legacy = []; }
  function subscribe(user) {
    stop(); state.user = user; if (!user) { render(); return; }
    // Query separate: Firestore rules can validate each privacy boundary without exposing personal metadata.
    const merge = new Map();
    const watch = (query, key) => state.unsubs.push(query.onSnapshot({includeMetadataChanges:true}, (snap) => { snap.docChanges().forEach((change) => change.type === "removed" ? merge.delete(change.doc.id) : merge.set(change.doc.id,{id:change.doc.id,...change.doc.data()})); state.documents = [...merge.values()]; render(); }, () => render()));
    watch(db.collection("documents").where("ownerUserId","==",user.uid), "owner");
    watch(db.collection("documents").where("visibility","==","shared").where("sharedUserIds","array-contains",user.uid), "shared");
    watch(db.collection("documents").where("visibility","==","shared").where("sharedToAll","==",true), "all");
    watch(db.collection("documents").where("visibility","==","global"), "global");
    state.unsubs.push(db.collection("privateDocuments").doc(user.uid).collection("items").onSnapshot((snap) => { state.legacy = snap.docs.map((d) => ({id:d.id,...d.data()})); render(); }));
    loadReferenceData().then(render);
  }

  function showDialog() { if (typeof dialog.showModal === "function") dialog.showModal(); else dialog.setAttribute("open",""); setTimeout(() => $("document-name").focus(), 30); }
  function closeDialog() { if (dialog.open && dialog.close) dialog.close(); else dialog.removeAttribute("open"); }
  function openForm(doc = null, replace = false) {
    if (!state.user) return alert("Devi fare login per usare i documenti.");
    form.reset(); $("document-id").value = doc?.id || ""; $("document-form-title").textContent = replace ? "Sostituisci file" : doc ? "Modifica documento" : "Nuovo documento";
    $("document-name").value = doc?.name || ""; $("document-category").value = doc?.category || ""; $("document-description").value = doc?.description || ""; $("document-expiration").value = doc?.expirationDate || ""; $("document-commessa").value = doc?.commessaId || state.commessaOnly || ""; $("document-favorite").checked = (doc?.favoriteBy || []).includes(uid());
    const vis = doc?.visibility || state.tab || "personal"; form.querySelector(`[name=document-visibility][value=${vis}]`).checked = true;
    $("document-recipients").classList.toggle("hidden", vis !== "shared"); $("document-share-all").checked = Boolean(doc?.sharedToAll); form.querySelectorAll("#document-users input").forEach((input) => input.checked = (doc?.sharedUserIds || []).includes(input.value));
    $("document-file").required = Boolean(replace); $("document-form-feedback").textContent = ""; showDialog();
  }

  async function upload(file) {
    if (!file) return {};
    if (navigator.onLine === false) throw new Error("Connessione assente. Il documento sarà sincronizzato appena torni online.");
    if (typeof window.uploadBlobToDrive === "function") { const result = await window.uploadBlobToDrive(file, file.name, file.type || "application/octet-stream", "", {driveType:"DOCUMENTI",commessaName:"Documenti"}); return { fileUrl: result.webViewLink || "", driveFileId: result.fileId || "" }; }
    if (file.size > 700000) throw new Error("File troppo grande per il salvataggio locale: collega Google Drive e riprova.");
    const fileUrl = await new Promise((resolve,reject) => { const reader = new FileReader(); reader.onload=()=>resolve(String(reader.result)); reader.onerror=reject; reader.readAsDataURL(file); });
    return { fileUrl };
  }

  async function syncExpiration(docId, data, previous = {}) {
    const ref = eventRef(docId);
    if (!data.expirationDate) { if (previous.calendarEventId) await ref.delete().catch(() => {}); return ""; }
    const authorizedUserIds = data.visibility === "personal" ? [data.ownerUserId] : data.sharedToAll || data.visibility === "global" ? [] : data.sharedUserIds;
    await ref.set({ type:"SCADENZA_DOCUMENTO", title:`Scadenza documento – ${data.name}`, compactTitle:`📄 ${data.name}`, startDate:data.expirationDate, endDate:data.expirationDate, allDay:true, documentId:docId, documentVisibility:data.visibility, ownerUserId:data.ownerUserId, authorizedUserIds, sharedToAll:Boolean(data.sharedToAll || data.visibility === "global"), commessaId:data.commessaId || "", worksite:data.commessaId ? `Commessa: ${commessaName(data.commessaId)}` : "", description:`Scadenza documento${data.description ? `\n${data.description}` : ""}`, link:data.fileUrl || previous.fileUrl || "", createdByUid:data.createdBy || uid(), createdByName:state.user.displayName || state.user.email || "Utente", reminderDate:new Date(new Date(`${data.expirationDate}T12:00:00`).getTime()-7*86400000).toISOString().slice(0,10), reminderDaysBefore:7, updatedAt:serverTime() }, {merge:true});
    return ref.id;
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault(); const button = $("document-save"); button.disabled = true;
    try {
      const id = $("document-id").value || db.collection("documents").doc().id; const ref = db.collection("documents").doc(id); const old = state.documents.find((x) => x.id === id) || {}; const file = $("document-file").files?.[0];
      const selectedVisibility = form.querySelector("[name=document-visibility]:checked").value; const sharedUserIds = [...form.querySelectorAll("#document-users input:checked")].map((x) => x.value); const sharedToAll = $("document-share-all").checked;
      if (selectedVisibility === "shared" && !sharedToAll && !sharedUserIds.length) throw new Error("Seleziona almeno un destinatario per il documento condiviso.");
      const uploaded = await upload(file); const data = { name:$("document-name").value.trim(), category:$("document-category").value.trim(), description:$("document-description").value.trim(), visibility:selectedVisibility, ownerUserId:old.ownerUserId || uid(), sharedUserIds:selectedVisibility === "shared" ? sharedUserIds : [], sharedTeamIds:old.sharedTeamIds || [], sharedToAll:selectedVisibility !== "personal" && (selectedVisibility === "global" || sharedToAll), authorizedUserIds:selectedVisibility === "personal" ? [uid()] : sharedUserIds, commessaId:$("document-commessa").value, expirationDate:$("document-expiration").value, reminderDate:$("document-expiration").value ? new Date(new Date(`${$("document-expiration").value}T12:00:00`).getTime()-7*86400000).toISOString().slice(0,10) : "", reminderDaysBefore:7, reminderEnabled:Boolean($("document-expiration").value), favoriteBy:$("document-favorite").checked ? [...new Set([...(old.favoriteBy || []),uid()])] : (old.favoriteBy || []).filter((x) => x !== uid()), fileName:file?.name || old.fileName || "", fileType:file?.type || old.fileType || "", fileSize:file?.size || old.fileSize || 0, fileUrl:uploaded.fileUrl || old.fileUrl || "", driveFileId:uploaded.driveFileId || old.driveFileId || "", createdBy:old.createdBy || uid(), updatedBy:uid(), createdAt:old.createdAt || serverTime(), updatedAt:serverTime(), status:"active" };
      const calendarEventId = await syncExpiration(id,data,old); data.calendarEventId = calendarEventId; await ref.set(data,{merge:true}); closeDialog();
    } catch (error) { $("document-form-feedback").textContent = error.message || "Salvataggio non riuscito."; }
    finally { button.disabled = false; }
  });

  async function act(card, action) {
    const doc = [...state.documents,...state.legacy.map(normalizeLegacy)].find((x) => x.id === card.dataset.documentId); if (!doc) return;
    if (action === "menu") return card.querySelector(".document-actions").classList.toggle("hidden");
    if (action === "open" || action === "download") { if (navigator.onLine === false && !String(doc.fileUrl || "").startsWith("data:")) return alert("Connessione assente. Il file non può essere aperto."); if (!doc.fileUrl) return alert("Nessun file allegato."); const link=document.createElement("a"); link.href=doc.fileUrl; link.target="_blank"; link.rel="noopener"; if(action==="download") link.download=doc.fileName||doc.name; link.click(); return; }
    if (action === "edit") return openForm(doc);
    if (action === "replace") return openForm(doc,true);
    if (action === "favorite") { if (doc.legacy) return; const active=(doc.favoriteBy||[]).includes(uid()); await db.collection("documents").doc(doc.id).update({favoriteBy:active ? firebase.firestore.FieldValue.arrayRemove(uid()) : arrayUnion(uid()),updatedAt:serverTime()}); return; }
    if (action === "unlink") { await db.collection("documents").doc(doc.id).update({commessaId:"",updatedAt:serverTime()}); await syncExpiration(doc.id,{...doc,commessaId:""},doc); return; }
    if (action === "delete" && confirm(`Eliminare definitivamente “${doc.name}”?`)) { await eventRef(doc.id).delete().catch(()=>{}); await db.collection("documents").doc(doc.id).delete(); }
  }

  page.addEventListener("click", (event) => { const tab=event.target.closest("[data-doc-tab]"); if(tab){state.tab=tab.dataset.docTab;state.commessaOnly="";render();return;} const filter=event.target.closest("[data-doc-filter]");if(filter){state.filter=filter.dataset.docFilter;render();return;} const action=event.target.closest("[data-doc-action]");if(action)act(action.closest(".document-card"),action.dataset.docAction); });
  $("documents-search").addEventListener("input", (event) => { state.query=event.target.value; render(); });
  $("documents-new-btn").addEventListener("click", () => openForm()); $("document-dialog-close").addEventListener("click",closeDialog); $("document-cancel").addEventListener("click",closeDialog);
  form.addEventListener("change", (event) => { if(event.target.name === "document-visibility") $("document-recipients").classList.toggle("hidden",event.target.value !== "shared"); });
  $("open-private-docs-upload-btn")?.addEventListener("click", () => setTimeout(() => openForm(),80));
  $("commessa-documents-btn")?.addEventListener("click", () => { const match=location.hash.match(/commessa=([^&]+)/); state.commessaOnly=match?decodeURIComponent(match[1]):""; state.tab="global"; $("open-private-docs-btn")?.click(); setTimeout(render,60); });
  const online = () => $("documents-offline").classList.toggle("hidden",navigator.onLine); window.addEventListener("online",online);window.addEventListener("offline",online);online();renderFilters();auth.onAuthStateChanged(subscribe);
  window.HeraDocuments = { open: (options={}) => { state.commessaOnly=options.commessaId||""; state.tab=options.visibility||"global"; $("open-private-docs-btn")?.click(); setTimeout(()=>options.create?openForm():render(),50); } };
})();
