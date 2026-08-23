(() => {
  "use strict";
  if (window.HeraOccasionalMultiSiteHours?.installed) return;

  const COMMESSA_ID = "lavori-occasionali";
  const state = { observer: null, mode: false };

  const txt = (id) => String(document.getElementById(id)?.textContent || document.getElementById(id)?.value || "").trim();
  const normalize = (v) => String(v || "").trim().replace(/\s+/g, " ").toLocaleUpperCase("it-IT");
  const hash = (value) => { let h = 2166136261; for (const c of String(value || "")) { h ^= c.charCodeAt(0); h = Math.imul(h, 16777619); } return (h >>> 0).toString(36); };
  const parseCoords = (value) => {
    const m = String(value || "").replace(/;/g, ",").match(/-?\d+(?:[.,]\d+)?/g) || [];
    if (m.length < 2) return null;
    const lat = Number(m[0].replace(",", ".")); const lng = Number(m[1].replace(",", "."));
    return Number.isFinite(lat) && Number.isFinite(lng) ? { lat, lng, text: `${lat.toFixed(6)}, ${lng.toFixed(6)}` } : null;
  };
  const collectionName = () => typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse";
  const currentDateKey = () => document.getElementById("squadra-riferimento")?.value || new Date().toLocaleDateString("en-CA", { timeZone: "Europe/Rome" });
  const currentRows = () => typeof readSquadraRows === "function" ? (readSquadraRows() || []).map((r) => ({ ...r })) : [];
  const currentUserData = () => ({ uid: String(typeof currentUser !== "undefined" ? currentUser?.uid || "" : ""), name: typeof getOperatorDisplayName === "function" ? getOperatorDisplayName() : String(typeof currentUser !== "undefined" ? currentUser?.displayName || currentUser?.email || "" : "") });
  const serverNow = () => typeof firebase !== "undefined" && firebase.firestore?.FieldValue?.serverTimestamp ? firebase.firestore.FieldValue.serverTimestamp() : new Date();

  function metadata() {
    const coordinates = parseCoords(txt("lavoro-occasionale-coordinate"));
    const nome = normalize(txt("lavoro-occasionale-nome"));
    return { nome, descrizione: txt("lavoro-occasionale-descrizione"), comune: txt("lavoro-occasionale-comune"), indirizzo: txt("lavoro-occasionale-indirizzo"), codicePrezzo: txt("lavoro-occasionale-codice-prezzo"), numeroPreventivo: txt("lavoro-occasionale-numero-preventivo"), coordinates };
  }

  function plantId(meta) {
    const slug = meta.nome.toLocaleLowerCase("it-IT").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 48) || "lavoro";
    return `occasionale-${slug}-${hash(`${meta.nome}|${meta.coordinates?.text || ""}`)}`;
  }

  function clearWorkFields() {
    ["lavoro-occasionale-nome","lavoro-occasionale-descrizione","lavoro-occasionale-coordinate","lavoro-occasionale-comune","lavoro-occasionale-indirizzo","lavoro-occasionale-codice-prezzo","lavoro-occasionale-numero-preventivo"].forEach((id) => { const el = document.getElementById(id); if (el) el.textContent = ""; });
    const pdf = document.getElementById("lavoro-occasionale-preventivo"); if (pdf) pdf.value = "";
    const status = document.getElementById("lavoro-occasionale-preventivo-status"); if (status) status.textContent = "Nessun PDF selezionato.";
    document.getElementById("lavoro-occasionale-nome")?.focus();
  }

  async function saveAdditionalSite() {
    const meta = metadata(); const rows = currentRows();
    if (!meta.nome || !meta.coordinates) return alert("Seleziona il cantiere sulla mappa e verifica nome e coordinate.");
    if (!rows.length) return alert("La squadra non è compilata.");
    if (typeof db === "undefined" || !db) return alert("Database non disponibile.");
    const id = plantId(meta); const dateKey = currentDateKey(); const user = currentUserData();
    const root = db.collection(collectionName()).doc(COMMESSA_ID);
    const plantRef = root.collection("impianti").doc(id);
    const payload = {
      id, commessaId: COMMESSA_ID, denominazione: meta.nome, nome: meta.nome, comune: meta.comune,
      indirizzo: meta.indirizzo, descrizioneVia: meta.indirizzo, latitudine: meta.coordinates.lat, longitudine: meta.coordinates.lng,
      gpsY: meta.coordinates.lat, gpsX: meta.coordinates.lng, coordinate: meta.coordinates.text, tipologiaImpianto: "LAVORO OCCASIONALE",
      codicePrezzo: meta.codicePrezzo, codiceVocePrezzo: meta.codicePrezzo, numeroPreventivo: meta.numeroPreventivo,
      tipologiaIntervento: meta.descrizione, tipologiaLavorazione: meta.descrizione, lavorazioniRichieste: meta.descrizione,
      note: meta.descrizione, lavoroOccasionale: true, multiCantiere: true, updatedAt: serverNow(), updatedBy: user.uid, updatedByName: user.name
    };
    await plantRef.set(payload, { merge: true });
    await root.collection("assegnazioniOccasionali").doc(`${dateKey}_${id}`).set({
      dateKey, plantId: id, commessaId: COMMESSA_ID, cantiere: meta.nome, comune: meta.comune, indirizzo: meta.indirizzo,
      coordinates: meta.coordinates.text, squadra: rows, createdAt: serverNow(), createdBy: user.uid, createdByName: user.name
    }, { merge: true });
    state.mode = false;
    alert(`Cantiere aggiunto alla stessa squadra: ${meta.nome}`);
    clearWorkFields();
  }

  function installFormControls() {
    const field = document.getElementById("lavoro-occasionale-field");
    if (!field || field.querySelector(".occasional-multi-controls")) return;
    const box = document.createElement("div"); box.className = "occasional-multi-controls";
    box.innerHTML = '<button type="button" class="btn occasional-new-site">➕ NUOVO CANTIERE STESSA SQUADRA</button><button type="button" class="btn btn-primary occasional-save-extra hidden">💾 SALVA CANTIERE AGGIUNTIVO</button><small class="muted">La composizione squadra resta invariata; ogni cantiere avrà ore proprie.</small>';
    field.appendChild(box);
    box.querySelector(".occasional-new-site").addEventListener("click", () => { state.mode = true; clearWorkFields(); box.querySelector(".occasional-save-extra").classList.remove("hidden"); });
    box.querySelector(".occasional-save-extra").addEventListener("click", async () => { try { await saveAdditionalSite(); box.querySelector(".occasional-save-extra").classList.add("hidden"); } catch (e) { console.error(e); alert("Non è stato possibile salvare il cantiere aggiuntivo."); } });
  }

  function hoursModal(plant) {
    const old = document.getElementById("occasional-hours-modal"); old?.remove();
    const modal = document.createElement("section"); modal.id = "occasional-hours-modal"; modal.className = "occasional-hours-modal";
    modal.innerHTML = `<div class="occasional-hours-card"><button type="button" class="occasional-hours-close">✕</button><h3>+ ORE · ${String(plant.denominazione || plant.nome || "Cantiere")}</h3><label>Data<input type="date" data-date value="${currentDateKey()}"></label><div class="occasional-hours-grid"><label>Ora inizio<input type="time" data-start></label><label>Ora fine<input type="time" data-end></label></div><label>Pausa (minuti)<input type="number" min="0" step="5" value="0" data-pause></label><label>Ore totali<input type="number" min="0" step="0.25" inputmode="decimal" data-hours placeholder="Es. 2,5"></label><label>Nota<input type="text" data-note placeholder="Facoltativa"></label><button type="button" class="btn btn-primary" data-save>💾 SALVA ORE CANTIERE</button></div>`;
    document.body.appendChild(modal); document.body.style.overflow = "hidden";
    const close = () => { modal.remove(); document.body.style.overflow = ""; };
    modal.querySelector(".occasional-hours-close").onclick = close;
    modal.querySelector("[data-save]").onclick = async () => {
      let ore = Number(String(modal.querySelector("[data-hours]").value || "").replace(",", "."));
      const start = modal.querySelector("[data-start]").value, end = modal.querySelector("[data-end]").value, pause = Number(modal.querySelector("[data-pause]").value || 0);
      if ((!Number.isFinite(ore) || ore <= 0) && start && end) {
        const [sh,sm] = start.split(":").map(Number), [eh,em] = end.split(":").map(Number); ore = Math.max(0, ((eh*60+em)-(sh*60+sm)-pause)/60);
      }
      if (!Number.isFinite(ore) || ore <= 0) return alert("Inserisci ore valide oppure ora inizio e fine.");
      const user = currentUserData(), dateKey = modal.querySelector("[data-date]").value || currentDateKey();
      const ref = db.collection(collectionName()).doc(COMMESSA_ID).collection("impianti").doc(String(plant.id || plant.docId)).collection("oreCantiere").doc();
      await ref.set({ dateKey, ore, oraInizio: start || "", oraFine: end || "", pausaMinuti: pause, nota: modal.querySelector("[data-note]").value.trim(), operatoreUid: user.uid, operatore: user.name, createdAt: serverNow() });
      close(); refreshCardHours(plant);
    };
  }

  async function refreshCardHours(plant, card) {
    try {
      const target = card || document.querySelector(`[data-occasional-hours-plant="${CSS.escape(String(plant.id || plant.docId || ""))}"]`)?.closest("article,div");
      const snap = await db.collection(collectionName()).doc(COMMESSA_ID).collection("impianti").doc(String(plant.id || plant.docId)).collection("oreCantiere").get();
      let total = 0; snap.forEach((d) => { total += Number(d.data()?.ore || 0) || 0; });
      const label = target?.querySelector(".occasional-hours-total"); if (label) label.textContent = `${total.toLocaleString("it-IT", { maximumFractionDigits: 2 })} ore cantiere`;
    } catch (_) {}
  }

  function plants() {
    const all = []; try { if (Array.isArray(currentImpianti)) all.push(...currentImpianti); } catch (_) {}
    try { const x = impiantiByCommessaId instanceof Map ? impiantiByCommessaId.get(COMMESSA_ID) : null; if (Array.isArray(x)) all.push(...x); } catch (_) {}
    return [...new Map(all.filter((p) => p?.lavoroOccasionale === true).map((p) => [String(p.id || p.docId || p.denominazione), p])).values()];
  }

  function decorateCards() {
    plants().forEach((plant) => {
      const id = String(plant.id || plant.docId || ""), name = normalize(plant.denominazione || plant.nome);
      document.querySelectorAll("button").forEach((b) => {
        if (normalize(b.textContent) !== "FATTO") return;
        let card = b.parentElement; while (card && card !== document.body && !normalize(card.textContent).includes(name)) card = card.parentElement;
        if (!card || card === document.body || card.querySelector(`[data-occasional-hours-plant="${CSS.escape(id)}"]`)) return;
        const wrap = document.createElement("div"); wrap.className = "occasional-hours-actions";
        wrap.innerHTML = `<button type="button" class="btn btn-primary" data-occasional-hours-plant="${id}">+ ORE CANTIERE</button><small class="occasional-hours-total muted">Ore cantiere…</small>`;
        b.parentElement?.insertAdjacentElement("afterend", wrap);
        wrap.querySelector("button").onclick = () => hoursModal(plant); refreshCardHours(plant, card);
      });
    });
  }

  function style() {
    if (document.getElementById("occasional-multi-hours-style")) return; const s = document.createElement("style"); s.id = "occasional-multi-hours-style";
    s.textContent = '.occasional-multi-controls{display:grid;gap:8px;margin-top:14px}.occasional-hours-actions{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:8px}.occasional-hours-modal{position:fixed;inset:0;z-index:2147483001;background:rgba(15,23,42,.55);display:grid;place-items:center;padding:16px}.occasional-hours-card{width:min(520px,100%);max-height:90vh;overflow:auto;background:#fff;border-radius:20px;padding:20px;display:grid;gap:12px;position:relative}.occasional-hours-card label{display:grid;gap:5px;font-weight:700}.occasional-hours-card input{min-height:44px;border:1px solid #cbd5e1;border-radius:12px;padding:8px 10px;font:inherit}.occasional-hours-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}.occasional-hours-close{position:absolute;right:12px;top:12px;border:0;background:transparent;font-size:22px}.hidden{display:none!important}'; document.head.appendChild(s);
  }

  function refresh() { style(); installFormControls(); decorateCards(); }
  state.observer = new MutationObserver(() => queueMicrotask(refresh)); state.observer.observe(document.documentElement, { childList: true, subtree: true });
  document.addEventListener("change", (e) => { if (e.target?.id === "squadra-commessa") refresh(); });
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", refresh, { once: true }); else refresh();
  window.HeraOccasionalMultiSiteHours = { installed: true, version: "1.0.0", refresh };
})();
