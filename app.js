firebase.initializeApp(firebaseConfig);

const auth = firebase.auth();
const db = firebase.firestore();

const ui = {
  loginBtn: document.getElementById("login-btn"),
  logoutBtn: document.getElementById("logout-btn"),
  user: document.getElementById("user"),
  form: document.getElementById("impianto-form"),
  denominazione: document.getElementById("denominazione"),
  comune: document.getElementById("comune"),
  indirizzo: document.getElementById("indirizzo"),
  excelFile: document.getElementById("excel-file"),
  importBtn: document.getElementById("import-btn"),
  importFeedback: document.getElementById("import-feedback"),
  lista: document.getElementById("impianti-lista")
};

let pendingRows = [];
let unsubscribeImpianti = null;

ui.loginBtn.addEventListener("click", loginWithGoogle);
ui.logoutBtn.addEventListener("click", logout);
ui.form.addEventListener("submit", onImpiantoSubmit);
ui.excelFile.addEventListener("change", onExcelSelected);
ui.importBtn.addEventListener("click", importPendingRows);

function loginWithGoogle() {
  const provider = new firebase.auth.GoogleAuthProvider();
  provider.setCustomParameters({ prompt: "select_account" });

  auth.signInWithPopup(provider).catch((error) => {
    if (error.code === "auth/popup-blocked" || error.code === "auth/cancelled-popup-request") {
      return auth.signInWithRedirect(provider);
    }

    console.error("Errore login:", error);
    alert("Errore login Google: " + error.message);
  });
}

function logout() {
  auth.signOut().catch((error) => {
    console.error("Errore logout:", error);
    alert("Errore logout: " + error.message);
  });
}

auth.onAuthStateChanged((user) => {
  if (user) {
    ui.user.textContent = `Loggato: ${user.displayName || "Utente"} (${user.email || "email non disponibile"})`;
    ui.loginBtn.disabled = true;
    ui.logoutBtn.disabled = false;
    subscribeImpianti();
  } else {
    ui.user.textContent = "Non loggato";
    ui.loginBtn.disabled = false;
    ui.logoutBtn.disabled = true;
    stopImpiantiSubscription();
    ui.lista.innerHTML = "<p class='muted'>Effettua il login per vedere gli impianti.</p>";
  }

  ui.importBtn.disabled = !user || pendingRows.length === 0;
});

async function onImpiantoSubmit(event) {
  event.preventDefault();

  const user = auth.currentUser;
  if (!user) {
    alert("Devi fare login prima di salvare un impianto.");
    return;
  }

  const payload = {
    denominazione: ui.denominazione.value.trim(),
    comune: ui.comune.value.trim(),
    indirizzo: ui.indirizzo.value.trim()
  };

  if (!payload.denominazione) {
    alert("La denominazione è obbligatoria.");
    return;
  }

  try {
    await saveImpianto(payload, user, false);
    ui.form.reset();
  } catch (error) {
    console.error("Errore salvataggio:", error);
    alert("Errore salvataggio impianto: " + error.message);
  }
}

function saveImpianto(payload, user, importedFromExcel) {
  return db.collection("impianti").add({
    denominazione: payload.denominazione || "",
    comune: payload.comune || "",
    indirizzo: payload.indirizzo || "",
    creatoDa: user.displayName || "",
    emailOperatore: user.email || "",
    importedFromExcel: Boolean(importedFromExcel),
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  });
}

function onExcelSelected(event) {
  const file = event.target.files && event.target.files[0];
  pendingRows = [];
  ui.importBtn.disabled = true;
  ui.importFeedback.classList.remove("error");

  if (!file) {
    ui.importFeedback.textContent = "Nessun file selezionato.";
    return;
  }

  const reader = new FileReader();

  reader.onload = (e) => {
    try {
      const workbook = XLSX.read(e.target.result, { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: "", raw: false });

      pendingRows = rows.map(normalizeRow).filter((row) => row.denominazione);

      if (!pendingRows.length) {
        ui.importFeedback.textContent = "Nessuna riga valida trovata. Servono colonne con denominazione/comune/indirizzo.";
        return;
      }

      ui.importFeedback.textContent = `Righe pronte da importare: ${pendingRows.length}`;
      ui.importBtn.disabled = !auth.currentUser;
    } catch (error) {
      console.error("Errore parsing Excel:", error);
      ui.importFeedback.textContent = "Errore lettura file Excel.";
      ui.importFeedback.classList.add("error");
    }
  };

  reader.readAsArrayBuffer(file);
}

async function importPendingRows() {
  const user = auth.currentUser;
  if (!user) {
    alert("Devi fare login prima di importare.");
    return;
  }

  if (!pendingRows.length) {
    alert("Nessuna riga da importare.");
    return;
  }

  ui.importBtn.disabled = true;

  try {
    await importRowsInChunks(pendingRows, user);
    pendingRows = [];
    ui.excelFile.value = "";
    ui.importFeedback.classList.remove("error");
    ui.importFeedback.textContent = "Import completato con successo.";
  } catch (error) {
    console.error("Errore importazione:", error);
    ui.importFeedback.classList.add("error");
    ui.importFeedback.textContent = "Errore durante l'importazione.";
    alert("Errore importazione Excel: " + error.message);
  } finally {
    ui.importBtn.disabled = !auth.currentUser || pendingRows.length === 0;
  }
}

async function importRowsInChunks(rows, user) {
  const MAX_BATCH = 450;

  for (let i = 0; i < rows.length; i += MAX_BATCH) {
    const chunk = rows.slice(i, i + MAX_BATCH);
    const batch = db.batch();

    chunk.forEach((row) => {
      const ref = db.collection("impianti").doc();
      batch.set(ref, {
        denominazione: row.denominazione,
        comune: row.comune,
        indirizzo: row.indirizzo,
        distretto: row.distretto,
        idSap: row.idSap,
        voceRiferimento: row.voceRiferimento,
        sfalciAreeVerdiMqPotaturaSiepiM: row.sfalciAreeVerdiMqPotaturaSiepiM,
        frequenzaAnnuaMinima: row.frequenzaAnnuaMinima,
        tipologiaIntervento: row.tipologiaIntervento,
        coordinateGpsY: row.coordinateGpsY,
        coordinateGpsX: row.coordinateGpsX,
        creatoDa: user.displayName || "",
        emailOperatore: user.email || "",
        importedFromExcel: true,
        createdAt: firebase.firestore.FieldValue.serverTimestamp()
      });
    });

    await batch.commit();
  }
}

function normalizeRow(row) {
  const normalized = normalizeKeys(row);

  return {
    denominazione: getFirstValue(normalized, [
      "denominazione",
      "denominazioneimpianto",
      "impianto",
      "nomeimpianto"
    ]),
    comune: getFirstValue(normalized, [
      "comune",
      "comuneubicazioneimpianto",
      "citta",
      "città"
    ]),
    indirizzo: getFirstValue(normalized, [
      "indirizzo",
      "via",
      "address",
      "viaecivicodiubicazioneimpianto"
    ]),
    distretto: getFirstValue(normalized, ["distretto"]),
    idSap: getFirstValue(normalized, ["idsap"]),
    voceRiferimento: getFirstValue(normalized, ["vocediriferimentoelencoprezzi"]),
    sfalciAreeVerdiMqPotaturaSiepiM: getFirstValue(normalized, [
      "sfalciareeverdimqpotaturasiepim"
    ]),
    frequenzaAnnuaMinima: getFirstValue(normalized, [
      "frequenzaannuaminimasfalcieopotaturasiepin"
    ]),
    tipologiaIntervento: getFirstValue(normalized, ["tipologiadisfalciointervento"]),
    coordinateGpsY: getFirstValue(normalized, ["coordinategpsy"]),
    coordinateGpsX: getFirstValue(normalized, ["coordinategpsx"])
  };
}

function normalizeKeys(obj) {
  return Object.entries(obj).reduce((acc, [key, value]) => {
    const compactKey = normalizeHeaderKey(key);
    acc[compactKey] = String(value || "").trim();
    return acc;
  }, {});
}

function normalizeHeaderKey(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function getFirstValue(obj, keys) {
  for (const key of keys) {
    if (obj[key]) return obj[key];
  }
  return "";
}

function subscribeImpianti() {
  if (unsubscribeImpianti) return;

  unsubscribeImpianti = db
    .collection("impianti")
    .orderBy("createdAt", "desc")
    .onSnapshot((snapshot) => {
      ui.lista.innerHTML = "";

      if (snapshot.empty) {
        ui.lista.innerHTML = "<p class='muted'>Nessun impianto presente.</p>";
        return;
      }

      snapshot.forEach((doc) => {
        const impianto = doc.data();
        const item = document.createElement("article");

        item.className = "impianto-item";
        item.innerHTML = `
          <strong>${escapeHTML(impianto.denominazione || "(senza denominazione)")}</strong>
          <p><b>Comune:</b> ${escapeHTML(impianto.comune || "-")}</p>
          <p><b>Indirizzo:</b> ${escapeHTML(impianto.indirizzo || "-")}</p>
          <p><b>Operatore:</b> ${escapeHTML(impianto.creatoDa || "-")} (${escapeHTML(impianto.emailOperatore || "-")})</p>
        `;

        ui.lista.appendChild(item);
      });
    }, (error) => {
      console.error("Errore lettura impianti:", error);
      ui.lista.innerHTML = "<p class='error'>Errore nel caricamento degli impianti.</p>";
    });
}

function stopImpiantiSubscription() {
  if (unsubscribeImpianti) {
    unsubscribeImpianti();
    unsubscribeImpianti = null;
  }
}

function escapeHTML(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
