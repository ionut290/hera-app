firebase.initializeApp(firebaseConfig);

const auth = firebase.auth();
const db = firebase.firestore();

function login() {
  const provider = new firebase.auth.GoogleAuthProvider();
  auth.signInWithPopup(provider)
    .then((result) => {
      console.log("Login riuscito", result.user.displayName);
    })
    .catch((error) => {
      console.error("Errore login:", error);
      alert("Errore login: " + error.message);
    });
}

function logout() {
  auth.signOut();
}

auth.onAuthStateChanged((user) => {
  const el = document.getElementById("user");
  if (!el) return;

  if (user) {
    el.innerHTML = "Loggato: " + user.displayName + " (" + user.email + ")";
  } else {
    el.innerHTML = "Non loggato";
  }
});

function salvaImpianto() {
  const user = auth.currentUser;
  if (!user) {
    alert("Devi fare login prima.");
    return;
  }

  const denominazione = document.getElementById("denominazione").value.trim();
  const comune = document.getElementById("comune").value.trim();
  const indirizzo = document.getElementById("indirizzo").value.trim();

  if (!denominazione) {
    alert("Inserisci la denominazione impianto.");
    return;
  }

  db.collection("impianti").add({
    denominazione: denominazione,
    comune: comune,
    indirizzo: indirizzo,
    creatoDa: user.displayName || "",
    emailOperatore: user.email || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  })
  .then(() => {
    document.getElementById("denominazione").value = "";
    document.getElementById("comune").value = "";
    document.getElementById("indirizzo").value = "";
    alert("Impianto salvato.");
  })
  .catch((error) => {
    console.error("Errore salvataggio:", error);
    alert("Errore salvataggio: " + error.message);
  });
}

db.collection("impianti")
  .orderBy("createdAt", "desc")
  .onSnapshot((snapshot) => {
    const lista = document.getElementById("impianti-lista");
    if (!lista) return;

    lista.innerHTML = "";

    snapshot.forEach((doc) => {
      const impianto = doc.data();

      const div = document.createElement("div");
      div.style.border = "1px solid #ccc";
      div.style.padding = "10px";
      div.style.marginBottom = "10px";

      div.innerHTML = `
        <strong>${impianto.denominazione || ""}</strong><br>
        Comune: ${impianto.comune || ""}<br>
        Indirizzo: ${impianto.indirizzo || ""}<br>
        Operatore: ${impianto.creatoDa || ""} (${impianto.emailOperatore || ""})
      `;

      lista.appendChild(div);
    });
  }, (error) => {
    console.error("Errore lettura impianti:", error);
  });