# VARGA CANTIERI — prova Firestore reale ma solo in emulatore

**Non è un deploy Firebase e non attiva la cancellazione degli utenti.**

Il nuovo workflow `check-account-deletion-request-firestore-emulator.yml`
usa il vero sorgente delle due callable di richiesta/stato e il vero Firestore
Emulator con progetto fittizio `demo-account-deletion-request`. Solo
l'autenticazione e il wrapper HTTPS sono simulati.

Verifica transazioni, timestamp server, richieste duplicate, accesso per UID,
rifiuto di richieste non autenticate o con token vecchio, campi minimi,
passaggio simulato `in_review` e protezione di documenti fittizi non interessati.

Il test rifiuta host diversi da loopback, progetti diversi dal progetto demo e
credenziali di servizio. I dati sono interamente sintetici.

**Non sono stati collaudati Apple Sign in/revoca OAuth su dispositivo fisico,
né il processo di cancellazione effettiva, l'approvazione GDPR, l'account
reviewer o il rilascio App Store.**

Riferimento: https://firebase.google.com/docs/emulator-suite/connect_firestore
