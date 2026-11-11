# Build Apple selezionabile per App Review — senza pubblicazione automatica

La build interna TestFlight 4 è stata caricata con `testFlightInternalTestingOnly: true`, quindi **Apple la rende non selezionabile nella scheda iOS 1.0**. Non è un errore dell'interfaccia o dell'account.

## Correzione sicura
Aggiunto workflow **`Varga Cantieri iOS App Store Candidate (manual upload only)`**: riusa gli stessi controlli già riusciti, Firebase iOS e la firma Apple del workflow TestFlight, ma non aggiunge il flag Internal Only. Carica solo un candidato idoneo ad App Store Connect; non invia la versione alla revisione, né pubblica l'app.

- Richiede esecuzione manuale e selezione `UPLOAD_APP_STORE_CANDIDATE_ONLY`.
- Richiede SHA esatto del commit revisionato e nuovo numero build (ad es. 5).
- Usa l'ambiente Apple già configurato `ios-internal-testflight`, senza copiare segreti.
- Nessuna modifica a sorgenti app, FATTO, WhatsApp, Firestore, Android, privacy o utenti.
- Nessuna esecuzione automatica e nessun upload finché l'utente non avvia il workflow.

Dopo l'upload e l'elaborazione di Apple: App Store Connect → Distribuzione → iOS App Version 1.0 → Add Build → scegliere la **nuova** build, non la 4. Verificare privacy, account demo e informazioni per i revisori prima di Add for Review.