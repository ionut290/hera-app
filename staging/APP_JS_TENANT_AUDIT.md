# Analisi TinyFish del codice originale app.js

Fonte letta direttamente: `app.js` nel branch `varga-cantieri-v2-staging` (1.386.659 byte).

## Punti verificati

- Firebase inizializzato alle righe 42–45.
- `getCommesseCollectionName()`, riga 537: restituisce `neve_commesse` nel contesto neve, altrimenti `commesse`.
- `isSnowServiceContext()`, riga 1069: determina il contesto neve.
- `createCommessa(event)`, riga 636: crea una commessa con `nome`, `codice`, `parentCommessaId`, `creatoDa`, `createdAt`, passa da `canManageData()`, poi può inizializzare un foglio Google tramite `driveAccessToken`.
- `addManagementImpianto()`, riga 265: usa la form originale, valida la denominazione, legge gli impianti esistenti per il numero progressivo, aggiunge il documento in `{commesseCollection}/{commessaId}/impianti`, aggiorna il parent con `nextImpiantoNumber`, poi aggiorna lista e mappa.
- `normalizeCommessaDocument()`, riga 409.
- `getMainCommesse()`, riga 1022; `getSubcommesse(parentCommessaId)`, riga 1026.
- L'autenticazione Firebase usa `auth.onAuthStateChanged(...)` e avvia `runAfterFirstRender()` / `runDeferredStartupTasks()`.

## Percorsi commesse originali

- `db.collection(getCommesseCollectionName())`
- `db.collection(getCommesseCollectionName()).doc(commessaId).collection("impianti")`
- `db.collection(getCommesseCollectionName()).doc(commessaId).collection("noteCommessa")`
- Altri percorsi distinti: `globalCommesse` e `globalImpianti`.

## Collezioni individuate in app.js

`commesse`, `neve_commesse`, `globalCommesse`, `globalImpianti`, `users`, `platformUsers`, `adminUsers`, `squadre`, `squadreStorico`, `personale`, `mezzi`, `resources`, `appConfig`, `appNotifications`, `userAlerts`, `chatMessages`, `gpsUpdateRequests`, `programmazioni`, `oreReports`, `oreApprovalRequests`, `commessaResources`, `ferie`, `safetyContacts`, `impiantoSafetyProcedures`, `weatherAlerts`, `weatherForecast`, `bannerConfig`, `workBanner`, `userActivityLogs`, `driveAdminSecret`, `driveBridge`, `posDocuments`, `notificationEmailQueue`, `privateDocs`, `calendarEvents`.

## Conseguenza tecnica

`getCommesseCollectionName()` è un punto di adattamento molto utile per commesse e subcommesse nel file `app.js`, ma non basta per clonare in sicurezza l'intera app. Le altre collezioni sopra elencate e gli altri moduli JS/CSS devono essere inventariati e instradati separatamente al tenant, con regole Firestore e storage appropriate. Non cambiare questo helper da solo e dichiarare completato il multi-tenant.

## Vincoli di compatibilità da preservare

- Non cambiare la UI o i campi delle form originali.
- Mantenere la gerarchia commessa principale/subcommessa.
- Mantenere il numero progressivo degli impianti e il campo `nextImpiantoNumber`.
- Mantenere aggiornamento della lista e della mappa.
- Mantenere la creazione opzionale del foglio Google solo con integrazione configurata per quel tenant, non ereditata da Avola.
- Non alterare FATTO/WHAZZUP, gestione impianti protetta, personale/utenti e relativi collegamenti.
- Non puntare la staging a produzione e non distribuire finché le regole tenant non sono verificate.
