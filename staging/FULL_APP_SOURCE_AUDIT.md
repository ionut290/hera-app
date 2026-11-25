# Audit sorgenti applicazione originale — staging

Data audit: 25 novembre 2026  
Branch: `varga-cantieri-v2-staging`  
Repository: `ionut290/hera-app`

## Risultato

La cartella principale del branch contiene già la shell originale `index.html` (151.035 caratteri) e la base di codice dell'app. La via corretta è mantenere questa applicazione come runtime operativo unico e usare il portale React esistente soltanto come prototipo/area amministrativa finché non viene realizzato un ingresso tenant sicuro. Non serve ricostruire le schermate.

- Script locali caricati direttamente da `index.html`: **72**
- Fogli CSS locali caricati direttamente: **24**
- Riferimenti JS aggiuntivi individuati in `firebase-config.js`: **24**
- Riferimenti locali non trovati nel tree GitHub: **0** (vedi elenco).

## Script originali caricati da index.html

- `loading-humor.js`
- `app-pure-utils.js`
- `app-worklimate.js`
- `app-atex.js`
- `app-documents.js`
- `app-calendar.js`
- `snow-castel-maggiore-catalog.js`
- `snow-castel-maggiore-kml-real.js`
- `app-snow.js`
- `app-availability.js`
- `heavy-libs-lazy-loader.js`
- `firebase-config.js`
- `android-prelogin-permission-gate.js`
- `native-android-runtime.js`
- `notification-center.js`
- `auth-login-fix.js`
- `varga-session-transfer.js`
- `approval-access.js`
- `first-login-password.js`
- `password-recovery-code.js`
- `login-retry-fix.js`
- `account-privacy-controls.js`
- `fuel-stations-core.js`
- `fuel-stations-national-cache.js`
- `fuel-stations-search.js`
- `app.js`
- `green-assistant.js`
- `android-play-store-install.js`
- `maintenance-price-code-display-guard.js`
- `commessa-stats-cache-optimizer.js`
- `shared-static-views-client.js`
- `squadra-current-save-sync.js`
- `lavori-occasionali.js`
- `documentazione-cantiere.js`
- `lavori-occasionali-pdf-storage.js`
- `activity-logs-read-disable.js`
- `management-core.js`
- `registry-google-sheet-sync.js`
- `management-v2.js`
- `documents.js`
- `coordinate-repair.js`
- `inrete-work-items-v2.js`
- `accounting-v2.js`
- `commessa-impianti-menu.js`
- `commessa-produced-widget.js`
- `accounting-view-guard.js`
- `operational-import-repair.js`
- `today-summary-interactions.js`
- `today-live-hours-vehicles.js`
- `identity-feature-lazy-loader.js`
- `fuel-stations-integration.js`
- `fatto-button-immediate.js`
- `administrative-calendar.js`
- `mezzi-alimentazione-fix.js`
- `hours-export-range.js`
- `header-menu-runtime.js`
- `timbratura-ios-notifications.js`
- `timbratura-ios-geofence.js`
- `timbratura-reminders.js`
- `cobo-mowing-work-orders.js`
- `tree-work-orders.js`
- `potature-followup.js`
- `tree-search.js`
- `green-areas.js`
- `urban-furniture.js`
- `wastewater-plants.js`
- `commessa-navigation-repair.js`
- `snow-plan-operations.js`
- `salatura-strada-routes.js`
- `salatura-strada-integration.js`
- `gestionale-desktop-return.js`
- `rdi-player.js`

## CSS originali caricati da index.html

- `style.css`
- `menu-grouping.css`
- `timbratura-reminders.css`
- `management-v2.css`
- `accounting-v2.css`
- `first-login-password.css`
- `login-retry-fix.css`
- `private-documents-v2.css`
- `squadre-restyle.css`
- `calendar-feature.css`
- `notification-center.css`
- `approval-access.css`
- `documentazione-cantiere.css`
- `green-assistant.css`
- `tree-search.css`
- `cobo-mowing-work-orders.css`
- `green-areas.css`
- `green-area-sheet.css`
- `urban-furniture.css`
- `wastewater-plants.css`
- `wastewater-infrastructure.css`
- `ios-safe-area.css`
- `snow-plan-operations.css`
- `salatura-strada-integration.css`

## Script caricati dinamicamente da firebase-config.js

- `storage-quota-guard.js`
- `app-notifications-read-guard.js`
- `firestore-operation-diagnostics.js`
- `firestore-diagnostics-dashboard-v4.js`
- `firestore-safe-optimizer.js`
- `native-android-runtime.js`
- `firestore-inflight-read-coalescer.js`
- `firestore-diagnostics-optimizer-extension.js`
- `shared-static-views.js`
- `active-commesse-first-boot-guard.js`
- `firestore-startup-cost-optimizer.js`
- `shared-static-views-ui.js`
- `admin-user-access-tools.js`
- `admin-user-access-share-fix.js`
- `lavori-occasionali-workflow-v2.js`
- `lavori-occasionali-squad-first.js`
- `lavori-occasionali-google-places.js`
- `lavori-occasionali-multi-cantiere-ore.js`
- `lavori-occasionali-pdf-storage.js`
- `squadre-lavori-occasionali-cantieri.js`
- `notification-session-enhancements.js`
- `update-app-feature.js`
- `google-sheet-two-way-sync.js`
- `personnel-training-manager.js`

## Riferimenti locali non trovati nel tree

- Nessuno

## Rischi e blocchi rilevati

1. `firebase-config.js` è già predisposto per un Firebase staging, ma richiede `window.__VARGA_STAGING_FIREBASE_CONFIG__` e usa un placeholder se manca la chiave API. Non va inserita alcuna chiave di produzione.
2. `netlify.toml` blocca deliberatamente la build: `STAGING deployment blocked pending isolation audit`. Non rimuovere il blocco finché l'isolamento non è verificato.
3. `staging/portal/src/App.jsx` è un portale semplificato e mostra ancora moduli operativi in sviluppo; non è la copia completa richiesta.
4. Le regole Firestore legacy root non dimostrano isolamento multi-tenant. Non collegare il runtime completo a Firestore finché tutti i percorsi, listener, cache, storage e integrazioni non sono isolati.
5. TinyFish ha letto il file originale `app.js` (1,38 MB) e ha verificato i punti di ingresso commesse: `getCommesseCollectionName()` (riga 537), `createCommessa()` (riga 636), `addManagementImpianto()` (riga 265), `firebase.initializeApp()` (riga 42) e i percorsi Firestore relativi. Questo helper centralizza i percorsi delle commesse nel solo `app.js`, ma NON centralizza tutte le collezioni dell'intera applicazione.
6. Il file `index.html` fa riferimento a icone sotto `icons/` che non risultano nel tree corrente: `icons/varga-cantieri-32.png`, `icons/varga-cantieri-180.png`, `icons/varga-cantieri-512.png`. Verificare se vengono generate da altra pipeline prima di una build.

## Prossimo gate tecnico

- Usare la root app originale, senza riscrivere UI o moduli.
- Aggiungere solo un gate di autenticazione/selezione tenant davanti alla root app.
- Introdurre un data adapter tenant-aware compatibile con il modello Firestore originale, poi migrare e testare un modulo alla volta.
- Inventariare ogni accesso Firebase e integrazione esterna prima di attivare il runtime.
- Preservare FATTO/WHAZZUP e i dati Personale/Utenti secondo `AGENTS.md`.
- Mantenere deploy bloccato, staging only, nessun dato Avola copiato e nessuna modifica alla produzione.
