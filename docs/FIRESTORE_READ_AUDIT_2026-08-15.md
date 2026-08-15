# Audit letture Firestore — Hera App / Varga Cantieri

Data audit: 2026-08-15
Base analizzata: branch `main`
Obiettivo: ridurre letture Firestore senza compromettere caricamento commesse/squadre, FATTO, WhatsApp/WHAZZUP, ore, mappe e sincronizzazione multi-dispositivo.

## Metodo

Audit statico dell'intero repository orientato alle chiamate Firestore (`onSnapshot`, `.get()`, accessi `firebase.firestore`, query e moduli che intercettano le letture) e verifica dei meccanismi di ottimizzazione già presenti.

Questo documento distingue:
- letture realtime permanenti;
- letture one-shot;
- fallback che possono riaprire query native;
- letture amministrative/on-demand;
- moduli già protetti da cache, viste aggregate o guardie;
- punti da misurare con la diagnostica runtime prima di qualsiasi rimozione.

## Stato generale

L'app contiene già un sistema avanzato di riduzione letture. Non conviene applicare una nuova ottimizzazione globale indiscriminata: il rischio sarebbe duplicare wrapper Firestore già presenti e rompere la sincronizzazione. La strategia corretta è intervenire solo sui punti che la diagnostica runtime dimostra essere ancora costosi.

## Protezioni già presenti

### `firestore-operation-diagnostics.js`
- conta documenti letti e operazioni;
- distingue letture one-shot e consegne dei listener;
- registra collection, area, funzione chiamante, schermata, query e durata;
- registra listener aperti/chiusi e picco dei listener simultanei.

### `firestore-safe-optimizer.js`
- condivide listener fisici identici;
- evita aperture duplicate per query realmente equivalenti;
- protegge `commesse`, `squadreStorico`, `userAlerts`;
- usa una grace period prima di chiudere il listener fisico.

### `firestore-startup-cost-optimizer.js`
- impedisce l'avvio automatico di `chatMessages` quando non necessario;
- impedisce l'avvio automatico di `commessaResources` quando non necessario;
- blocca listener legacy `userAlerts` fuori dal centro notifiche;
- sostituisce, quando possibile, `squadreStorico` con viste statiche condivise;
- mantiene un fallback Firestore nativo se la vista condivisa non è disponibile.

### `activity-logs-read-disable.js`
- disattiva la lettura dell'archivio `activityLogs` nella schermata operativa;
- usa `appConfig/activeCommesse` per limitare le commesse caricate all'avvio;
- evita listener impianti per commesse disattivate;
- carica tutte le commesse solo quando viene aperta esplicitamente la gestione amministrativa.

### `shared-static-views.js` / `shared-static-views-client-core.js`
- spostano dati operativi ripetitivi verso documenti aggregati condivisi;
- possono sostituire molte letture con una singola lettura/listener.

### `firestore-registry-read-optimizer.js` / `registry-device-cache.js`
- riducono letture delle anagrafiche/registri con cache per dispositivo.

## Aree da controllare ancora

La ricerca statica mostra accessi Firestore o chiamate `.get()` anche nei seguenti moduli. La presenza di una chiamata non significa automaticamente che sia un problema: va distinta una lettura occasionale da una lettura ripetuta ad ogni apertura/render.

### Priorità A — avvio e uso quotidiano
- `app.js`
- `active-commesse-first-boot-guard.js`
- `squadra-current-save-sync.js`
- `shared-static-views.js`
- `shared-static-views-client-core.js`
- `firestore-startup-cost-optimizer.js`
- `firestore-safe-optimizer.js`
- `firestore-registry-read-optimizer.js`
- `approval-access.js`
- `auth-login-fix.js`
- `login-retry-fix.js`

Da verificare:
- nessuna doppia sottoscrizione alla stessa query;
- listener chiusi quando si cambia pagina/commessa;
- nessun `get()` ripetuto durante render multipli;
- nessuna lettura completa di collection quando basta un documento o una query filtrata;
- nessun fallback nativo `squadreStorico` attivato inutilmente dopo timeout;
- nessuna lettura profilo/accesso duplicata durante login + retry + approvazione.

### Priorità B — moduli operativi aperti su richiesta
- `global-archive-sync.js`
- `operational-import-repair.js`
- `inrete-work-items-v2.js`
- `hours-export-range.js`
- `app-worklimate.js`
- `app-atex.js`
- `private-documents-v2.js`
- `operator-profile-feature.js`
- `identity-card-feature.js`
- `google-sheet-two-way-sync.js`
- `registry-google-sheet-sync.js`

Regola: questi moduli non devono generare letture Firestore finché la rispettiva funzione/pagina non viene realmente aperta, salvo dati indispensabili alla Home.

### Priorità C — Preventivi / contabilità
- `preventivi-core.js`
- `preventivi-firestore-chunks.js`
- `preventivi-persistenza-selezioni-fix.js`
- `preventivi-commessa-search-bridge.js`
- `preventivi-storage-config.js`
- `accounting-v2.js`

Questi moduli possono avere dataset più grandi. Devono restare lazy/on-demand e preferire cache/chunk/indici mirati invece di letture complete ripetute.

### Backend / funzioni
- `functions/index.js`
- `functions/shared-operational-views.js`
- `functions/user-notifications.js`
- `functions/run-calendar-admin-reminders.js`
- script di rebuild/backfill

Le letture server non pesano sul caricamento del client ma contribuiscono ai costi Firestore. Vanno auditati separatamente per trigger duplicati, scansioni complete e rebuild troppo frequenti.

## Rischi individuati

### 1. Troppi wrapper sullo stesso `Query.prototype.onSnapshot`
Il progetto utilizza più moduli che intercettano `onSnapshot`. È utile ma delicato. Ogni nuovo wrapper deve preservare correttamente il precedente e l'ordine di caricamento.

Azione: non aggiungere nuovi wrapper generici se la stessa ottimizzazione può essere fatta nel modulo specifico.

### 2. Fallback delle viste condivise
`squadreStorico` può ricadere sul listener Firestore nativo dopo timeout. Se la vista condivisa è lenta/non inizializzata, si rischia di pagare sia la preparazione della vista sia il listener originale.

Azione: misurare `staticSquadreFallbacks`. Se > 0 nell'uso normale, priorità alta.

### 3. Listener duplicati ma non identici
Il multiplexer unisce solo query equivalenti. Due query quasi uguali con filtri/ordinamenti differenti continueranno a creare due listener fisici.

Azione: usare la diagnostica `queries` per individuare query diverse sulla stessa collection che possono essere consolidate.

### 4. Letture amministrative accidentali all'avvio
La gestione completa delle commesse è già on-demand. Lo stesso principio va verificato per utenti, documenti, Preventivi, contabilità, registri e moduli speciali.

### 5. `get()` ripetuti da retry/render
I moduli login/accesso e alcuni bridge possono effettuare letture one-shot. Una singola lettura è corretta; la stessa lettura ripetuta più volte in pochi secondi no.

Azione: cercare nella diagnostica funzioni con molte `readOperations` e pochi documenti per operazione.

## Soglie proposte per sessione standard
Sessione: apertura app -> Home -> apertura di una commessa -> un impianto, senza pannelli amministrativi.

- listener duplicati identici: 0;
- `activityLogs` letti: 0;
- `chatMessages` letti senza aprire chat: 0;
- `commessaResources` letti senza aprire risorse: 0;
- raccolte Preventivi/contabilità lette senza aprire i moduli: 0;
- documenti privati letti senza aprire Documenti: 0;
- fallback `squadreStorico`: 0 in condizioni normali;
- letture di commesse disattivate: 0 all'avvio;
- letture impianti di commesse non operative: da minimizzare.

## Ordine di intervento
1. Eseguire una diagnostica pulita della build corrente per 60-120 secondi.
2. Ordinare per `collections`, `queries`, `functions` e `listenerInstances`.
3. Correggere prima i listener permanenti con più documenti consegnati.
4. Correggere poi i `get()` ripetuti.
5. Spostare su viste aggregate solo dataset realmente ripetitivi e condivisi.
6. Verificare dopo ogni modifica che FATTO, commesse, squadre, ore e WhatsApp restino invariati.
7. Applicare una sola ottimizzazione per volta e confrontare due diagnostiche prima/dopo.

## Funzioni critiche da preservare
- caricamento affidabile di commesse e squadre;
- aggiornamenti multi-dispositivo necessari;
- persistenza e visualizzazione FATTO;
- ordinario/straordinario;
- foto e WHAZZUP/WhatsApp;
- gestione ore;
- navigazione e dati impianto;
- autenticazione e autorizzazioni;
- Android e PWA.

## Esito audit statico

Il codice contiene già molte ottimizzazioni valide e non richiede una riscrittura generale. Il margine ulteriore più probabile è in:
1. listener residui aperti all'avvio da moduli non visibili;
2. fallback di `squadreStorico`;
3. letture one-shot duplicate nei flussi login/accesso;
4. moduli Preventivi/contabilità/documenti caricati prima dell'apertura della pagina;
5. query simili ma non identiche che il multiplexer non può condividere;
6. funzioni backend con possibili scansioni complete.

La prossima ottimizzazione deve essere decisa dai numeri della diagnostica runtime della versione corrente.

## Checklist runtime
- avvio da app completamente chiusa;
- attendere caricamento Home;
- non aprire menu per 30 secondi;
- aprire una sola commessa;
- aprire un solo impianto;
- non premere FATTO durante il baseline;
- esportare la diagnostica;
- ripetere su Android e web/PWA.

Con questi due report sarà possibile quantificare esattamente le prossime riduzioni e fissare un obiettivo numerico per le letture iniziali.