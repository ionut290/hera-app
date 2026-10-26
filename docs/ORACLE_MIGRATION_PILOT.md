# Migrazione progressiva Oracle — stato e runbook pilota

## Perimetro realizzato

Questo ramo aggiunge un servizio separato in `oracle-pilot-api/`. Non modifica il client pubblicato, il login corrente, Firebase, Netlify, FATTO/WHAZZUP, gli impianti o i dati operativi.

Il servizio è intenzionalmente limitato a letture autenticate di:

- catalogo `commesse` / `impianti` tramite `varga_catalogo_api_v1`;
- coordinate valide tramite `varga_mappa_api_v1`.

L'utente PostgreSQL dell'applicazione deve essere `varga_reader`: non deve avere `SELECT` su `firestore_import_test_20261026`, né privilegi di scrittura o DDL.

## Inventario delle dipendenze che impediscono un cutover immediato

L'app corrente dipende almeno da:

- Firebase Authentication (Google, email/password, recupero, sessione e accesso da Varga Gestionale);
- Firestore e relativi listener, regole, indici, viste condivise e coda offline;
- Cloud Functions Firebase (accessi, Drive, calendari, viste aggregate, notifiche ed errori);
- funzioni Netlify per meteo, carburanti, Google Sheet, assistente verde, arredo urbano e sollevamenti;
- Firebase Storage, Realtime Database, Cloud Messaging e wrapper Android/Capacitor;
- Google Drive/Sheets e credenziali amministrative conservate lato backend;
- funzioni protette e immutabili FATTO/WHAZZUP/RESET con sincronizzazione Firestore.

Le 27.518 righe importate costituiscono un archivio di compatibilità, non ancora un modello relazionale operativo completo. Le due viste esistenti coprono soltanto catalogo e mappa; non coprono scritture, transazioni, autorizzazioni per-record, sincronizzazione offline, Storage, notifiche o funzioni backend.

## Autenticazione e autorizzazione pilota

Per ridurre il rischio, il primo stadio riusa soltanto l'identità Firebase già emessa al client e verifica l'ID token lato API. La verifica non legge né modifica Firestore. L'accesso è **deny by default**:

- email presente in `PILOT_ALLOWED_EMAILS`; oppure
- claim firmato `varga_access=true`;
- ruolo amministrativo solo tramite `PILOT_ADMIN_EMAILS` o claim `varga_role=admin`.

Questo non replica ancora lo stato account e i permessi dinamici conservati in Firestore. Prima di ampliare il pilota è obbligatorio migrare utenti/ruoli in tabelle dedicate preservando UID ed email storici, oppure emettere claim affidabili da un processo amministrativo controllato. Non usare un token o un ruolo fornito dal browser senza verifica della firma.

La modalità iniziale verifica firma, progetto, scadenza ed email verificata. `CHECK_REVOKED_TOKENS=false` evita di distribuire una credenziale Firebase Admin sul server pilota; di conseguenza una sessione revocata resta valida fino alla scadenza naturale del relativo ID token. Prima di estendere il pilota occorre fornire una credenziale minima e impostare il controllo a `true`, oppure completare il passaggio a OIDC Oracle.

Il servizio supporta anche `AUTH_MODE=oidc` per una futura Identity Domain Oracle. Il passaggio da Firebase a OIDC è una fase successiva e richiede mappatura UID, test di revoca, recupero account, utenti disabilitati e rollback.

## Preparazione sul server Oracle (nessun cutover)

1. Copiare soltanto `oracle-pilot-api/` in una directory pilota separata.
2. Creare `.env` da `.env.example`, con password PostgreSQL in un secret manager o file `0600`; non salvarla nel repository.
3. Verificare che `db-postgres-1` e l'API condividano una rete Docker dedicata e che PostgreSQL non sia esposto pubblicamente.
4. Come proprietario delle viste, eseguire `sql/001_api_views.sql`; poi rieseguire i controlli SQL sotto riportati.
5. Costruire il container e avviarlo su `127.0.0.1:18082`. L'accesso esterno deve passare da reverse proxy HTTPS; non pubblicare direttamente la porta.
6. Conservare il web pilota su `127.0.0.1:18081` e inserire il suo origin esatto in `ALLOWED_ORIGINS`.
7. Inserire inizialmente una sola email di collaudo in `PILOT_ALLOWED_EMAILS`.
8. Verificare `healthz`, `readyz`, rifiuto senza token, rifiuto di un utente non ammesso e risposte con token ammesso.

Esempio di controlli non distruttivi:

```sql
SELECT has_table_privilege('varga_reader','public.firestore_import_test_20261026','SELECT') AS archivio_deve_essere_false;
SELECT has_table_privilege('varga_reader','public.varga_catalogo_api_v1','SELECT') AS catalogo_deve_essere_true;
SELECT has_table_privilege('varga_reader','public.varga_mappa_api_v1','SELECT') AS mappa_deve_essere_true;
SELECT collection_name, COUNT(*) FROM public.varga_catalogo_api_v1 GROUP BY collection_name ORDER BY collection_name;
SELECT collection_name, COUNT(*) FROM public.varga_mappa_api_v1 GROUP BY collection_name ORDER BY collection_name;
```

Valori di riferimento comunicati per le viste sorgenti:

- archivio importato: 27.518 documenti;
- vista catalogo grezza: 3.264 righe classificate come `commesse` e 67 impianti;
- mappa: 1.679 commesse e 67 impianti;
- archivio non leggibile da `varga_reader`.

L'analisi dei percorsi ha mostrato che le 3.264 righe `commesse` includono numerose sottocollezioni. Il catalogo API filtra perciò i soli documenti principali: 20 commesse (`commesse/<id>`) e 67 impianti (`impianti/<id>`). La mappa mantiene invece i 1.679 punti commessa validi perché rappresentano elementi geografici delle sottocollezioni. Ogni differenza futura va spiegata prima di procedere.

## Gate obbligatori prima di collegare una schermata

- TLS valido e reverse proxy con timeout e limite dimensione richieste;
- backup PostgreSQL verificato con prova di ripristino in ambiente isolato;
- scansione segreti e dipendenze senza vulnerabilità high/critical;
- test di autorizzazione con utente ammesso, non ammesso, disabilitato e token scaduto/revocato;
- confronto conteggi e campione di record per catalogo e mappa;
- log privi di token, password, payload Firestore e dati personali;
- metriche di latenza/errori e procedura di rollback;
- nessun cambio al DNS, a Netlify o a Firebase durante il pilota.

## Strategia di migrazione successiva

1. **Shadow read**: una pagina pilota separata confronta Oracle e Firebase senza influire sull'app reale.
2. **Read canary**: pochi utenti leggono catalogo/mappa da Oracle; fallback esplicito al comportamento corrente.
3. **Modello relazionale**: progettare tabelle e vincoli per utenti, personale, commesse, impianti e relazioni, mantenendo gli ID Firestore.
4. **Dual-write controllato**: solo dopo test di idempotenza, audit e riconciliazione; mai sul flusso FATTO/WHAZZUP sigillato senza una soluzione esterna compatibile con il blocco del repository.
5. **Migrazione dei servizi**: sostituire separatamente Functions, Storage/Drive, notifiche, offline e Netlify.
6. **Cutover**: richiede approvazione esplicita, finestra, backup ripristinabile, prove E2E web/Android e rollback provato.

## Componenti non migrati

Non sono migrati: login del client, profili e permessi dinamici, scritture CRUD, FATTO/WHAZZUP/RESET, offline sync, Personale, Utenti, squadre, calendario, ore, mezzi, notifiche, file/Storage, Drive/Sheets, Cloud Functions, funzioni Netlify, Android e DNS/TLS pubblico. Pertanto la migrazione completa e il cutover **non sono completati**.

## Adattatore per l'interfaccia originale

`oracle-read-adapter.js` prepara il contratto dati per la futura copia canary dell'interfaccia originale. È disattivato per impostazione predefinita, non viene incluso dall'HTML di produzione e non avvia richieste all'installazione. Quando esplicitamente configurato nella sola copia pilota, riusa il token dell'utente già autenticato e normalizza catalogo e coordinate in oggetti compatibili, marcati `oraclePilotReadOnly`.

L'adattatore non legge o scrive Firestore, non crea listener e non modifica HTML/CSS. Il collegamento a `commesseById` e alle mappe dell'app originale resta intenzionalmente non attivo finché non saranno definiti test di parità e fallback che escludano interferenze con FATTO/WHAZZUP e con gli impianti operativi.

La copia canary `varga-original-canary` serve gli stessi file dell'app dal checkout, senza modificarli. Nginx inietta soltanto configurazione, adattatore e `oracle-canary-bootstrap.js` nella risposta HTML della copia locale. Il bootstrap non modifica il DOM: dopo il login confronta gli ID delle commesse già caricate dall'app con il catalogo Oracle e salva il risultato nel solo `localStorage` del canary. La porta `18085` resta vincolata a loopback e le directory tecniche o sensibili non sono servite.

Il confronto considera riuscita la copertura quando ogni commessa visibile all'utente autenticato è presente nel catalogo Oracle. L'uguaglianza completa viene registrata separatamente: l'applicazione applica filtri di assegnazione e può quindi caricare intenzionalmente un sottoinsieme delle 20 commesse principali.

Il catalogo espone inoltre una allow-list di campi operativi: nome/codice/stato e relazioni non personali delle commesse; denominazione, ID SAP, comune, indirizzo, distretto, tipologia, voce di riferimento, frequenza e consistenza verde degli impianti. Email operatore, autori delle modifiche, identificativi Google Sheet e payload JSON originali non sono esposti. Al controllo del 26 ottobre 2026 risultano 19 commesse nominate su 20, 67 impianti nominati su 67 e 67 ID SAP su 67.

Il reverse proxy del canary usa `Referrer-Policy: strict-origin-when-cross-origin`, necessaria per rispettare la policy del provider delle tessere OpenStreetMap. Non sono state apportate modifiche al layout o al codice cartografico dell'app.
