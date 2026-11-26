# Varga Cantieri 2.0 — inventario dei confini operativi

Stato: integrazione non ancora pronta per staging live o TestFlight. L'originale
app.js è invariato. Le prove precedenti hanno dimostrato il servizio multi-azienda
e il caricamento originale delle commesse in ambiente sintetico, non l'intera app.

## Confini già identificati nel codice originale

| Area | Evidenza | Lavoro necessario prima dell'attivazione |
|---|---|---|
| Commesse, impianti e FATTO | app.js: collection("commesse"), getCommesseCollectionName(), setImpiantoDone() | Definire un confine del datastore compatibile con il codice sigillato. Non cambiare resolver, wrapper o comportamento FATTO per aggirarlo. |
| Utenti e ruoli | platformUsers, appConfig/adminUsers, canManageData() | Distinguere identità globale e autorizzazioni aziendali. Approvazione e login originale protetti; una modifica del login richiede consenso esplicito. |
| Personale, mezzi, squadre, ore | Resolver normali/neve, squadreStorico, oreLocks, sharedStaticViews | Conservare relazioni, regole di modifica, deduplicazione e storico. Le anteprime sintetiche non sostituiscono i moduli originali. |
| Cache, stato e offline | COMMESSE_LOCAL_CACHE_KEY, LAST_OPENED_COMMESSA_KEY, PENDING_OFFLINE_MUTATIONS_KEY, WHAZZUP_PENDING_DONE_KEY | Isolare identità e azienda e mantenere le code originali. Non cancellare cache o azioni pendenti storiche. |
| Drive, Sheets, Storage | uploadCentralDriveFile, centralDriveFolders, appConfig, storage.rules | Verificare proprietà tenant e destinazioni, nessun link o credenziale condiviso tra aziende. |
| Notifiche, trigger, Gestionale | functions/index.js, functions/main.js, user-notifications, varga-gestionale-sync | Eventi e destinatari confinati all'azienda; nessuna distribuzione globale delle operazioni. |
| Release | firebase.json staging vuoto, Netlify build bloccata | Conservare blocchi di deploy fino al superamento delle prove native e alla verifica dei confini. |

## Decisione ancora aperta

I namespace tenant sono già verificati per il nuovo backend. Non basta sostituire
il database globale con un adattatore generico: ciò può alterare i flussi FATTO,
WHAZZUP e RESET protetti. Un datastore separato per azienda, con percorsi
originali, è una possibile alternativa da valutare; non è una decisione approvata
e può richiedere nuovi progetti/database, configurazione di identità e costi.
Non si creano risorse cloud aggiuntive e non si abilita fatturazione.

## Inventario ripetibile

`node tests/multi-tenant/legacy-boundary-inventory.cjs` analizza file JavaScript
e regole del repository, senza connessioni ai servizi. Esegue una ricerca
euristica di path, resolver, callable, chiavi cache, storage, trigger e servizi
esterni. Il report contiene file/riga e identificatori, non righe sorgente,
credenziali, contenuti dei dati o host specifici. I riferimenti dinamici non
riconosciuti, native plugins e configurazioni richiedono ancora revisione
manuale. L'inventario non certifica l'isolamento. Il workflow produce un
artefatto per la revisione sul ramo di sviluppo.

Non migrare dati, non pubblicare e non dichiarare l'app pronta sulla base di
questo inventario.

## Verifica dei costi della separazione fisica
La documentazione ufficiale Firestore (https://firebase.google.com/docs/firestore/pricing, consultata il 26 novembre 2026) conferma una sola banca dati con quota gratuita per progetto; ulteriori database richiedono l'upgrade del piano di fatturazione. Non attivare database aggiuntivi o fatturazione senza una decisione separata. Il laboratorio resta a costo cloud nullo ed emulator-only.
