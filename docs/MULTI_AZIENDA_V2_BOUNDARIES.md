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

## Alternativa concreta a database multipli: percorsi originali, tenant esplicito

Il 26 novembre, il confronto con le funzioni sigillate ha confermato che setImpiantoDone usa il percorso `commesse/{id}/impianti` (o neve), `impiantiFisici`, `lavorazioni` e batch; isImpiantoPersistedAsDone usa lo stesso percorso. Le funzioni di coda Whazzup accedono direttamente a localStorage. Non è ammesso avvolgerle o riscriverle.

Esperimento separato, non attivo nel runtime: mantenere i percorsi originali in un solo database, ID commessa univoci globalmente, metadato tenant sul padre e autorizzazione server dell'azienda selezionata. Le query globali non filtrate devono fallire, mentre quelle esplicitamente filtrate devono leggere solo l'azienda attiva. Le sottocollezioni verificano il padre. Il test tenta batch misti e verifica che non siano parzialmente applicati, nega cambi tenant arbitrari, verifica revoca e cambio azienda. Le sole mutazioni consentite sono marker sintetici; le scritture FATTO vere vengono negate intenzionalmente. Non certifica la compatibilità FATTO, non implementa una variante del pulsante, non modifica app.js e non cambia lo schema esistente.

Questa ipotesi evita di presupporre nuovi database a pagamento, ma richiede una revisione delle query non sigillate, regole operative complete (quote e campi autorizzati), ID univoci con relazioni intatte, separazione di tutte le raccolte globali, cache e coda offline. Una selezione server per UID, come nel test, ha il limite delle sessioni contemporanee su più dispositivi; serve una soluzione prima dell'uso reale. Verificare anche storage, integrazioni, prezzari, aggregati e notifiche. Non cambiare risolutori o cache protette per forzare l'integrazione. Non migrare Avola, non scegliere una architettura definitiva e non pubblicare sulla base della prova.

## Sessioni contemporanee e caricatore originale — nuovo blocco

Nuovo esperimento separato: percorsi Firestore originali commesse/neve_commesse, query tenant esplicita, contesto per singola sessione invece di selezione globale per UID. Le regole di laboratorio richiedono un claim sintetico firmato nel test, documento sessione con azienda, stato e scadenza, e appartenenza attiva; il client non può emettere, modificare o leggere sessioni. Il test esegue le funzioni originali subscribeCommesse, normalizeCommessaDocument e getCommesseCollectionName con l'SDK emulator, due sessioni dello stesso UID, commesse ordinarie/neve e prezziario, cambio di una sessione, revoca e scadenza. Nessuna scrittura operativa, nessun accesso a cache storiche, nessun deploy. I token per singolo dispositivo sono una ipotesi da verificare con Auth reale: le fixture sintetiche NON provano emissione, refresh o revoca di token reali, né sono una decisione architetturale definitiva. Il caricatore resta un adattatore di test in sola lettura; non collegare il runtime originale live o modificare funzioni sigillate. Il nuovo test è incluso nel workflow emulator; esito da confermare.

## Revoca: anomalia osservata e prova Auth effettiva

Il run 45861336379 è fallito al controllo successivo alla revoca della sessione telefonica: getDocsFromServer su una query già attiva ha restituito un risultato. Le prove precedenti nello stesso job sono passate. Non considerare la revoca in tempo reale risolta: una risposta di un listener già attivo può precedere la rivalutazione delle regole. Il nuovo test conserva un'osservazione esplicita di quel caso, chiude il runtime e richiede il rifiuto di una lettura diretta e di una query su un nuovo client. Se anche la richiesta nuova riesce, il test fallisce. Serve un protocollo di chiusura/invalidazione delle sessioni e della UI prima del live, indipendentemente dal risultato. La nuova prova separata usa Auth Emulator e Admin SDK per due login custom-token dello stesso UID, senza scrivere il claim nelle custom claims globali, verifica il refresh forzato e controlla la revoca su nuove connessioni. L'emissione locale non è un endpoint, non è un meccanismo anti-replay e non prova l'emissione o il refresh cloud. La durata di 30 minuti è solo un parametro del laboratorio. Nessun deploy, nessun accesso ad Avola e nessuna modifica al runtime o ai flussi protetti. Esito del nuovo workflow da confermare.

## Archivi originali e chiusura sessione — laboratorio

Run 45861708456: entrambi i job superati; token Auth emulator per singola sessione e refresh forzato verificati, nuove letture dopo revoca negate. L'osservazione della query già aperta è rimasta fulfilled: non è una soluzione alla revoca della UI. Il nuovo blocco consente la lettura del solo record di controllo della sessione associata al claim; nessuna enumerazione o modifica dal client. Monitor locale verifica stato/scadenza, ascolta il record, ripete una verifica server e richiama la chiusura. Il test usa un buffer sintetico, non la cache storica dell'app; la frequenza e il timeout sono parametri di laboratorio, non garanzia di revoca istantanea o scelta sui costi live. Aggiunte regole read-only su personale, mezzi, squadre correnti/storiche, ore e richieste (normale/neve). Si eseguono i sette selettori originali estratti da app.js senza cambiare il sorgente; test di query autorizzate, accessi incrociati, ruoli e scritture negate. I moduli operativi completi e l'installazione del monitor nel runtime originale restano da realizzare. Nessuna pubblicazione regole, migrazione o modifica FATTO/WHAZZUP. Esito da verificare in CI.

## Inserimento originale con quote server — laboratorio

Run 45864017649: browser passato; la suite archivi si è fermata al parsing delle funzioni estratte, prima di eseguire il monitor. Corretto il test con parser AST Acorn, installato solo tra le dipendenze di test, senza modificare app.js. Nuovo adattatore limitato a createCommessa e addManagementImpianto: esegue il codice originale con fixture UI e traduce esclusivamente i due payload iniziali nei servizi transazionali già esistenti. Usa il namespace annidato tenants, NON scrive anche nei percorsi flat. Verifica concorrenza normale/neve con quota una commessa, dieci impianti e coppia anagrafica fisica/proiezione, ruoli falsificati e accesso incrociato. L'adattatore è solo emulator, non gestisce altri handler, non è un sostituto del runtime e non può essere pubblicato. Il percorso flat rimane ipotesi separata; niente doppia scrittura o migrazione. Le richieste non hanno retry automatico e l'idempotenza tra nuovi click non è implementata. Nessun FATTO/WHAZZUP/RESET modificato. Esito CI da confermare.
