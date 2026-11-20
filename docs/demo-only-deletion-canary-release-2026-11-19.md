# VARGA CANTIERI — rilascio canary della cancellazione (solo demo)

**Data:** 19 novembre 2026

## Perché questo rilascio è limitato

La build TestFlight 1.0 (20011) ha superato il collaudo di **Accedi con Apple** su iPhone fisico, ma la cancellazione account non è stata ancora distribuita né provata con un account di test. L'audit dei dati di aziende e lavoratori, la conservazione esterna e la privacy restano incompleti: NON approvare per questo motivo il rilascio generale e NON affermare che Apple ha approvato la pubblicazione.

## Fase 1: canary senza cancellazioni di utenti reali

Il backend `functions/account-deletion.js` ora impone tre condizioni prima di eseguire qualsiasi pulizia di documenti:

1. **Doppio controllo sui Firebase ID token:** `appReviewDemo === true` e `apple_review === true`, creati soltanto dal workflow protetto per la demo;
2. **Controllo server-side:** anche il record Firebase Auth letto con Admin SDK deve possedere entrambi i claim;
3. **Confronto SHA-256 dell'email del record Auth** con l'alias disposable già autorizzato. Non basta una email scelta dal client né il parametro `uid`.

Il valore `DEMO_ONLY_RELEASE = true` è fisso nel sorgente: non esiste un input GitHub o una variabile di ambiente in grado di attivare gli utenti reali. In assenza di queste condizioni la funzione termina con `permission-denied` prima di leggere o modificare Firestore.

Il workflow manuale `deploy-account-deletion-demo-canary.yml` è separato dal rilascio generale, richiede esatto SHA del ramo `main`, approvazione esplicita, ambiente protetto `firebase-account-deletion-production` con reviewer, e lettura limitata del demo DISABILITATO con verifica di 20 possibili collegamenti al lavoro. Distribuisce **solo** `functions:deleteOwnVargaAccount`. Se il demo non è isolato, non viene distribuito.

### Cosa prova il workflow (e cosa NO)

- Prova: versione di codice demo-only, controlli sintetici, identità demo disabilitata e senza i 20 riferimenti conosciuti, distributore Firebase limitato a una singola callable.
- **Non prova:** inventario esaustivo di tutte le collezioni, copia Drive o backup, affidabilità della revoca Apple reale, cancellazione su iPhone, legalità e titolarità di dati di terzi.
- Il deploy è un rilascio canary, **non un'attestazione `PRIVACY_REVIEW_APPROVED`**, che rimane obbligatoria per il rilascio generale.

## Fase 2: collaudo reale demo

Dopo il deploy canary, verificare la presenza della callable in `europe-west1` e l'esclusione esplicita di utenti non-demo. Attivare temporaneamente **soltanto** l'account demo dopo una nuova verifica di isolamento e con reviewer; non attivare utenti reali. Servono credenziali demo effettivamente utilizzabili e la prova su iPhone per il flusso Apple; la demo attuale è DISABILITATA e non si possono inventare login/password.

La UI della beta 20011 resta nascosta perché `window.__VARGA_ACCOUNT_DELETION_ENABLED` non è impostata a true. Per il video serve una build successiva che mostri la cancellazione esclusivamente nel contesto di test appropriato.

## Ulteriori protezioni nel codice (20 novembre 2026)

Prima di una futura valutazione del rilascio generale, il canary applica anche due controlli fail-closed:

- La pre-verifica dei file personali accetta soltanto percorsi Firebase Storage canonici `documents/{uid}/{documentId}/{filename}` con UID del richiedente, quattro segmenti semplici e nome file compatibile con `documents.js`/`storage.rules`. Percorsi annidati, traversal `..`, riferimenti Drive/legacy, nomi malformati o oggetti che sembrano appartenere a un'altra persona producono `failed-precondition`, **prima di cancellare qualsiasi dato**. I riferimenti storici non canonici richiedono revisione manuale, non una cancellazione forzata.
- La funzione accetta `mode: "preflight"` per il controllo non distruttivo e **assenza di `mode`** per la richiesta di cancellazione espressamente confermata. Qualsiasi altra modalità viene rifiutata con `invalid-argument`; non può cadere accidentalmente nella procedura irreversibile.

I test `scripts/check-account-privacy-runtime.js` verificano con stub e identità **sintetiche** la mancata cancellazione di Auth, Storage, Firestore e comunicazioni condivise in questi casi. Il controllo non equivale alla prova su iPhone e **non abilita account ordinari**: il blocco demo-only, il workflow protetto e la revisione privacy rimangono obbligatori.

## Fase 3: rilascio generale e App Store

Solo dopo l'inventario GDPR, titolarità con le aziende, backup/export, politica di minimizzazione e collaudo end-to-end documentati, proporre una **nuova PR di promozione** che rimuova il blocco demo, con revisione e test separati. Usare il workflow `deploy-account-deletion-protected.yml` con signoff legale vero, poi abilitare il pulsante per gli utenti reali in una nuova build. Non auto-approvare i gate, non cancellare account di dipendenti e non inviare una revisione App Store che dichiari completata la cancellazione quando ancora non lo è.

Riferimenti: issue [#1609](https://github.com/ionut290/hera-app/issues/1609) e test della build [#11](https://github.com/ionut290/hera-app/actions/runs/44607013274).
