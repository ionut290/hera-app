# VARGA CANTIERI — richiesta di cancellazione autenticata (solo preparazione)

**Stato al 20 novembre 2026: NON DISTRIBUITA e NON ABILITATA.** Non usare questo documento come prova di conformità App Store o GDPR.

## Perché esiste questa procedura

Apple consente che l'eliminazione richiesta dall'app sia completata manualmente entro un periodo ragionevole, a condizione che l'utente possa iniziare dall'app, sia informato sui tempi e riceva conferma al completamento. Una semplice email obbligatoria all'assistenza, come nell'attuale `delete-account.html`, generalmente non basta per un'app non appartenente a settori altamente regolati.

Riferimento Apple: https://developer.apple.com/help/app-review/guideline-reference/5-1-1-account-deletion/

L'opzione preparata qui permette a un utente **autenticato** di avviare una richiesta non distruttiva, mentre i dati operativi/contabili con conservazione necessaria vengono esaminati prima di qualsiasi rimozione.

## Implementazione tecnica preparata

- `functions/account-deletion-request.js` contiene `requestOwnVargaAccountDeletion` e `getOwnVargaAccountDeletionStatus`. Richiedono Firebase Authentication valida; per inviare una nuova richiesta serve un accesso recente (massimo 10 minuti) e una conferma esatta. UID e destinazione Firestore sono ricavati **esclusivamente dal token verificato**, non da input del client.
- Il server registra solo `uid`, `status=pending`, orario server, origine autenticata e indicazione che serve una valutazione prima della cancellazione. Non copia email, numeri di telefono, nomi dei dipendenti o dati di lavoro. Le richieste ripetute sullo stesso UID non generano duplicati né azzerano una richiesta in revisione.
- La collection `accountDeletionRequests` **non compare** tra quelle consentite nel fallback di `firestore.rules`: l'app non può leggere/scrivere richieste di altri utenti. L'accesso avviene solo via Cloud Function protetta dall'identità del chiamante.
- `account-privacy-controls.js` mostra l'azione «Richiedi cancellazione account» in profilo, menu, vista utente in attesa e accesso **solo** se `window.__VARGA_ACCOUNT_DELETION_REQUEST_ENABLED === true`, se l'eliminazione immediata non è attiva e se `window.__VARGA_ACCOUNT_DELETION_REQUEST_SLA_DAYS` è un intero numerico da 1 a 90 approvato e conforme al processo reale. Per impostazione predefinita il pulsante è invisibile. La ricevuta usa solo la durata approvata, senza promettere automaticamente 30 giorni.
- Il client prepara anche un pulsante «Verifica stato cancellazione», che legge tramite `getOwnVargaAccountDeletionStatus` lo stato del **solo** utente autenticato, senza creare nuove richieste. Mostra «in attesa», «in revisione», «completata» (senza presumere di aver ricevuto una conferma ufficiale) o «chiusa senza cancellazione». In caso di errore non indica falsamente che la cancellazione sia riuscita. Anche questo pulsante è nascosto senza il flag di rilascio.
- Le due nuove callable **non sono esportate da `functions/main.js`** e il normale deploy Firebase esclude `functions/account-deletion-request.js`. **Non sono raggiungibili dagli utenti su Firebase** finché non vengono integrate, distribuite e verificate tramite un rilascio separato approvato.
- `scripts/check-account-deletion-request-staged.cjs` e `scripts/check-account-privacy-client-runtime.js` esercitano la funzione reale in VM con dati fittizi e l'interfaccia con stub Firebase, verificando autenticazione, conferma, isolamento per UID, idempotenza, mancata eliminazione dei dati e messaggi corretti. CI: `check-account-deletion-request-staged.yml`.

## Procedura che deve esistere PRIMA dell'attivazione

1. Il titolare del trattamento competente convalida i ruoli delle diverse aziende, inventario dati, finalità e conservazioni (compresi `personale`, `oreReports`, GPS, chat, documenti condivisi, Storage, esportazioni Google Drive, copie di backup e comunicazioni esterne).
2. Viene individuato il responsabile operativo **reale** del trattamento delle richieste, una coda controllata e il processo tracciato di presa in carico, verifica delle eventuali eccezioni di conservazione, eliminazione/anonimizzazione e chiusura. Il responsabile deve poter dare conferma verificabile all'utente, non soltanto cambiare lo stato Firestore.
3. Devono essere approvati termini di presa in carico, completamento, gestione di rifiuti motivati, retention delle richieste stesse e notifiche al richiedente. L'attuale pagina pubblica `delete-account.html` menziona «normalmente entro 30 giorni». Questa promessa va verificata o corretta prima di attivare la funzione in-app e allineata alla durata effettivamente approvata; il solo valore nel client non costituisce prova di un processo reale.
4. Per accessi Apple occorre completare e collaudare la revoca dei token nel processo di cancellazione effettiva. Registrare la richiesta non equivale a revocare il provider.
5. Dopo la verifica privacy, esportare le funzioni da `functions/main.js` in una modifica controllata, distribuire **solo** quelle con un workflow protetto, provare con un utente fittizio separato e poi impostare il flag client esplicitamente.
6. Dimostrare su **iPhone fisico**: pulsante visibile, richiesta con account fittizio, messaggio di ricevuta, gestione amministrativa, cancellazione completata e confermata, dati riservati di terzi esclusi. Infine aggiornare la build TestFlight e le note App Store Connect con credenziali valide.

## Cosa NON è stato eseguito

Nessun deploy Cloud Functions, nessun account creato/eliminato, nessuna lettura o scrittura Firebase, nessuna assegnazione di ruolo e nessuna dichiarazione di completamento Apple. Il workflow di pubblicazione `ios-appstore-candidate.yml` mantiene il blocco `--require-ready`.

**Rischio ancora aperto:** finché la richiesta non viene presa in carico da un processo reale verificato, anche l'avvio autenticato dall'app rimane una bozza tecnica e non una funzione pronta a soddisfare la revisione Apple.
