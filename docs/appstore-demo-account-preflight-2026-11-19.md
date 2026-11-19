# VARGA CANTIERI — Preflight account dimostrativo per Apple (19/11/2026)

**Stato:** controllo **solo nel ramo PR draft #1606**, non eseguito su Firebase produzione. Nessun account è stato creato, modificato o eliminato da questa procedura. La sua riuscita **NON autorizza** il deploy di `deleteOwnVargaAccount`.

## Perché serve

Apple deve poter osservare su un dispositivo iOS vero una cancellazione completa. La Cloud Function è potenzialmente distruttiva: bisogna assicurarsi che l'identità impiegata nella dimostrazione non corrisponda a un lavoratore, amministratore o account collegato a commesse e documenti reali.

La verifica è preliminare e prudenziale, non un inventario di tutte le fonti Firebase, Storage, Google Drive o backup. Firestore non impone uno schema universale e gli identificativi legacy possono utilizzare campi diversi.

## Controllo automatico preparato

- Script: `scripts/check-app-review-demo-readiness.cjs`.
- Test senza credenziali: `npm run check:demo-readiness`.
- Workflow protetto: `.github/workflows/app-review-demo-readonly-preflight.yml`.

Il workflow è **manuale**, limitato al branch `main` e a uno SHA esatto, con approvazione nell'environment già protetto `firebase-account-deletion-production`. Per avviarlo in futuro servono la conferma `APPROVE_DEMO_METADATA_READ_ONLY`, l'UID del **solo** account usa-e-getta e la revisione umana dell'ambiente. Poiché ora il workflow è soltanto nel ramo draft, **NON è avviabile per la produzione tramite un normale workflow_dispatch fino a quando il codice non venga reso disponibile in modo approvato su main**. Non fare merge solo per attivarlo: il merge presenta rischi di deploy automatico di altre funzioni.

Prima di accedere a Firestore, lo script verifica l'UID e l'Auth custom claim attendibile `appReviewDemo === true` del profilo Firebase Authentication. Questo claim deve essere imposto **esplicitamente da un amministratore autorizzato tramite Firebase Admin SDK**, solo dopo aver identificato un account nuovo e usa-e-getta. Non è sufficiente cambiare displayName, email o un campo scrivibile dal client. Non impostare questo claim su utenti reali, admin, dipendenti o account di lavoro.

Dopo la verifica del claim, lo script effettua esclusivamente letture limitate ai riferimenti UID nelle collezioni `personale`, `documents`, `chatMessages`, `appNotifications`, `notifications`, `userAlerts`, `operatorCommessaHistory`, `oreReports`, `gpsUpdateRequests`, `notificationEmailQueue`, `activityLogs`. Controlla inoltre l'esistenza di documenti con chiave UID in `privateDocuments`, `operatorPositions` e `userOperatorLinks`.

Gli output sono esclusivamente booleani, numeri di categorie e nomi di categorie che richiedono revisione: mai UID, nomi, messaggi, email, coordinate o ID documento. Errori di permesso/lettura, claim mancante e dati collegati devono fermare il preflight.

**Attenzione:** le ricerche per `userId`, `linkedUserId` ecc. non rilevano tutte le possibili varianti di schema, dati in testo libero o esportazioni. Anche con `preliminaryGatePassed: true`, `grantsDeletionAuthorization: false` e `comprehensiveDataInventory: false` rimangono sempre false.

## Sequenza sicura per il collaudo reale (ancora NON avviata)

1. Completare l'audit GDPR/retention e definire titolarità, backup e dati aziendali con i responsabili autorizzati.
2. Preparare un'**identità Apple di prova**, non associata all'account Apple personale usato da altri sistemi. Registrare con essa un nuovo account VARGA CANTIERI senza attribuire commesse, aziende, ore, personale, documenti o privilegi amministrativi.
3. Far verificare l'identità del demo a un amministratore Firebase; soltanto dopo attribuire il custom claim `appReviewDemo: true`, preservando ogni altro claim esistente e senza esporre credenziali nel repository.
4. Eseguire l'inventario preliminare *in sola lettura* approvato: eventuali dati collegati causano uno stop e devono essere valutati manualmente. Il preflight non è un consenso alla cancellazione.
5. Ottenere signoff privacy e gate di produzione documentati. Distribuire la Cloud Function con il workflow mirato e protetto soltanto quando le verifiche precedenti sono complete; non fare merge o deploy indistinti.
6. Su **iPhone fisico**, provare preflight applicativo, accesso Apple, doppia conferma, revoca delle credenziali e cancellazione dello **stesso** account demo. Verificare su Firebase che l'utente sia assente; verificare separatamente che un altro account demo e i dati condivisi siano intatti.
7. Registrare e ripulire il video senza token/PII; inviarlo ad Apple con una nuova build **solo dopo** che i test hanno dimostrato il risultato.

## Criterio di stop

Se è necessario usare un account reale, se i dati possono appartenere a terzi, se manca il custom claim affidabile, se la lettura fallisce, se non sono verificate le fonti legacy o se l'app non può confermare `deleted: true`, **non procedere** con una cancellazione di prova e non dichiarare la funzione operativa ad Apple.
