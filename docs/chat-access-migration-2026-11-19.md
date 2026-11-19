# VARGA CANTIERI — Migrazione accesso chat privata (19 novembre 2026)

**Stato:** codice e test su PR draft #1606. Nessuna regola/index e nessun client ancora distribuiti. Prima del rilascio serve collaudo coordinato.

## Problema osservato

La versione precedente carica una query globale su `chatMessages` (`orderBy(createdAt).limit(500)`) e nasconde i messaggi privati solo con `canViewMessage()`. Le regole Firestore precedenti consentivano `read,create,update` a tutti gli utenti autenticati. La protezione grafica non impedisce letture API, né modifiche abusive del destinatario.

## Modifiche proposte nel ramo

- Firestore: lettura broadcast solo quando `recipientId` è vuoto/mancante; altrimenti autore, destinatario o amministratore. Scrittura di un messaggio con `senderId` uguale all'UID del chiamante (o amministratore); modifica/eliminazione riservata agli amministratori.
- Client `subscribeChat()`: tre listener separati (broadcast, destinatario, autore), unione per ID, limite/sort per data, niente notifiche sui primi snapshot. Non esegue più download globale per gli utenti ordinari.
- Due indici compositi `chatMessages` (`recipientId + createdAt`, `senderId + createdAt`).
- Test emulator con messaggi sintetici: accesso partecipanti, negazione a terzi, impossibilità di spoofing del mittente, tombstone del destinatario eliminato, query globali negate e query filtrate autorizzate. Controlli statici anti-regressione.

## Ordine obbligatorio del rilascio (non avviato)

1. Eseguire test CI + emulatori sul commit esatto e confermare non-regressione dei flussi principali.
2. Distribuire **soltanto gli indici Firestore** e attenderne lo stato Ready prima di distribuire il client o cambiare le regole.
3. Rilasciare la versione web/iOS/Android con query per partecipante e verificarla su account sintetici. Le versioni vecchie con la query globale falliranno dopo il cambio regole: definire compatibilità, aggiornamento obbligatorio o disattivazione mirata delle vecchie versioni.
4. Applicare le regole Firestore restrittive in una finestra controllata, poi ripetere verifiche di permessi, messaggi pubblici e privati e amministrazione; monitorare errori `permission-denied`.
5. Verificare separatamente la policy di Storage/Drive degli allegati e i messaggi legacy senza `recipientId`. L'archivio chat e le notifiche di tipo `notification_ack` possono avere restrizioni operative ulteriori; una query broadcast non equivale a prova della corretta classificazione di ogni record.
6. Solo dopo la sicurezza chat, completare l'audit privacy della cancellazione account, test demo e video fisico richiesto da Apple.

## Limiti e decisione

- Gli amministratori mantengono volutamente accesso alle chat come dato aziendale: questo privilegio va verificato e documentato con il titolare competente.
- Gli URL allegati condivisi su Drive/Storage non sono automaticamente protetti dalle sole regole chat.
- Le chat con recipientId mancante non compariranno nella nuova query broadcast finché non saranno migrate/classificate; non dedurre che siano tutte pubbliche.
- Il codice client contiene un precedente ciclo di pulizia chat di circa 24 ore: questo non dimostra l'attuazione delle proposte di retention per i dati aziendali.
- **Blocco distribuzione:** non distribuire regole restrittive prima della migrazione dei client. Non unire PR #1606 né distribuire `deleteOwnVargaAccount` per effetto di questo documento.


## Blocco automatico e fasi protette — modifica codice 19/11/2026

**Misura applicata nella PR (NON ancora in produzione):** `.github/workflows/deploy-firebase-functions.yml` non pubblica più regole o indici Firestore automaticamente su push di `main`. Continuerà a occuparsi soltanto delle Cloud Functions espressamente elencate. Questo evita che il merge della PR renda i vecchi client incompatibili con i permessi dei messaggi privati.

Il nuovo workflow `.github/workflows/deploy-firestore-chat-protected.yml` richiede dispatch manuale su `main`, esatto SHA revisionato e un environment GitHub **già creato** con nome `firebase-private-chat-production` e reviewer obbligatori. Se il nuovo environment non esiste o non è protetto, il job viene rifiutato **prima** di accedere a credenziali Firebase. Non autorizzare una creazione automatica senza reviewer.

1. **Indici soltanto** (`APPROVE_INDEXES_ONLY`): pubblica `firestore.indexes.json` senza cambiare i permessi di lettura. Attendere che gli indici risultino entrambi `READY`.
2. **Distribuzione client**: pubblicare le nuove query chat su web/iOS/Android in modo controllato, verificando tutti i client ancora utilizzati. Non presumere che una beta TestFlight aggiorni automaticamente ogni dispositivo aziendale.
3. **Regole soltanto** (`APPROVE_RULES_ONLY`): richiede `PRIVACY_REVIEW_APPROVED`, `CLIENTS_MIGRATED` e verifica via **Firestore Admin API in sola lettura** che gli indici di `recipientId` e `senderId` con `createdAt` siano effettivamente `READY`. Il controllo è implementato da `scripts/check-chat-indexes-ready.cjs`; il test locale `--self-test` non accede a Firebase.
4. **Convalida post-pubblicazione**: messaggi privati accessibili soltanto agli interessati, broadcast visibili, notifica senza duplicati, non interrompere storico e operazioni aziendali. Eventuali `permission-denied` su vecchi client vanno risolti prima della finestra di rilascio.

**IMPORTANTE:** tutte queste fasi sono predisposte come codice. Non sono state attivate né è stata verificata la presenza dell'environment `firebase-private-chat-production` o la reale migrazione di tutte le installazioni. Un merge non conclude la migrazione.


## Verifica del vecchio workflow di registrazione — 19 novembre 2026

Durante l'audit è emerso un **secondo percorso di pubblicazione automatica**: `.github/workflows/deploy-self-registration.yml` eseguiva `firebase-tools deploy --only firestore:rules` a ogni push su `main` che modificava le regole, aggirando la nuova procedura a fasi. È stato convertito in workflow di **sola verifica** `Verify Firestore registration rules (NEVER deploy)`, senza credenziali, secrets o comando deploy. Il controllo `scripts/check-account-privacy.js` impedisce la reintroduzione del vecchio deploy. Il nuovo workflow di rilascio protetto `deploy-firestore-chat-protected.yml` rimane l'unico percorso **previsto** per cambiare queste regole.

L'environment GitHub `firebase-private-chat-production` è stato **creato e verificato dal browser** con reviewer `ionut290`, distribuzione consentita soltanto da `main`, bypass amministrativo disabilitato. **Nessun deploy Firestore è avvenuto** in questa preparazione.

## Ulteriori regole per coordinate GPS ed email (19 novembre 2026)

L'audit ha individuato che `gpsUpdateRequests` e `notificationEmailQueue` erano comprese nel fallback `allow read, write: if signedIn()`. **Il solo controllo visuale non protegge Firestore.**

Le regole proposte nella PR eliminano entrambe dal fallback:
- `gpsUpdateRequests`: un operatore autenticato può creare solo richieste `pending` col proprio `operatorId`, leggere la propria richiesta e interrogare solo il proprio `operatorId` con limite massimo 100; un amministratore può visualizzare l'elenco completo e approvare/rifiutare. Nessun altro lavoratore può leggere le coordinate.
- `notificationEmailQueue`: lettura/scrittura esclusiva per l'amministratore autenticato. Il codice esistente crea le email dal pannello di gestione; eventuali consumer privilegiati tramite Admin SDK devono essere verificati separatamente.
- Il client non avvia più il listener globale delle richieste GPS per un operatore. Anche eventuali callback ritardati del vecchio account amministrativo vengono ignorati e i dati vengono azzerati.
- Prove automatiche Auth/Firestore Emulator e test comportamentale con utenze sintetiche verificano accessi autorizzati, negazioni, creazione di false approvazioni e isolamento dopo cambio utente.

**Rischi di compatibilità da controllare prima del rilascio:** versioni mobili/web obsolete che leggono l'elenco GPS globale o scrivono la coda email fuori dall'interfaccia amministrativa; integrazioni webhook; usi privilegiati non documentati. Non aggiornare le regole in produzione prima delle verifiche e della migrazione dei client.

**Limiti:** non elimina coordinate già salvate, email già inviate, copie Drive o backup. La retention e gli accordi GDPR restano aperti. Questa è una proposta di codice, non un rilascio già avvenuto.


## Verifica aggiuntiva delle comunicazioni operative — 19/11/2026

Il controllo della PR ha trovato altre due raccolte nel permesso generico `allow read, write: if signedIn()`:

- `appNotifications`: per preservare le notifiche operative globali, la lettura rimane ai lavoratori con account attivo. La creazione di eventi da un utente ordinario è ora consentita **solo** con `createdByUid` uguale al proprio UID autenticato, un `eventType` appartenente agli eventi realmente usati dal client e un `createdAt` server-side. Gli eventi di sistema, così come le modifiche e le cancellazioni, richiedono amministratore. L'app pubblica eventi anche per FATTO, navigazione, ore, squadre e vicinanza/allontanamento impianto: la visibilità generale dei dettagli operativi ai lavoratori **è una scelta di prodotto che necessita valutazione privacy con il titolare**; l'autorizzazione attuale non applica destinatari per singola commessa.
- `hoursDeadlineAlerts`: il controllo automatico delle ore mancanti viene eseguito dal client solo se `canManageData()` è vero; pertanto soltanto gli amministratori possono creare/modificare/eliminare i record di allerta. La lettura continua per gli utenti attivi.

Aggiunte prove Firebase Emulator con operatori, amministratori, utenti non autenticati e account di revisione Apple, oltre a verifiche statiche contro il ripristino del fallback generico. **Le nuove regole non sono state distribuite.** Devono attraversare gli stessi passaggi manuali protetti indici/client/regole descritti sopra. Prima dell'attivazione verificare gli eventi legacy, i client ancora installati e la coerenza della visibilità globale con le basi giuridiche aziendali.


## Conferme di lettura personale — ulteriore controllo Firestore

La collezione `userAlertAcknowledgements` conteneva identificativi e nomi dei lavoratori ma utilizzava il vecchio `allow read, write: if signedIn()`, che permetteva a ogni account attivo di leggere o modificare le conferme altrui. Nel ramo della PR si applica una regola esplicita:

- Ogni lavoratore può eseguire `get` soltanto sul documento con ID `<notificationId>__<proprio UID>`, anche prima che il documento esista, come richiesto dall'interfaccia.
- La creazione è ammessa soltanto per l'ID deterministico, con `userId` corrispondente all'autenticazione e timestamp server.
- La lista completa per data, la modifica e la cancellazione sono riservate agli amministratori, che usano il calendario di conferma in gestione.
- La vecchia autorizzazione generica è eliminata; emulatori Firebase testano account proprietario, altro lavoratore, reviewer Apple e amministratore.

Questa restrizione non risolve da sola la raccolta `userAlerts`: il client ordinario carica ancora il feed generale e filtra sul dispositivo le notifiche dirette ad altri utenti. Per evitare interruzioni o mancata consegna degli avvisi legacy basati su nomi di squadra, occorre **progettare e collaudare una migrazione della lettura con regole server-side e query per destinatari** prima di applicare restrizioni simili. La revisione di `userAlerts` rimane un elemento privacy **aperto**, non una funzione dichiarata sicura. Non distribuire automaticamente la nuova regola delle conferme a client che non hanno superato i test di compatibilità.


## Inventario obbligatorio prima di separare gli avvisi personali `userAlerts`

La lettura dei messaggi `userAlerts` rimane oggi globale con filtro destinatari nel dispositivo, anche per notifiche con contenuti personali (inclusi gli avvisi automatici sulle assenze). Limitare immediatamente `allow read` in Firestore bloccherebbe i vecchi client che usano la query generale. Inoltre la Cloud Function `functions/user-notifications.js` crea gli avvisi `calendar-absence` basandosi su `targetMemberNames` (nomi testuali), non UID Firebase. **Non associare automaticamente nomi e UID con corrispondenza fuzzy**: nomi omonimi, alias o squadre variate potrebbero consegnare una notifica al dipendente sbagliato.

È stato aggiunto `scripts/audit-user-alert-targeting.cjs` con due modalità:

- `node scripts/audit-user-alert-targeting.cjs --self-test`: test di classificazione con sole fixture sintetiche; parte della CI `check:account-privacy`.
- `ALLOW_READ_ONLY_USER_ALERT_AUDIT=1 GOOGLE_APPLICATION_CREDENTIALS=<credenziale> node scripts/audit-user-alert-targeting.cjs --live`: futuro audit *in sola lettura*, bloccato senza esplicita variabile di abilitazione e credenziale con `project_id=hera-app-6cd2b`. Richiede `firebase-admin` installato nella cartella `functions`. La query seleziona solo i campi di indirizzamento `sendToAllRegistered`, `targetUserIds`, `targetUserId`, `targetMemberNames` e `source`; stampa esclusivamente i **conteggi aggregati**, senza messaggi, nomi o ID. **Non è stato lanciato su Firebase produzione.** Programmarlo solo con un account di servizio a privilegi minimi e procedura privacy approvata.

Classificazioni: broadcast, destinatari UID espliciti, UID + nomi legacy, soli nomi legacy non risolti, destinatari sconosciuti, metadati malformati. `needsReview>0` restituisce exit code 2 e blocca l'abilitazione di regole restrittive. Un risultato `needsReview=0` non è sufficiente da solo: restano da verificare destinatari sbagliati, ID non più attivi, client vecchi e indici.

**Procedura di migrazione prevista (non eseguita):** inventariare gli avvisi; validare con un amministratore i `targetMemberNames` contro anagrafiche e UID effettivi; correggere le Cloud Functions che producono assenze per scrivere destinatari UID deterministici; distribuire gli indici e client con query separate (broadcast, UID array, UID singolo); solo dopo i test sul dispositivo e la migrazione degli avvisi storici, applicare regole server-side che non consentano più la query globale. Necessita conferma del titolare del trattamento e non deve modificare automaticamente avvisi personali o dati aziendali.


### Audit con approvazione tramite GitHub Actions

In aggiunta al comando manuale locale, è stato preparato il workflow **`.github/workflows/audit-user-alert-targeting-protected.yml`** (non eseguito):
- Avvio solo manuale `workflow_dispatch` su `main`, con `APPROVE_METADATA_AUDIT`, `RECIPIENT_METADATA_REVIEW_APPROVED` e SHA di commit esatto.
- Environment protetto `firebase-private-chat-production` con reviewer obbligatorio già creato. Il job verifica nuovamente che l'environment abbia regole `required_reviewers` valide. Se manca, il processo viene interrotto.
- Il servizio accede solo per leggere i **metadati di destinatari** con identità Firebase del progetto previsto; i risultati del job sono solo conteggi. Non archivia documenti, non genera file di esportazione, non cancella o modifica Firestore e distrugge la credenziale temporanea a fine esecuzione.
- Un risultato con `needsReview > 0` blocca il job intenzionalmente. Nessuna migrazione automatica per associazione dei nomi, per evitare di divulgare assenze o comunicazioni ad omonimi.

**Stato:** soltanto codice e test sintetici nella PR draft. Non è stato richiesto, approvato o eseguito alcun accesso ai documenti `userAlerts` di produzione.


## Destinatari delle push Android — invio fail-closed (19/11/2026)

È stata individuata un'ulteriore fonte di possibile divulgazione: `functions/user-notifications.js` usava `userMatches()` con `includes()` per associare operatori, ore e destinatari delle chat a profili Firebase, senza gestire correttamente nomi omonimi, alias o token FCM duplicati. Un match parziale può essere ambiguo, soprattutto per nomi corti.

La modifica nel ramo della PR:
- Per chat **private**, risolve `recipientId` soltanto come **UID Firebase esatto e con distinzione tra maiuscole/minuscole**, oppure come **email legacy esatta e univoca** (email senza distinzione di maiuscole/minuscole). Non invia mai contenuto chat a corrispondenze approssimate per nome. I messaggi generali restano broadcast. Anche il confronto con il `senderId` è case-sensitive per non scambiare utenti distinti.
- Per push individuali sulle **ore** o sull'**assegnazione a squadre**, preferisce l'identità esatta; in sua assenza consente un match legacy parziale **solo se unico tra tutti i profili**, inclusi quelli senza push token. In caso di omonimia o alias ambiguo non invia alcuna push personale: la registrazione aziendale resta disponibile nel gestionale, ma la notifica richiede una successiva correzione dell'anagrafica.
- Un token FCM associato a più profili, o mancante, non riceve push personali, evitando il rischio di recapitare informazioni del lavoratore a un dispositivo associato a un altro account.
- `scripts/check-push-recipient-isolation.cjs` contiene test con profili sintetici, nomi duplicati e token condivisi, agganciati a `check:account-privacy`.

**Compatibilità e limiti:** i lavoratori con identificativi legacy ambigui potrebbero non ricevere alcune push finché non vengono associati a UID univoci; non viene cancellato o modificato alcun dato aziendale. I test non inviano push vere. Questa modifica **non risolve** il client `userAlerts` a query globale né autorizza la distribuzione della PR senza i controlli GDPR e l'account demo Apple.


## Identità profilo come origine delle notifiche — controllo aggiuntivo

Nell'ulteriore verifica delle push Android è emerso che `platformUsers/{UID}` contiene campi `uid` e `email` e che il server usava `data.uid || doc.id` per decidere chi riceve una notifica privata. Le regole per l'aggiornamento del profilo non impedivano esplicitamente a un dipendente di modificare questi campi, e tale comportamento avrebbe potuto creare una falsa corrispondenza con un altro utente.

Nel ramo della PR:
- `functions/user-notifications.js` usa il **document ID Firestore**, che coincide con il vero UID Firebase, e non più il campo `data.uid` modificabile per determinare l'utente della push. Il campo `data.uid` viene escluso anche dagli alias ricercabili per l'instradamento. I profili esplicitamente **bannati, in attesa, rifiutati o con uno stato diverso da `attivo`** sono esclusi dalle push; i vecchi profili senza campo di stato mantengono la compatibilità con le regole di accesso legacy. Il confronto degli UID della chat privata resta esatto e distingue maiuscole e minuscole.
- In `firestore.rules`, il self-update del profilo consente di modificare `uid` **soltanto** per ripristinare il proprio `request.auth.uid`; un cambio di `email` è consentito soltanto se coincide con la nuova email già associata all'account autenticato Firebase. Altri cambiamenti ordinari al profilo continuano a essere supportati. L'amministratore conserva le autorizzazioni attuali.
- Test Firebase Emulator esercitano il blocco dei tentativi di impersonare UID o email di un collega, gli aggiornamenti ordinari consentiti e il cambio email autenticato. Un test statico impedisce di tornare a usare il campo `data.uid` come autorità.

**Limitazioni e rilascio:** queste correzioni sono nella sola PR. Prima di attivarle, verificare eventuali dati legacy incoerenti (UID del documento vs `data.uid` e email reale Auth vs profilo) e client che aggiornano il profilo dopo cambio indirizzo. Le nuove regole restano subordinate alla procedura protetta, al controllo GDPR e alla migrazione dei client; nessun dato account è stato modificato in produzione.
