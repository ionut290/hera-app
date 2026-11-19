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
