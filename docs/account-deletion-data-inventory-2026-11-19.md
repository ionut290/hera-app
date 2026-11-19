# VARGA CANTIERI — Inventario dati per la cancellazione account
**Data verifica:** 19 novembre 2026  
**Progetto Firebase:** `hera-app-6cd2b`  
**Stato:** audit *parziale*, NON autorizza la distribuzione o la cancellazione di account reali.

## Metodologia e limiti
Il proprietario ha autorizzato verifiche in sola lettura sulla console Firebase. Sono stati verificati **il nome del bucket** `hera-app-6cd2b.firebasestorage.app` e **45 nomi di collezioni** Firestore nel database `(default)`. Non sono stati aperti documenti, osservati contenuti, identificativi utenti o dati personali. Non sono state ispezionate le sottocollezioni, i documenti in Google Drive, le esportazioni o i backup. La cartella Storage visibile `lavori-occasionali/` contiene potenzialmente file operativi che non sono soggetti alla cancellazione automatica attuale.

L'inventario è una **mappa di aree da verificare**, non una prova che le singole collezioni contengano informazioni personali, né una prova di cancellazione completa.

## Ambiti della procedura attualmente proposta in PR #1606
| Archivio o tipo dati | Azione proposta | Garanzie o riserve |
| --- | --- | --- |
| Firebase Authentication | Elimina solo l'utente autenticato, per ultimo | Autenticazione recente e conferma `ELIMINA`; l'azione è irreversibile |
| `platformUsers/{uid}` | Rimuove ricorsivamente | Solo UID del richiedente |
| `privateDocuments/{uid}`, `operatorPositions/{uid}`, `userOperatorLinks/{uid}` | Rimuove ricorsivamente | Solo chiave UID del richiedente |
| `documents` | Rimuove metadati di documenti personali associati al proprio UID | Include documenti legacy senza `visibility`; non elimina documenti condivisi/global |
| Firebase Storage `documents/{uid}/...` | Rimuove allegati personali noti | Percorsi convalidati; non rimuove file richiamati da documenti aziendali. Percorsi legacy ignoti fanno fallire il preflight |
| `documentReminderAcknowledgements`, `alertReadConfirmations`, `userAlertAcknowledgements`, `userAlerts`, `passwordRecoveryChallenges` | Rimuove conferme e sfide di recupero password per l’UID dell’utente | Le password-reset challenges vengono eliminate solo quando la colonna `uid` corrisponde al richiedente; la corrispondenza di altre collezioni dipende dallo schema |
| `personale` | Scollega `linkedUserId` e attributi di login derivati; conserva la scheda aziendale | Non cancella il rapporto di lavoro, la foto aziendale e i documenti amministrativi |
| `chatMessages` (messaggi di lavoro) | Per le righe con `senderId` uguale all'UID eliminato, elimina `senderId` e `senderEmail`, sostituisce `senderName` con `Account eliminato` | Conserva testo, metadati di commessa e identità di tutti gli altri utenti. Se `recipientId` corrisponde esattamente all'UID eliminato, lo sostituisce con un identificatore casuale `deleted-recipient-<UUID>` non riconducibile all'utente: **non lasciare il campo vuoto**, perché i client legacy potrebbero mostrare quel messaggio come chat generale. Le varianti legacy per email, i riferimenti nel contenuto e la retention di 12 mesi dalla chiusura richiedono ancora una valutazione separata |
| `appNotifications` | Per gli avvisi con `createdByUid` uguale all'UID elimina `createdByUid` e `createdByEmail`, sostituisce `createdByName` con `Account eliminato` | Conserva titolo, testo e altri destinatari; la retention programmata di 90 giorni non è ancora automatizzata |
| `notifications` (notifiche centrali) | Rimuove l’`actorId`, anonimizza `actorName` se corrispondono all’UID cancellato, elimina l’UID da `recipientUserIds`, `readBy`, `deletedBy` | Conserva il messaggio di lavoro e gli altri destinatari; il testo e `metadata` richiedono una distinta valutazione privacy |

## Aree osservate da esaminare prima del rilascio
| Area e collezioni osservate | Rischio da verificare | Stato |
| --- | --- | --- |
| Messaggi e notifiche non interamente coperte: destinatari e testi di `chatMessages`, corpi e retention di `appNotifications`, `notificationEmailQueue`, `hoursDeadlineAlerts`, testi e metadata di `notifications` | UIDs, email, testo dei messaggi, liste destinatari: definire cancellazione, anonimizzazione o conservazione motivata | **Non inventariato a livello di campi** |
| Ore e registri operativi: `oreReports`, `oreApprovalRequests`, `oreLocks`, `ferieColleghi`, `operatorCommessaHistory` | Rapporti lavorativi ed eventuali obblighi di conservazione; evitare la perdita delle ore degli altri dipendenti | **Retention da approvare** |
| Log, sicurezza e autorizzazioni: `activityLogs`, `auditLogsWhazzup`, `operatorUserLinkHistory`, `securityPrivate`, `passwordRecoveryAttempts`, `passwordRecoveryChallenges`, `appErrorGroups` | Dati di sicurezza e storico di accesso; definire minimizzazione e tempi di conservazione | **Retention da approvare** |
| Geolocalizzazione e impianti: `gpsUpdateRequests`, `impianti`, `impiantoSafetyChecks`, `operatorPositions` | Identificativi operatore, posizione e cronologia: verificare se ci sono dati personali persistenti | **Non inventariato a livello di campi** |
| Risorse e documenti aziendali: `commesse`, `globalCommesse`, `commessaResources`, `posDocuments`, `pianoNeveUsers` | Contenuti o riferimenti personali dentro dati aziendali condivisi | **Retention da approvare** |
| Backup/snapshot: `registryImportBackups`, `gestionaleSyncChangeIndex`, `gestionaleSyncSnapshots`, `gestionaleSyncState` | Ricomparsa di PII dopo il ripristino; tempi di cancellazione e politiche di backup | **Non inventariato a livello di campi** |
| Google Drive, esportazioni, allegati legacy Storage `lavori-occasionali/` | Copie esterne di informazioni sull'utente che la funzione Firebase non tocca | **Fuori copertura attuale** |

Le altre collezioni osservate comprendono `appConfig`, `companies`, `companyCodes`, `contabilita`, `mezzi`, `neve_commesse`, `neve_squadre`, `neve_squadre_storico`, `pianoNeveStreetWork`, `safetyContacts`, `servizioNeveClienti`. La loro presenza non implica automaticamente dati personali e richiede classificazione a livello di campi.

## Audit di schema in sola lettura — 19 novembre 2026

Verificati nella console Firestore **solamente i nomi e i tipi dei campi** di **un documento campione** per ciascuna delle sei collezioni seguenti. **Nessun valore di email, UID, testo, nome, coordinate o ID di documento è stato trascritto.** Firestore è schemaless: un campione non dimostra che ogni documento abbia gli stessi campi. Questa verifica NON costituisce approvazione della retention né autorizzazione a cancellare i dati.

| Collezione | Campi personali o rilevanti osservati (solo schema) | Classificazione prudenziale e prossima verifica |
| --- | --- | --- |
| `chatMessages` | `senderId`, `recipientId`, `senderEmail`, `senderName`, `text`, `metadata`, `createdAt` | Comunicazioni anche aziendali: possibile PII dell'autore e del destinatario. Valutare rimozione dei collegamenti di accesso, anonimizzazione mirata e conservazione di conversazioni operative; NON cancellare interi thread senza approvazione |
| `appNotifications` | `createdByUid`, `createdByEmail`, `createdByName`, `body`, `title`, `commessaId` | Dati del creatore associati a notifiche di lavoro. Audit destinatari/permessi e necessità di retention; scollegare identità solo se non pregiudica le comunicazioni operative |
| `notificationEmailQueue` | `to`, `createdBy`, `notificationId`, `html`, `text`, `status`, `createdAt` | Destinatario e corpo email potrebbero contenere PII; verificare code pending, gestione post-cancellazione, contenuto e durata dei messaggi senza interrompere invii aziendali legittimi |
| `activityLogs` | `userId`, `userEmail`, `userName`, `deviceInfo`, `detail`, `technicalError`, `createdAt` | Log tecnici con identificativi e contenuti diagnostici; informativa indica **90 giorni programmati**, da verificare se realmente applicati e quali eventi abbiano bisogno di retention maggiore |
| `gpsUpdateRequests` | `operatorId`, `operatorEmail`, `operatorName`, `operatorLat`, `operatorLng`, `approvedBy`, `createdAt` | Dati localizzativi e identificativi, potenzialmente particolarmente delicati per il lavoro. Non diffonderli, non cancellare i record operativi senza una politica approvata |
| `passwordRecoveryAttempts` | `failCount`, `lastAttemptAt`, `lastResult`, `lockedUntil`, `windowStartedAt` | Nessuna email o UID fra i **campi del campione**; verificare però se l'**ID del documento** identifica un utente (ID non letto) e definire cancellazione/TTL |

**Regola per la release:** la callable attuale non include la pulizia automatica di tutti questi archivi. Non dichiarare che l'eliminazione rimuove chat, storico GPS, log o code email senza prove; conservare i registri con motivazione e retention approvate dalle aziende titolari. Preferire anonimizzazioni selettive e verificabili. Le scelte giuridiche e contrattuali non devono essere dedotte dai nomi dei campi.

**Backup:** Google Cloud Console indica backup Firestore giornalieri con retention configurata di **98 giorni**, PITR disattivato. La cancellazione attiva non elimina retroattivamente le copie; un ripristino deve riconciliare gli account cancellati e impedire il ripristino di accessi revocati.

## Controllo separato del rilascio in produzione

La funzione `deleteOwnVargaAccount` è stata **esclusa** dal workflow generale `.github/workflows/deploy-firebase-functions.yml`. Un workflow dedicato `.github/workflows/deploy-account-deletion-protected.yml` impone una richiesta manuale esplicita, SHA esatto su `main`, dichiarazione di audit privacy approvato, account demo predisposto e ambiente GitHub `firebase-account-deletion-production` già esistente con **required reviewers**. Se manca l'ambiente protetto, il workflow deve fermarsi prima di accedere alle credenziali. **Questa preparazione non equivale a un deploy:** la funzione resta disattivata in produzione finché le condizioni non sono soddisfatte.

**Approvazione dei criteri:** il referente ha approvato il 19 novembre 2026 i periodi dettagliati in `docs/privacy-retention-proposal-2026-11-19.md`. Le nuove anonimizzazioni per `chatMessages` e `appNotifications` sono proposte di codice nel ramo della PR #1606, non sono distribuite. Le scadenze di conservazione, le eccezioni legali e i controlli sui fornitori esterni richiedono implementazione separata prima di dichiarare che il sistema le applica automaticamente.

**Verifica aggiuntiva della logica chat e notifiche:** una rimozione completa del campo `recipientId` avrebbe potuto essere interpretata dai vecchi client come una chat generale. La PR usa invece un tombstone casuale non associato a un account reale; i test conservano il corpo dei messaggi e verificano che i campi dell'altro operatore non cambino. Per `notifications` vengono individuati i flag `readBy` / `deletedBy` impostati sia a `true` sia a `false`. Questi comportamenti sono soltanto proposti su ramo fino alla validazione CI e alla distribuzione deliberata.

## Criteri di accettazione per il rilascio
1. Confermare struttura e schema dei record per gli archivi rilevanti con letture limitate e autorizzate, evitando di trascrivere o esporre informazioni personali nei log.
2. Definire e documentare quali dati aziendali devono essere conservati e per quanto tempo; predisporre minimizzazione/scollegamento degli identificativi personali quando possibile.
3. Definire la gestione di copie nei backup, dei dati di Google Drive e dei file Storage legacy, includendo assistenza privacy quando la rimozione non è completamente automatizzabile.
4. Eseguire il **preflight e la cancellazione** solamente su un account Firebase creato appositamente per il collaudo, senza collegamenti a personale reale e senza commesse reali; verificare anche che il resto dell'azienda rimanga intatto.
5. Testare su iPhone/iPad reale accedi con Apple, creazione account, eventuale stato pending, eliminazione e revoca Apple; registrare il video completo richiesto da Apple.
6. Confermare build firmata successiva alla 1.0 (21) e aggiornamento Apple App Review, senza dichiarare già pubblicate funzioni ancora in bozza.

**Decisione di rilascio:** NON distribuire la callable `deleteOwnVargaAccount` fino alla verifica dei punti 1–4. I test GitHub sul codice simulato non sostituiscono una prova end-to-end.
