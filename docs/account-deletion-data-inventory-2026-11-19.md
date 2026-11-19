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
| `documentReminderAcknowledgements`, `alertReadConfirmations`, `userAlertAcknowledgements`, `userAlerts` | Rimuove record associati con campi UID specifici | La corrispondenza dipende dai campi previsti nel codice |
| `personale` | Scollega `linkedUserId` e attributi di login derivati; conserva la scheda aziendale | Non cancella il rapporto di lavoro, la foto aziendale e i documenti amministrativi |

## Aree osservate da esaminare prima del rilascio
| Area e collezioni osservate | Rischio da verificare | Stato |
| --- | --- | --- |
| Messaggi e notifiche: `chatMessages`, `notifications`, `appNotifications`, `notificationEmailQueue`, `hoursDeadlineAlerts` | UIDs, email, testo dei messaggi, liste destinatari: definire cancellazione, anonimizzazione o conservazione motivata | **Non inventariato a livello di campi** |
| Ore e registri operativi: `oreReports`, `oreApprovalRequests`, `oreLocks`, `ferieColleghi`, `operatorCommessaHistory` | Rapporti lavorativi ed eventuali obblighi di conservazione; evitare la perdita delle ore degli altri dipendenti | **Retention da approvare** |
| Log, sicurezza e autorizzazioni: `activityLogs`, `auditLogsWhazzup`, `operatorUserLinkHistory`, `securityPrivate`, `passwordRecoveryAttempts`, `passwordRecoveryChallenges`, `appErrorGroups` | Dati di sicurezza e storico di accesso; definire minimizzazione e tempi di conservazione | **Retention da approvare** |
| Geolocalizzazione e impianti: `gpsUpdateRequests`, `impianti`, `impiantoSafetyChecks`, `operatorPositions` | Identificativi operatore, posizione e cronologia: verificare se ci sono dati personali persistenti | **Non inventariato a livello di campi** |
| Risorse e documenti aziendali: `commesse`, `globalCommesse`, `commessaResources`, `posDocuments`, `pianoNeveUsers` | Contenuti o riferimenti personali dentro dati aziendali condivisi | **Retention da approvare** |
| Backup/snapshot: `registryImportBackups`, `gestionaleSyncChangeIndex`, `gestionaleSyncSnapshots`, `gestionaleSyncState` | Ricomparsa di PII dopo il ripristino; tempi di cancellazione e politiche di backup | **Non inventariato a livello di campi** |
| Google Drive, esportazioni, allegati legacy Storage `lavori-occasionali/` | Copie esterne di informazioni sull'utente che la funzione Firebase non tocca | **Fuori copertura attuale** |

Le altre collezioni osservate comprendono `appConfig`, `companies`, `companyCodes`, `contabilita`, `mezzi`, `neve_commesse`, `neve_squadre`, `neve_squadre_storico`, `pianoNeveStreetWork`, `safetyContacts`, `servizioNeveClienti`. La loro presenza non implica automaticamente dati personali e richiede classificazione a livello di campi.

## Criteri di accettazione per il rilascio
1. Confermare struttura e schema dei record per gli archivi rilevanti con letture limitate e autorizzate, evitando di trascrivere o esporre informazioni personali nei log.
2. Definire e documentare quali dati aziendali devono essere conservati e per quanto tempo; predisporre minimizzazione/scollegamento degli identificativi personali quando possibile.
3. Definire la gestione di copie nei backup, dei dati di Google Drive e dei file Storage legacy, includendo assistenza privacy quando la rimozione non è completamente automatizzabile.
4. Eseguire il **preflight e la cancellazione** solamente su un account Firebase creato appositamente per il collaudo, senza collegamenti a personale reale e senza commesse reali; verificare anche che il resto dell'azienda rimanga intatto.
5. Testare su iPhone/iPad reale accedi con Apple, creazione account, eventuale stato pending, eliminazione e revoca Apple; registrare il video completo richiesto da Apple.
6. Confermare build firmata successiva alla 1.0 (21) e aggiornamento Apple App Review, senza dichiarare già pubblicate funzioni ancora in bozza.

**Decisione di rilascio:** NON distribuire la callable `deleteOwnVargaAccount` fino alla verifica dei punti 1–4. I test GitHub sul codice simulato non sostituiscono una prova end-to-end.
