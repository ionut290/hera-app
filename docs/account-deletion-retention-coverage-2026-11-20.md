# VARGA CANTIERI — inventario tecnico per cancellazione account e conservazione

**20 novembre 2026 — stato: NON APPROVATO PER RILASCIO GENERALE.**

Questo documento registra soltanto il comportamento osservabile nel codice di `functions/account-deletion.js`, nelle regole Firestore e nell'informativa pubblicata. **Non è una valutazione GDPR**, non dimostra di conoscere tutti i dati realmente conservati e non autorizza cancellazioni o accessi a dati di lavoratori.

## Cosa è già coperto nel codice (canary SOLO DEMO)

| Raccolta | Comportamento attuale |
| --- | --- |
| `documents` | Cancella documenti personali attribuiti all'UID autenticato e i file canonici non referenziati da documenti mantenuti; preserva i documenti condivisi/aziendali |
| `documentReminderAcknowledgements`, `alertReadConfirmations`, `userAlertAcknowledgements`, `userAlerts` | Elimina record riferiti all'utente mediante i campi di proprietà previsti dal codice |
| `passwordRecoveryChallenges` | Elimina sfide temporanee associate all'UID |
| `personale` | Scollega alcuni campi di autenticazione; conserva il registro del personale e le informazioni contrattuali |
| `notifications` | Rimuove riferimenti UID/stato personale da notifiche condivise; conserva il contenuto operativo |
| `chatMessages` | Scollega l'identità del mittente e sostituisce il destinatario UID con un valore non instradabile; non elimina messaggi aziendali |
| `appNotifications` | Scollega l'identità dell'autore mantenendo le comunicazioni |
| `notificationEmailQueue` | Minimizza la coda indirizzata esattamente all'UID/email verificata, senza richiamare email già recapitate |
| `activityLogs` | Minimizza i campi identificativi a corrispondenza esatta; conserva il dettaglio operativo |
| `privateDocuments`, `operatorPositions`, `userOperatorLinks`, `platformUsers` | Cancella il documento UID dedicato |
| Firebase Authentication | Cancella l'utente **per ultimo**, soltanto se tutti i passaggi precedenti sono riusciti |

Le raccolte `commesse`, `globalCommesse`, `oreReports` e `calendarEvents` non sono cancellate da questa funzione: possono includere dati personali e richiedono una valutazione distinta per finalità, ruolo delle aziende, conservazione e accesso.

## Termini già pubblicati nella privacy

L'informativa `privacy.html` descrive **termini programmati**, soggetti a verifiche e obblighi: profili/account fino a 12 mesi dopo la disattivazione (salvo motivo documentato), log tecnici 90 giorni, foto e allegati operativi 24 mesi dalla chiusura del lavoro, e certe registrazioni amministrativo-contabili 10 anni quando necessari e giustificati. Il semplice termine stampato nell'informativa non equivale all'attivazione di TTL o procedure automatiche.

Lo script protetto `scripts/audit-retention-readiness.cjs` esamina attualmente soltanto metadati di **cinque raccolte** (`activityLogs`, `notificationEmailQueue`, `gpsUpdateRequests`, `appNotifications`, `notifications`). Restano fuori da quell'audit gli allegati e le registrazioni aziendali. Le sue classificazioni sono **candidati a revisione manuale, non ordini di cancellazione**.

## Blocchi che devono restare aperti

- Verificare ruoli privacy e istruzioni documentate delle aziende titolari e delle altre parti interessate, senza equiparare l'accesso tecnico al consenso per la diffusione di dati personali
- Verificare con inventario separato gli allegati Storage/legacy, miniature, file condivisi, esportazioni Google Drive, copie di backup e conservazione su piattaforme esterne
- Riesaminare alias email legacy, testo libero di chat/documenti, dati personali in registrazioni lavoro, GPS e calendario, e casi di contenzioso/legal hold
- Definire procedure di rettifica, opposizione e conservazione residuale dopo la cancellazione di un account, con evidenze della decisione dell'azienda competente
- Verificare su iPhone fisico accesso Apple, revoca e cancellazione completa tramite un'identità fittizia, senza modificare account reali
- Rendere la funzione disponibile agli utenti ordinari solo dopo revisione, collaudo e distribuzione protetta; `DEMO_ONLY_RELEASE = true` resta obbligatorio allo stato attuale

## Verifica riproducibile senza Firebase

Eseguire `node scripts/check-account-retention-coverage.cjs`. Legge soltanto i file di progetto, verifica le categorie già classificate, confronta i termini pubblicati, impone il blocco demo e produce un rapporto `BLOCKED_PENDING_CONTROLLER_AND_RETENTION_REVIEW`. La suite `check-account-deletion-isolated.yml` esegue automaticamente il controllo sulle PR pertinenti.

**Non viene eseguito alcun inventario live o deploy.** Per l'inventario reale resta necessario il workflow manuale protetto `audit-retention-readonly-protected.yml` e la relativa approvazione specifica.
