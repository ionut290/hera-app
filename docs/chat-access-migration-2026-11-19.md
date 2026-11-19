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
