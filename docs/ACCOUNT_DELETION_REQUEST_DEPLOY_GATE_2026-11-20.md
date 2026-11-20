# VARGA CANTIERI — rilascio protetto della richiesta di cancellazione (preparato)

**Stato: SOLO WORKFLOW PREPARATO. Nessuna funzione nuova distribuita, nessun flag client attivato.**

La console Firebase, verificata il 20 novembre 2026 in sola lettura, mostra `deleteOwnVargaAccount` fra le funzioni deployate ma NON `requestOwnVargaAccountDeletion` né `getOwnVargaAccountDeletionStatus`.

## Dipendenze obbligatorie prima di eseguire il workflow manuale

1. Ottenere approvazione documentata del titolare/responsabile competente su categorie di dati, eventuali retention legali, ruoli, backup, verifiche GDPR, approvazioni dei terzi interessati e comunicazioni agli utenti.
2. Identificare una procedura **effettivamente presidiata** per istruire, completare e confermare le richieste di cancellazione, con scadenza realistica verificata; non dichiarare 30 giorni se il processo non lo può mantenere. Fare prima prove con un account sintetico e un'altra identità isolata.
3. Configurare e confermare il GitHub Environment `firebase-account-deletion-production` con **required reviewers**. Richiedere che un revisore competente approvi il job. Non usare un ambiente non protetto.
4. Solo dopo gli atti precedenti, sull'attuale `main` usare `deploy-account-deletion-request-protected.yml` con SHA esatto, attestazioni e riferimento approvazione (senza PII). Il workflow distribuisce esclusivamente due callable non distruttive: richiesta e stato per l'utente autenticato.
5. Verificare in produzione tramite account sintetici la ricezione, isolamento per UID, visualizzazione dello stato, gestione e conferma completamento. **La funzione NON cancella di per sé account o dati**, non riempie automaticamente il processo amministrativo.
6. Attivare la UI mediante `__VARGA_ACCOUNT_DELETION_REQUEST_ENABLED === true` **e** `__VARGA_ACCOUNT_DELETION_REQUEST_SLA_DAYS` (intero 1–90, tempistica approvata dal responsabile reale) in una modifica separata e rivista, con un nuovo rilascio iOS. Senza tempistica approvata il pulsante resta nascosto anche con il flag attivo. Allineare anche la pagina pubblica `delete-account.html` alla SLA effettiva, verificare il deploy Hosting separato con controllo privacy e notificare il completamento effettivo.
7. Verificare su iPhone vero login Apple, revoca dei grant Apple, cancellazione e pieno accesso reviewer, poi procedere ad Apple solo se l'audit complessivo è davvero pronto.

## Garanzie tecniche

- Workflow manuale, `DO_NOT_DEPLOY` predefinito, solo `main`, SHA esatto, due attestazioni e riferimento documentale.
- Primo job senza segreti di produzione, verifica presenza di reviewer richiesti nell'environment GitHub.
- Secondo job dietro environment protetto; Firebase service account validato per `hera-app-6cd2b`.
- Esporta temporaneamente le due callable SOLO nel checkout del runner; `functions/main.js` su GitHub resta invariato e il normale deployment le esclude.
- Non esegue `deleteOwnVargaAccount`, non cambia regole Firestore, non attiva pulsanti, non crea account revisore, non firma o carica build Apple.
- Test `scripts/check-account-deletion-request-release-guard.cjs` verifica i limiti di distribuzione in modo offline.

**Non trattare le attestazioni testuali come prova legale automatica.** Lo stato rimane bloccato finché signoff e procedura reale non esistono davvero.
