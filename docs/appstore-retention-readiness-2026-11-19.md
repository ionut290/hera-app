# Verifica retention e stato App Store — 19 novembre 2026

## Scopo

Questa pagina descrive le verifiche possibili **senza pubblicare** VARGA CANTIERI e separa gli automatismi già testati da quelli non ancora reali. Riguarda la PR [#1606](https://github.com/ionut290/hera-app/pull/1606), attualmente **draft**.

## Risultati già verificati

- Tutti i 21 workflow relativi al commit `650976037b86eb7169d7ff59f24640b8ced9b5f9` sono stati completati con esito **SUCCESS** il 19/11/2026. Dopo ulteriori commit bisogna verificare il nuovo HEAD, non riutilizzare questo conteggio.
- Il login Apple e il pulsante per eliminare l'account sono stati visualizzati su una beta iPhone, come documentato nel piano di collaudo.
- Test sintetici con Firebase Auth e Firestore Emulator hanno verificato l'isolamento di due utenti fittizi e la conservazione di dati condivisi.
- **Non** è stato eseguito un test distruttivo sul backend di produzione. La Cloud Function `deleteOwnVargaAccount` non è distribuita.
- Il tentativo di verifica diretta del 19/11/2026 in App Store Connect ha richiesto una nuova autenticazione. Non sono state inviate build o risposte.

## Audit conservazione: strumento aggiunto

`scripts/audit-retention-readiness.cjs` classifica **solo possibili casi di riesame**, senza decidere o eseguire alcuna cancellazione. In modalità predefinita rifiuta l'accesso online.

Eseguire in un checkout isolato:

```bash
node scripts/audit-retention-readiness.cjs --self-test
```

La prova usa dati inventati e verifica che:
- code email `pending` non siano considerate chiuse;
- richieste GPS senza data di risoluzione non ricevano una data di scadenza inventata;
- record con `legalHold` o `retentionException` non siano segnalati come candidati automatici;
- notifiche senza classificazione `retentionPurpose=transient` rimangano da classificare;
- i periodi indicativi non diventino mai comandi di cancellazione.

Il codice prevede una **modalità live solo in lettura**, disabilitata finché non vengono forniti credenziali isolate, autorizzazione esplicita e ambiente di esecuzione protetto. Non usare questa modalità in una console locale senza revisione. Se sarà eseguita, restituirà solo conteggi aggregati per categoria (senza ID, email, nomi, messaggi, allegati o coordinate). Il test sintetico **non prova** il rispetto della retention sui dati reali.

### Limiti dell'inventario

È un controllo preliminare, **non** un audit completo GDPR. I campi `resolvedAt`, `retentionPurpose`, `legalHold` e `retentionException` sono requisiti conservativi di classificazione, **non** campi che risultano uniformemente presenti nel database. Nessuna riga viene eliminata in automatico. I riferimenti di chiusura delle commesse, le conversazioni private, le foto, le copie Drive, gli export, i backup e gli allegati Storage richiedono verifiche e decisioni a parte.

## Bloccanti e ordine di rilascio

1. Verificare titolarità e accordi GDPR di aziende e dipendenti, classificazioni/retention ed eccezioni, e definire la gestione dei dati in Drive e nei backup.
2. Completare i test di compatibilità della nuova chat: distribuire indici, verificare `READY`, aggiornare i client e applicare solo dopo le regole più restrittive nel workflow protetto.
3. Predisporre un account Apple/Firebase **demo nuovo e isolato**, mai collegato a `personale`, clienti, ore, commesse o documenti reali.
4. Dopo signoff e revisione del codice, distribuire **solo** la callable attraverso `deploy-account-deletion-protected.yml`, con SHA esatto di `main` e approvazione dell'ambiente. Non avviare il deploy genericamente.
5. Verificare preflight, cancellazione completa su demo, revoca Apple e assenza del demo in Firebase Auth, preservando un altro utente sintetico.
6. Registrare su **iPhone fisico** il video richiesto da App Review (login demo, profilo, doppia conferma, esito), evitando dati aziendali/personali.
7. Accedere ad App Store Connect, aggiornare la build e inviare a nuova revisione **solo quando** test, privacy e video sono completati.

**Decisione corrente:** lasciare PR #1606 in bozza. Nessuna distribuzione della cancellazione su account reali o delle regole chat restrittive prima dei gate di sicurezza.
