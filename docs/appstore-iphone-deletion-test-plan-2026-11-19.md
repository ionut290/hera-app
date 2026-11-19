# VARGA CANTIERI — Piano collaudo account e video per Apple

**Stato:** piano operativo. **NON eseguire ancora la cancellazione su iPhone**: la funzione `deleteOwnVargaAccount` non è attiva in Firebase produzione. La beta interna 1.0 (10009) è stata caricata su TestFlight e dichiarata pronta da Apple. Il proprietario ha confermato su iPhone fisico il login Apple e la presenza del pulsante; NON ha eseguito né confermato una cancellazione reale.

## Sicurezza del collaudo

- Non utilizzare l'account amministratore o l'account di un dipendente.
- Preparare un **nuovo account fittizio** senza collegamenti a `personale`, clienti, ore, commesse e Drive aziendale.
- Verificare nell'interfaccia Firebase che il nuovo UID non sia associato a documenti o record reali. Non inviare nel ticket screenshot contenenti email personali, token o password.
- Non premere `Elimina il mio account` in TestFlight fino a quando la callable non è stata distribuita e convalidata con esito sicuro; un pulsante presente non implica che il backend sia disponibile.
- La versione web pubblica può differire dalla beta del ramo PR #1606. Il collaudo va eseguito sulla build esatta mostrata in TestFlight.

## Test automatici richiesti prima della distribuzione

1. GitHub PR #1606: i controlli obbligatori `critical-flows`, `auth-startup`, `e2e-smoke`, `Unsigned iOS simulator build` tutti verdi.
2. Test di cancellazione reale **solo negli emulatori** Firebase Auth + Firestore (progetto `demo-review-isolation`), con due identità fittizie; il secondo account non deve subire cambiamenti.
3. Verifica dei dati aziendali conservati: registri personale, ore, messaggi di lavoro, allegati condivisi, notifiche degli altri utenti.
4. Definire un processo di riconciliazione per gli eventuali ripristini dei backup Firestore: backup giornalieri, retention configurata 98 giorni; una copia storica non deve riattivare account cancellati.
5. Completare e approvare l'inventario dati personale per chat, log, documenti esterni e backup. Se non si può stabilire se un dato è aziendale o privato, non cancellarlo in modo indiscriminato.

## Test manuale futuro su iPhone (soltanto account dimostrativo)

1. Aprire TestFlight e verificare versione/build esatta.
2. Registrare lo schermo dall'iPhone con il comando **Registrazione schermo** del Centro di Controllo (attivarlo in Impostazioni se assente).
3. Aprire VARGA CANTIERI. Mostrare la registrazione/acquisizione dell'account dimostrativo e il login con Apple.
4. Aprire il profilo (l'utente ha già verificato il pulsante in questa sezione); visualizzare `Elimina il mio account`.
5. Solo dopo autorizzazione del collaudo isolato e disponibilità della callable: avviare la doppia conferma ed eliminare l'account fittizio.
6. Verificare che l'app ritorni allo stato non autenticato e che lo stesso UID non sia più presente in Firebase Authentication; controllare che un altro account fittizio e i dati di lavoro condivisi rimangano intatti.
7. Rivedere la registrazione; **tagliare o oscurare token, indirizzi email reali, dati dipendenti e notifiche private**. Caricare il video nel campo **App Review Information → Notes** della prossima presentazione, come richiesto da Apple.

## Stato reale verificato al 19 novembre 2026

- **Beta 1.0 (10009):** login con Apple confermato esplicitamente dal proprietario su iPhone fisico (nessun video ancora allegato).
- Pulsante «Elimina il mio account» confermato visibile dal proprietario nella schermata profilo. Il pulsante **non** prova da solo che la cancellazione sia disponibile.
- Beta iOS 1.0 (10009) in stato **Testing** su App Store Connect e confermata via email Apple «now available to test», installabile su TestFlight interno; [upload GitHub riuscito](https://github.com/ionut290/hera-app/actions/runs/44552567645).
- Audit **aggregato e non distruttivo** di cinque categorie Firestore completato con SUCCESS [GitHub #44567705049](https://github.com/ionut290/hera-app/actions/runs/44567705049). Non è un inventario GDPR esaustivo: log parzialmente campionati, scopi di conservazione ancora da classificare, code non finali e coordinate/record GPS che richiedono valutazione.
- Cloud Function `deleteOwnVargaAccount` **non distribuita in produzione**: il video di cancellazione completo non è ancora registrabile correttamente e l'utente non deve premere il pulsante sul proprio account.
- PR #1606 **draft** e non unita a `main`; verificato `mergeable=true` dopo l'allineamento dello script audit sul ramo PR. Non equivale a un'approvazione della pubblicazione. Evitare deploy Firebase automatici o manuali finché signoff privacy, account demo isolato e revisione dei dati aziendali non sono completi.
- I criteri di conservazione di massima sono stati approvati dal gestore dell'app, **non** è ancora stato provato il ruolo GDPR delle aziende titolari dei dati e non si è verificata la cancellazione su dati reali.
- Nessun video del flusso completo è stato inviato ad Apple e non è stata inviata una nuova richiesta di App Review dopo il rifiuto della build 21.

**Condizione di uscita:** nessun invio Apple che dichiari la cancellazione funzionante finché non sono completati test end-to-end sul demo, audit privacy e video su dispositivo fisico.
