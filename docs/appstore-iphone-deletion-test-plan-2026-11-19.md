# VARGA CANTIERI — Piano collaudo account e video per Apple

**Stato:** piano operativo. **NON eseguire ancora la cancellazione su iPhone**: la funzione `deleteOwnVargaAccount` non è attiva in Firebase produzione. La beta 1.0 (10003) su TestFlight permette per ora di verificare il login Apple e la presenza del pulsante.

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

- Login con Apple confermato dall'utente su iPhone.
- Pulsante di eliminazione visibile nella schermata profilo dell'iPhone, come da screenshot dell'utente.
- Beta iOS 1.0 (10003) installabile tramite TestFlight interno.
- Cloud Function `deleteOwnVargaAccount` **non distribuita in produzione**, pertanto il video completo della cancellazione non è ancora registrabile correttamente.
- PR #1606 **draft** e non unita a `main`; evitare di attivare il workflow deploy Firebase che potrebbe distribuire funzioni non ancora approvate.

**Condizione di uscita:** nessun invio Apple che dichiari la cancellazione funzionante finché non sono completati test end-to-end sul demo, audit privacy e video su dispositivo fisico.
