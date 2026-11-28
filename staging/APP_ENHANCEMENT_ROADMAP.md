# Varga Cantieri 2.0 — audit e completamento funzionale

## Obiettivo

Mantenere Varga Cantieri 2.0 nella sua forma attuale e renderla affidabile e realmente utilizzabile. Il lavoro consiste nel verificare, correggere e completare le funzioni già presenti, rispettando interfaccia, flussi, dati e impostazioni esistenti.

Non aggiungere nuovi moduli o cambiare il modo di lavorare dell’app senza una richiesta esplicita.

## Regole di conservazione

- Lasciare invariati i flussi esistenti e la struttura generale dell’app.
- Proteggere FATTO, WhatsApp/WHAZZUP, impianti, Personale/Utenti, identificativi e storico.
- Non cancellare, azzerare o scollegare dati o relazioni esistenti.
- Mantenere compatibilità con i dati già presenti e con il funzionamento offline.
- Correggere una funzione solo dopo averne verificato dipendenze, permessi e dati coinvolti.
- Lavorare su staging/emulatore; nessuna pubblicazione di regole o modifiche in produzione durante l’audit.

## Verifiche funzionali

### Accesso e aziende

- Creazione account e azienda, accesso, uscita e riapertura dell’app.
- Collegamento a un’azienda tramite codice e approvazione/rifiuto dell’amministratore.
- Accesso Google e Apple nei percorsi previsti.
- Ruoli e permessi coerenti tra schermate, app e regole Firebase.
- Nessuna schermata vuota, blocco o ritorno improvviso al login.

### Piano prova e limiti

Verificare che il piano gratuito rispetti i limiti già definiti:

- 5 commesse;
- 100 impianti;
- 1 amministratore;
- 11 operatori;
- 11 mezzi.

Controllare che il limite venga applicato correttamente, che l’elemento successivo sia rifiutato con un messaggio comprensibile e che non si verifichino duplicati o conteggi errati.

### Commesse, impianti e completamento

- Creazione, apertura, modifica, ricerca, filtri e visualizzazione delle commesse.
- Inserimento manuale e gestione degli impianti, collegamenti e posizione in mappa.
- Passaggio Da fare/Fatti e conservazione di data, ora, operatore, note e storico.
- Pulsante FATTO e messaggio WhatsApp/WHAZZUP verificati senza alterarne la logica.
- Navigazione verso l’impianto e comportamento con rete assente/ripristinata.

### Personale, squadre e mezzi

- Inserimento e modifica manuale di persone, squadre e mezzi.
- Assegnazioni e collegamenti mantenuti nel tempo.
- Ruoli, permessi e storico coerenti.
- Nessuna perdita o duplicazione di record quando si sincronizza o si aggiorna l’elenco.

### Mappe e servizi

- Apertura mappa, posizione, ricerca e navigazione verso gli impianti.
- Servizio Neve: commesse, percorsi, importazione KML, stato dei tratti, modifica e reset del giro.
- Potature/Abbattimenti, Sfalcio COBO e QR alberi: verificare i flussi già presenti, senza sostituirli.
- Foto, allegati e documenti: acquisizione, visualizzazione e sincronizzazione.

### Contabilità, calendario e report

- Inserimento e consultazione dei dati contabili e delle ore.
- Calendario, assegnazioni e dati storici.
- Esportazioni e report: completezza dei dati, formattazione e collegamento alla commessa corretta.

### Offline e dispositivi

- Verificare la coda offline esistente e la sincronizzazione successiva.
- Ripetere le operazioni senza creare duplicati o sovrascrivere dati più recenti.
- Provare interfaccia, login, mappe, foto e pulsanti su iPhone e Android.
- Controllare che nessun comando sia coperto da notch, fotocamera o barra di sistema e che l’app resti utilizzabile con testo ingrandito.

## Ordine di lavoro

1. Inventariare schermate, pulsanti, percorsi utente e permessi già presenti.
2. Provare ogni percorso in staging con dati di test e annotare esito, errore e dipendenze.
3. Correggere prima blocchi di accesso, navigazione e operazioni fondamentali.
4. Correggere gli altri difetti in modifiche isolate, mantenendo i flussi e i dati attuali.
5. Ripetere i test di regressione per le parti protette e verificare i limiti del piano prova.
6. Eseguire prove sui dispositivi e preparare un riepilogo chiaro delle funzioni verificate e di quelle ancora bloccate.

## Criteri di completamento

- Ogni funzione già presente ha un percorso verificato dall’inizio alla fine oppure un difetto documentato con causa e correzione prevista.
- I limiti del piano prova corrispondono ai valori indicati sopra.
- FATTO, WhatsApp/WHAZZUP, impianti, Personale/Utenti, collegamenti e storico superano i test di regressione.
- Non si perdono dati durante uso offline, aggiornamento o sincronizzazione.
- Nessuna funzione è dichiarata operativa sulla sola base di un test dell’interfaccia: vengono verificati anche salvataggio, permessi e lettura dei dati.
- Nessuna modifica viene pubblicata in produzione senza un’autorizzazione e una verifica dedicate.
