# Verifica degli eventi disponibili in `activityLogs`

La view **Attività utente** legge esclusivamente i documenti già presenti nella collection `activityLogs` e non aggiunge nuovi tracciamenti.

## Eventi attualmente registrati dal codice

- accesso effettuato (`login_app`);
- apertura dell'app (`apertura_app`);
- pressione di NAVIGA (`pressione_naviga`);
- pressione di FATTO (`pressione_fatto`);
- pressione di FORZA (`pressione_forza`);
- invio WhatsApp (`invio_whatsapp`);
- apertura mappa (`apertura_mappa`);
- errori Firestore intercettati (`errore_firestore`).

I documenti possono contenere utente, ruolo, tipo e descrizione dell'azione, commessa, impianto, dettaglio, view, pulsante, dati tecnici dell'errore, dispositivo e data. La view nasconde ogni campo assente. Un ID SAP viene mostrato solo se esiste già in uno dei campi `impiantoSap`, `idSap` o `idSAP` del documento.

## Eventi richiesti ma non tracciati esplicitamente

Nel codice attuale non risultano chiamate dedicate per chiusura dell'app, apertura di una commessa, apertura della scheda di un impianto, inserimento ore, creazione o modifica di una segnalazione e sincronizzazione dati. Come richiesto, questa modifica non introduce automaticamente tali eventi e non cambia il salvataggio delle attività o la logica del pulsante FATTO.
