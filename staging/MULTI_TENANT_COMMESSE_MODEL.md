# Varga Cantieri 2.0 — modello multiaziendale delle commesse

## Regola funzionale
Sfalci COBO, Potature, Hera e Inrete sono esempi di COMMESSE, non moduli applicativi separati.

Ogni azienda:
- crea, rinomina e gestisce le proprie commesse secondo i propri permessi;
- inserisce i propri cantieri dentro ciascuna commessa;
- gestisce impianti, indirizzi, coordinate, foto, note, orari e stati dei lavori dei propri cantieri;
- assegna i propri dipendenti e permessi;
- parte da un archivio vuoto quando viene creata, senza commesse o cantieri pre-caricati automaticamente.

## Struttura dati
tenants/{tenantId}/commesse/{commessaId}/impianti/{impiantoId}

Tutte le letture, scritture, ricerche, allegati e azioni devono restare nel tenant autenticato. Nessun dato di Avola Coop deve essere copiato in un nuovo tenant.

## Interfaccia
Una sola interfaccia condivisa e identica per tutte le aziende. Non creare schermate o applicazioni differenti per singole aziende o singole commesse. Le funzionalità specialistiche già esistenti (es. Sfalci COBO) si attivano in base alla commessa e alla configurazione, mantenendo il flusso originale.

## Permessi e protezioni
La creazione delle commesse è riservata ai ruoli aziendali autorizzati; l'inserimento dei cantieri segue i permessi definiti dall'azienda. La logica protetta del pulsante FATTO resta centralizzata e non modificabile dai singoli utenti aziendali. I dati e le azioni di un tenant non devono essere accessibili da un altro.
