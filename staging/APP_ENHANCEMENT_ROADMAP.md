# Varga Cantieri 2.0 — ampliamenti funzionali

## Obiettivo

Integrare le idee più utili delle soluzioni di gestione del verde e delle squadre sul territorio, mantenendo l’interfaccia, i dati e i flussi già presenti in Varga Cantieri 2.0. Le nuove funzioni si aggiungono come moduli compatibili: non sostituiscono le procedure esistenti.

## Funzioni già presenti da riutilizzare

- Flussi specializzati per Potature/Abbattimenti e Sfalcio COBO.
- Pianificatore cartografico del Servizio Neve con percorsi, stati, import/export KML e coda offline specifica.
- Lettura QR per le schede degli alberi.
- Commesse, impianti, mappe, personale, mezzi, calendario, contabilità, foto e documenti.
- Strumenti di durata dati e sincronizzazione da verificare prima di estendere ad altri moduli.

Queste funzioni restano intatte e vanno usate come base quando hanno già la logica necessaria.

## Nuove funzioni selezionate

### 1. Ordini di lavoro generali — prima priorità

Un modulo ODL disponibile per le commesse ordinarie, senza rimpiazzare i moduli specializzati esistenti.

Ogni ODL dovrà poter collegare una commessa e, quando pertinente, uno o più impianti; indicare priorità, scadenza, squadra, operatori, mezzi, istruzioni e stato; raccogliere orari, materiali, note e foto prima/dopo; produrre un rapporto esportabile.

Gli stati e i permessi dovranno impedire chiusure o modifiche non autorizzate. Il completamento dovrà restare compatibile con i flussi originali FATTO/WHAZZUP e non aggirarli.

### 2. QR universale per impianti e mezzi

Estendere la scansione già disponibile per gli alberi affinché il codice di un impianto o mezzo apra la scheda corretta e, se consentito, l’ODL collegato. Ogni codice dovrà essere convalidato nel tenant selezionato: la scansione di un codice di un’altra azienda non deve esporre dati.

La compatibilità NFC va verificata separatamente su iOS e Android prima di includerla nel requisito.

### 3. Pianificazione ricorrente e capacità squadre

Aggiungere piani ripetuti per attività stagionali e controlli periodici, con assegnazione di squadra e mezzo, rilevamento delle sovrapposizioni e riprogrammazione suggerita quando il meteo o un imprevisto cambia il programma. Le modifiche dovranno essere confermate da una persona.

### 4. Segnalazioni trasformabili in ODL

Consentire a un utente autorizzato dell’azienda di inviare una segnalazione con posizione, impianto, foto e descrizione; l’amministratore o responsabile potrà convertirla in ODL senza reinserire i dati. Un eventuale portale pubblico per cittadini o committenti resta una fase successiva, dopo la definizione di accessi e privacy.

### 5. Lavoro offline e sincronizzazione uniforme

Verificare la coda offline esistente e riutilizzarla per ODL, mappe e prove fotografiche. La sincronizzazione dovrà essere idempotente, legata a utente e tenant, mostrare gli errori e presentare i conflitti prima di sovrascrivere dati.

## Ordine di lavoro

1. Chiudere l’audit delle regole Firebase staging attive e completare l’inventario tenant dei moduli già esistenti.
2. Progettare e collaudare ODL generale solo nello staging/emulatore, con regole server-side e test d’isolamento.
3. Estendere il QR a impianti e mezzi riutilizzando lo scanner esistente.
4. Aggiungere pianificazione ricorrente e gestione delle segnalazioni.
5. Estendere e testare offline/sincronizzazione su tutti i flussi approvati.
6. Provare su iPhone e Android con account e dati di test prima di qualsiasi attivazione operativa.

## Criteri di accettazione

- Tutti i nuovi documenti e allegati sono nel tenant corretto; test di lettura e scrittura incrociata tra tenant sempre negati.
- Ruoli e transizioni di stato sono controllati dalle regole Firestore, non solo dalla UI.
- Le prove coprono creazione, assegnazione, aggiornamento, rifiuto di modifiche non autorizzate, sincronizzazione ripetuta e conflitti offline.
- Le regressioni esistenti su commesse, impianti, Personale/Utenti, neve, FATTO e WHAZZUP restano verdi.
- Nessuna regola live viene pubblicata e nessuna funzione viene attivata in produzione prima dell’audit autorevole e delle prove su dispositivo.
