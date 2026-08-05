# Issue #913 — Ottimizzazione letture oreReports

## Problema osservato

Nel report `diagnostica-firestore-2026-08-05(13).json`, la Home mantiene un listener globale su `oreReports` tramite `subscribeHoursStats()`. Con circa 282 documenti presenti, una singola sessione ha prodotto tre consegne complete per un totale di 846 documenti letti dal listener.

## Obiettivo

Eliminare il listener globale su `oreReports` dalla Home senza modificare il comportamento visibile dell'app.

## Vincoli di sicurezza

- Non cancellare o migrare documenti esistenti.
- Non modificare le regole Firestore.
- Non alterare inserimento, modifica o eliminazione ore.
- Non alterare calendario personale o condiviso.
- Non alterare squadre, pulsante `+ORE`, riepiloghi e dashboard commesse.
- Le ore delle commesse disattivate devono restare nello storico.
- Mantenere sincronizzazione multi-dispositivo.

## Strategia

1. La Home usa `sharedStaticViews/calendario__YYYY-MM` come fonte primaria del mese corrente.
2. `oreReports` non viene più ascoltata senza filtri.
3. La pagina Ore effettua una sola query per il mese selezionato.
4. Le richieste mensili vengono deduplicate con una cache di Promise per chiave `YYYY-MM`.
5. Dopo una mutazione si aggiorna subito il singolo record locale e si invalida solo il mese interessato.
6. Se la vista condivisa non è disponibile, eseguire una singola query limitata all'intervallo del mese corrente.
7. `oreApprovalRequests` resta separata, ma deve essere limitata ai record rilevanti.

## Criteri di accettazione diagnostici

- Home: `oreReports:listener-delivery = 0`.
- Modifica di una singola ora: nessuna riconsegna dell'intera collection.
- Home: una lettura della vista condivisa mensile.
- Pagina Ore: una sola query del mese selezionato.
- Nessuna regressione su calendario, squadre, `+ORE`, export, commesse e offline/cache.

## Sequenza test

1. Aprire l'app e verificare riepilogo di oggi.
2. Aprire una commessa e controllare totali ore.
3. Inserire un'ora e verificare aggiornamento immediato.
4. Chiudere e riaprire l'app e verificare persistenza.
5. Modificare ed eliminare la stessa ora.
6. Cambiare mese nel calendario.
7. Verificare una commessa disattivata nello storico.
8. Ripetere come admin e utente normale.
9. Generare un nuovo report diagnostico e confrontarlo con quello del 5 agosto 2026.
