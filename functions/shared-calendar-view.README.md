# Sincronizzazione calendario condiviso

Le funzioni `syncSharedCalendarFromOreReports`, `syncSharedCalendarFromOreApprovalRequests`, `syncSharedCalendarFromImpianti` e `syncSharedCalendarFromLavorazioni` aggiornano automaticamente `sharedStaticViews/calendario__YYYY-MM` dopo ogni creazione, modifica o eliminazione nelle raccolte delle ore e nelle attività delle commesse.

I documenti originali restano la fonte autorevole. La vista condivisa contiene soltanto il riepilogo mensile necessario alla consultazione: ore, attività FATTO, operatore e importi. Ogni aggiornamento è transazionale e conserva le altre sezioni del mese.

Il client apre al massimo un listener sul documento del mese visibile. Non esegue query per giorno, commessa, impianto o lavorazione.

## Distribuzione

```bash
firebase deploy --only functions:syncSharedCalendarFromOreReports,functions:syncSharedCalendarFromOreApprovalRequests,functions:syncSharedCalendarFromImpianti,functions:syncSharedCalendarFromLavorazioni
```
