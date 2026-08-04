# Sincronizzazione calendario condiviso

Le funzioni `syncSharedCalendarFromOreReports` e `syncSharedCalendarFromOreApprovalRequests` aggiornano automaticamente `sharedStaticViews/calendario__YYYY-MM` dopo ogni creazione, modifica o eliminazione nelle raccolte delle ore.

Il documento originale delle ore resta la fonte autorevole. La vista condivisa contiene solo un riepilogo mensile destinato alla consultazione e viene aggiornata tramite transazione server-side.

## Distribuzione

```bash
firebase deploy --only functions:syncSharedCalendarFromOreReports,functions:syncSharedCalendarFromOreApprovalRequests
```
