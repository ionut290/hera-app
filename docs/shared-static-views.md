# Viste condivise e letture Firestore

All’avvio il client usa documenti di `sharedStaticViews` per personale, mezzi, squadre e ore:

- `registri__corrente` per personale e mezzi;
- `squadre__YYYY-MM-DD` per la composizione giornaliera;
- `calendario__YYYY-MM` per ore e richieste del calendario.

`shared-static-views.js` conserva una cache in memoria e locale e mantiene un solo listener Firestore per documento, condiviso tra i componenti. I log `[SHARED VIEWS]` indicano origine e quantità di record.

Le raccolte sorgenti restano autorevoli per le modifiche. Le Cloud Functions rigenerano le viste dopo scritture reali. Se una vista manca, il client mostra un errore senza tornare automaticamente a leggere l’intera raccolta sorgente.
