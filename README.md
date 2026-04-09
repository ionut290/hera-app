# Hera App (HTML/CSS/JS + Firebase)

Web app senza framework con:

- Login Google via Firebase Authentication
- CRUD base su Firestore (`impianti`)
- Importazione Excel (`.xlsx/.xls`) con SheetJS
- UI mobile-first
- Config pronta per Firebase Hosting

## Avvio locale

Apri `index.html` con un server statico (consigliato):

```bash
python3 -m http.server 8080
```

Poi visita `http://localhost:8080`.

## Deploy Firebase Hosting

1. Installa Firebase CLI.
2. Esegui login:

```bash
firebase login
```

3. Deploy:

```bash
firebase deploy --only hosting
```

## Note Firestore

La collezione usata è `impianti` con ordinamento per `createdAt` in `app.js`.
