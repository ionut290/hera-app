# Hera App (HTML/CSS/JS + Firebase)

Web app senza framework con:

- Login Google via Firebase Authentication
- CRUD base su Firestore (`impianti`)
- Importazione Excel (`.xlsx/.xls`) con SheetJS
- Collegamento Google Drive per:
  - salvataggio media chat (foto/video/audio) su Drive con link condivisibile
  - creazione automatica di un Google Sheet quando si preme **Fatto** su un impianto
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

## Note Google Drive / Google Sheets

- Dopo il login Firebase, usare il bottone **Collega Google Drive**.
- L'app crea (se non esistono) le cartelle:
  - `Hera App - Dati`
  - `Hera App - Dati/Chat Media`
  - `Hera App - Dati/Report Impianti`
- I media chat vengono caricati su Drive e salvati in Firestore come URL.
- Quando un impianto viene segnato come **Fatto**, viene creato un Google Sheet con i dati dell'impianto, data/ora esecuzione e operatore.
