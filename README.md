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

- Il Drive usato è centralizzato: **ionut29019@gmail.com** collega il proprio account una volta, poi tutti gli utenti usano quel bridge.
- Solo l'utente admin (`ionut29019@gmail.com`) vede attivo il pulsante **Collega Google Drive**.
- L'app crea (se non esistono) le cartelle:
  - `Hera App - Dati`
  - `Hera App - Dati/Chat Media`
  - `Hera App - Dati/Report Impianti`
- Token/folder del bridge Drive vengono salvati in `appConfig/driveBridge` su Firestore per essere riusati da tutti gli utenti autenticati.
- I media chat vengono caricati su Drive e salvati in Firestore come URL.
- Per ogni commessa viene usato un solo Google Sheet (`Commessa - <nome commessa>`) dentro `Report Impianti`.
- In creazione commessa puoi impostare opzionalmente un **ID/link Google Sheet** già esistente: da quel momento il pulsante **Fatto** appende sempre lì (senza creare un nuovo file).
- Quando un impianto viene segnato come **Fatto**, viene aggiunta una nuova riga nel foglio della commessa con i dati dell'impianto, data/ora esecuzione e operatore.

## Checklist manutenzione

- Quando si aggiunge/modifica una feature, aggiornare `appConfig/helpCenter` (oppure `appHelpFaq`) con domanda/risposta/passi e pubblicare un nuovo snapshot Drive.

## Passo-passo: mantenere web app intatta + pubblicare su Play Store

Questa procedura mantiene la versione web invariata: **la web app resta la sorgente principale**, Android è solo un contenitore.

1. Verifica PWA base (già predisposta in questo repo): `manifest.webmanifest`, icona SVG e `sw.js`.
2. Continua a distribuire la web app su Firebase Hosting come sempre.
3. Usa Capacitor (config già pronta in `capacitor.config.ts`) per creare il wrapper Android in `android/`.
4. Genera il file `.aab` da Android Studio e caricalo in Play Console.

### Comandi consigliati (quando vuoi attivare Android)

```bash
npm install
npm run android:add
npm run android:sync
npm run android:open
```

`capacitor.config.ts` è già incluso nel repository con `appId` `it.vargacantieri.hera`, quindi non devi inizializzarlo manualmente.

In Android Studio:
- Build > Generate Signed Bundle / APK
- Seleziona **Android App Bundle (AAB)**
- Firma e pubblica su Play Console

### Nota importante

Per mantenere la web app intatta, evita refactor lato UI/rotte solo per Android: usa plugin Capacitor solo se servono feature native (camera, notifiche, file).
