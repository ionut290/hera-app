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

## Notifiche PWA e attività in background

- Nell'area utente è presente il blocco **Notifiche** con:
  - pulsante **Attiva notifiche** (richiede permesso browser),
  - pulsante **Test notifica** (invia una notifica locale via Service Worker).
- `sw.js` ora gestisce:
  - evento `push` (notifiche remote),
  - evento `notificationclick` (apertura/focus app),
  - evento `sync` con tag `hera-app-background-check` (test background sync).
- Per attivare le push remote reali configura una chiave VAPID pubblica in **uno** di questi modi:
  1. variabile globale prima di caricare `app.js`:

```js
window.HERA_PUSH_PUBLIC_VAPID_KEY = "INSERISCI_LA_TUA_CHIAVE";
```

  2. meta tag nell`<head>`:

```html
<meta name="hera-push-vapid-key" content="INSERISCI_LA_TUA_CHIAVE" />
```

  3. localStorage (utile per test rapido):

```js
localStorage.setItem("heraPushPublicVapidKey", "INSERISCI_LA_TUA_CHIAVE");
```

Senza chiave VAPID l'app continua a funzionare normalmente: avrai notifiche locali di test, ma non invii push server->client.

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
3. Usa Capacitor (config già pronta in `capacitor.config.json`) per creare il wrapper Android in `android/`.
4. Genera il file `.aab` da Android Studio e caricalo in Play Console.

### Comandi consigliati (quando vuoi attivare Android)

```bash
npm install
npm run android:add
npm run android:sync
npm run android:open
```

`capacitor.config.json` è già incluso nel repository con `appId` `it.vargacantieri.hera`, quindi non devi inizializzarlo manualmente.

In Android Studio:
- Build > Generate Signed Bundle / APK
- Seleziona **Android App Bundle (AAB)**
- Firma e pubblica su Play Console

### Nota importante

Per mantenere la web app intatta, evita refactor lato UI/rotte solo per Android: usa plugin Capacitor solo se servono feature native (camera, notifiche, file).


## Android: geofence nativo (app chiusa)

È stato aggiunto un plugin nativo Capacitor (`HeraGeofence`) con logica lato Android per rispettare il vincolo *"anche con app spenta"*.

### Coordinate e raggio geofence

- lat: `44.562504656236015`
- lng: `11.356961975643515`
- radius: `200m`

### Logica nativa implementata

- Trigger geofence via `GeofencingClient` (Google Play Services), non via Web Geolocation.
- Gestione fasce orarie in receiver nativo (ora locale device):
  - `06:15–07:30` → notifica entrata
  - `15:30–17:00` → notifica uscita
- Deduplica persistente con `SharedPreferences` per `giorno + fascia`.
- Ripristino automatico geofence dopo riavvio/aggiornamento app (`BOOT_COMPLETED` + `MY_PACKAGE_REPLACED`).

### Permessi Android configurati

- `ACCESS_COARSE_LOCATION`
- `ACCESS_FINE_LOCATION`
- `ACCESS_BACKGROUND_LOCATION`
- `POST_NOTIFICATIONS` (Android 13+)
- `RECEIVE_BOOT_COMPLETED`

### Bridge opzionale in `app.js`

`app.js` espone (solo se plugin disponibile su Android nativo):

- `window.heraNativeGeofence.activate()`
- `window.heraNativeGeofence.deactivate()`
- `window.heraNativeGeofence.status()`

La logica di trigger notifiche resta comunque al livello nativo.

## Stabilità: cosa controllare prima di dire che è "perfetta"

In questo progetto non ci sono test automatici completi, quindi la stabilità dipende da controlli tecnici e test manuali.

### Check rapido (consigliato a ogni modifica)

```bash
npm run check:syntax
```

Questo comando verifica che `app.js` non abbia errori di sintassi JavaScript bloccanti.

### Check funzionali minimi

1. Login/logout con account Google.
2. Lettura/scrittura Firestore (creazione e aggiornamento di almeno un impianto).
3. Import Excel (`.xlsx/.xls`) con almeno un file reale.
4. Notifiche PWA (permesso + test notifica locale).
5. Se usi Android wrapper: `npm run android:sync` dopo ogni modifica web rilevante.

### Per migliorare ancora l'affidabilità

- Introdurre test end-to-end (es. Playwright) per i flussi critici.
- Aggiungere monitoraggio errori runtime (es. Sentry) per intercettare errori reali utenti.
- Definire una checklist di release con prova su browser mobile reale + almeno un dispositivo Android.
- Validare sempre configurazioni Firebase e permessi Android prima del rilascio.
