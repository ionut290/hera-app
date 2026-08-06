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


## Layer meteo / radar

- Il provider principale dei layer meteo in mappa fullscreen è **OpenWeatherMap Weather Maps 2.0**.
- Configura la chiave API senza inserirla nel codice, esponendo `VITE_OPENWEATHER_API_KEY` a runtime prima di `app.js`. Esempio consigliato:

```html
<script>
  window.__HERA_ENV__ = { VITE_OPENWEATHER_API_KEY: "INSERISCI_LA_TUA_CHIAVE" };
</script>
```

- Sono supportati anche un meta tag (`<meta name="VITE_OPENWEATHER_API_KEY" content="...">`) o `localStorage.setItem("VITE_OPENWEATHER_API_KEY", "...")` per test locali.
- In ambienti con pipeline Vite/build puoi sostituire il placeholder `%VITE_OPENWEATHER_API_KEY%` durante il build.
- RainViewer resta disponibile solo come fallback opzionale per pioggia/nuvole quando OpenWeatherMap non è configurato o la verifica del tile OpenWeatherMap fallisce.
- Se un layer non è esposto dal provider o dal piano attivo, l'app mostra “Dato non disponibile” e lascia la mappa navigabile.
- I tile meteo usano `maxNativeZoom` e `maxZoom` separati: oltre lo zoom nativo il layer viene scalato da Leaflet invece di richiedere tile non supportati.

## Meteo operativo e impianti

- Home, navigazione e card degli impianti usano il proxy Netlify `/api/weather`.
- Il provider principale è **Open-Meteo Best Match**: seleziona automaticamente il modello disponibile a risoluzione più alta per le coordinate richieste, incluso ItaliaMeteo/ARPAE ICON-2I dove applicabile.
- Per le decisioni operative vengono richiesti dati ogni 15 minuti su pioggia, rovesci, neve, vento, raffiche, visibilità e condizioni temporalesche, più una previsione oraria di 12 ore.
- Se Open-Meteo non risponde, il proxy normalizza automaticamente i dati del provider indipendente **MET Norway Locationforecast**. Se anche il proxy non è raggiungibile, web e Android tentano Open-Meteo direttamente.
- Il proxy usa cache CDN, validazione delle coordinate e timeout per non bloccare l'interfaccia. L'app mantiene inoltre l'ultimo dato locale degli impianti.
- Per l'uso commerciale configura `OPEN_METEO_API_KEY` nelle variabili ambiente Netlify. Il proxy passerà automaticamente all'endpoint riservato `customer-api.open-meteo.com`, senza esporre la chiave nel frontend.


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

- Il Drive usato è centralizzato: **ionut29019@gmail.com** collega il proprio account una volta, poi tutti gli utenti passano dal backend Firebase.
- Solo l'utente admin (`ionut29019@gmail.com`) vede attivo il pulsante **Collega Google Drive**.
- Tutti i nuovi upload vengono salvati dentro la cartella Drive principale `1s6qmv2SsiTUbCjqFX4yIk4VoPQayFrU0` (`Varga Cantieri`).
- L'app crea o riusa sempre le cartelle `Varga Cantieri / <Nome Commessa> / <tipo file>` (es. `FOTO`, `POS`, `ORE`, `EXPORT`, `SEGNALAZIONI`) senza creare nuove cartelle duplicate per la stessa commessa/tipo.
- Quando l'admin collega Drive, eventuali dati Drive vecchi trovati nella cartella `Hera App - Dati` vengono spostati in `Varga Cantieri / VECCHI DATI` per conservarli dentro la nuova root.
- Token e refresh token admin vengono salvati in `appConfig/driveAdminSecret`; `appConfig/driveBridge` espone solo lo stato centralizzato agli utenti autenticati.
- I media chat vengono caricati su Drive tramite backend e salvati in Firestore come metadati/link.
- Per ogni commessa viene usato un solo Google Sheet (`Commessa - <nome commessa>`) dentro `Varga Cantieri / <Nome Commessa> / EXPORT`.
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
# Accesso biometrico Android

L'app usa il plugin Capacitor locale `HeraBiometric` **1.0.0**, sviluppato per le API di
Capacitor 7.2.0 già presenti nel progetto. Il plugin usa `androidx.biometric:biometric:1.1.0`
e una chiave AES non esportabile di Android Keystore. Nel browser/PWA non viene registrato
e il login Firebase/Google continua a usare il flusso web esistente.

La biometria è un blocco locale della sessione Firebase persistente: non sostituisce Firebase
Authentication e non conserva password, token Firebase, impronte o immagini del volto. Il
logout chiude sempre Firebase; la preferenza biometrica resta disponibile e può essere rimossa
dalla sezione **Sicurezza**.

Prima di compilare l'APK, verificare che `android/app/google-services.json` appartenga al progetto
Firebase `hera-app-6cd2b` e contenga il client Android `it.vargacantieri.hera`. Il file non è
incluso nel repository e va copiato localmente da Firebase Console senza rigenerare o modificare
la SHA-1 del certificato.

Comandi Android (dopo avere completato/rigenerato, se necessario, il progetto nativo Capacitor):

```sh
npm install
npx cap sync android
npx cap open android
```

In Android Studio eseguire la sincronizzazione Gradle e provare su un dispositivo reale con
blocco schermo e almeno un'impronta o volto configurato.
Test workflow GitHub Actions.
Test workflow GitHub Actions.
