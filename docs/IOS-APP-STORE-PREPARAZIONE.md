# VARGA CANTIERI — preparazione iPhone / App Store

Stato: **prototipo tecnico su ramo separato**, NON distribuibile e NON firmato.

Ramo di lavoro: `feature/ios-app-store-prep-20261111`
Repository: `ionut290/hera-app`
Bundle ID proposto (da confermare nell'account Apple): `it.vargacantieri.hera`.

## Obiettivo e perimetro

Preparare il progetto web esistente con Capacitor 7 per una futura distribuzione iOS
senza richiedere un Mac personale e senza modificare la versione web/Android in uso.
Non vengono toccati `app.js`, login, flusso FATTO/Whazzup, stati impianti,
Firestore, Google Drive, dati operativi o regole di accesso.

Un workflow GitHub Actions su runner macOS genera temporaneamente il progetto
nativo iOS, applica le descrizioni dei permessi di sistema e prova a compilare
una **app per iOS Simulator priva di firma**. L'artefatto di questo workflow
**non è un IPA per iPhone**, non è installabile tramite App Store o TestFlight
e non costituisce una pubblicazione. Nessun secret Apple è richiesto o usato.

## Prova automatica

Workflow: `.github/workflows/build-ios-simulator.yml`

Controlla:
- test di regressione sui flussi FATTO e WhatsApp protetti;
- stabilità del login e controlli privacy/sicurezza presenti nel repo;
- sintassi JavaScript e generazione del bundle web;
- compatibilità del progetto generato con Capacitor iOS;
- compilazione iOS Simulator senza firma con `xcodebuild`;
- installazione e lancio senza login nel simulatore, con screenshot scaricabile per revisione manuale.

Per evitare una variazione delle dipendenze della produzione, solo il runner iOS
installa temporaneamente `@capacitor/ios` e `@capacitor/core` **7.6.10**
(mediante `npm install --no-save --package-lock=false`). Questa è una
versione corretta rispetto a CVE-2026-103922 / GHSA-rvm3-566m-v7fv.
L'attuale `package-lock.json` contiene una versione nativa Android 7.6.8:
**prima di qualunque rilascio** servirà aggiornare e verificare separatamente
tutte le dipendenze Capacitor vulnerabili, senza toccare le logiche protette.

Lo stato della build si consulta in GitHub → Actions →
**Varga Cantieri iOS Simulator (no deploy)**.

## Creazione locale su un Mac / runner macOS

Richiede Node.js 22+, **Xcode 26 o successivo (SDK iOS 26+)** e CocoaPods.\nIl workflow gira su `macos-26` e verifica esplicitamente che Xcode e SDK\nsoddisfino i minimi Apple vigenti da aprile 2026. La compilazione con Xcode\n16.4 era soltanto una prova tecnica e non è idonea a App Store Connect.

```bash
npm ci
npm run android:aab:prepare
npm install --no-save --package-lock=false @capacitor/ios@7.6.10 @capacitor/core@7.6.10
npx cap add ios
npx cap sync ios
python3 scripts/configure-ios-plist.py
open ios/App/App.xcworkspace
```

I comandi non effettuano upload, firma, deploy o modifica dei database.
Per le prove iPhone è necessario configurare Apple signing in Xcode/CI.

## Test simulatore e limiti\n\nIl comando `bash scripts/smoke-ios-simulator.sh /percorso/App.app /percorso/screenshot.png`\ninstalla e avvia il pacchetto NON firmato su un simulatore iPhone disponibile.\nLa CI carica una schermata iniziale in GitHub Actions. Il test **non**\ndimostra che login, Firebase, interfaccia interattiva o WhatsApp funzionino\nsu un dispositivo reale e non rileva automaticamente una schermata bianca.\nLa screenshot richiede una revisione visiva prima della pubblicazione.\n\n## Blocchi da superare PRIMA della pubblicazione

1. **Login iPhone:** verificare su un dispositivo reale il login Google/Firebase
   in WKWebView e la ripresa di sessione. La configurazione Android non garantisce
   il login nativo iOS; non modificare il login senza autorizzazione esplicita.
2. **FATTO / WhatsApp:** testare sul dispositivo reale il flusso protetto.
   Il plugin WhatsApp nativo presente nel repository è Android-specifico.
   Non presumere equivalenza su iOS e non inserire fallback a browser:
   il flusso sigillato non deve essere riscritto o aggirato.
3. **Foto, mappa, posizione, Google Drive, file e notifiche:** verificare i
   permessi effettivi e i flussi reali su iOS. Le descrizioni in `Info.plist`
   non abilitano da sole le funzioni; APNs e firma push sono ancora da predisporre.
4. **Branding e revisione Apple:** creare un'icona iOS originale 1024 × 1024,
   schermate App Store, informativa privacy aggiornata e dati del soggetto
   titolare; riesaminare l'informativa anche per i dati dei lavoratori.
5. **Accesso revisori:** prevedere un account demo e dati di esempio privi di
   informazioni aziendali reali. Evitare upload di segreti e dataset reali nei log.
6. **App Review 4.2:** dimostrare caratteristiche app-like e integrazione
   con funzioni native; Apple può rifiutare un semplice sito incorporato.
7. **Firma e consegna:** disporre di iscrizione Apple Developer, Bundle ID
   univoco, provisioning/certificati e compilazione firmata per dispositivi.
   Solo dopo prova reale preparare un IPA e caricarlo in App Store Connect,
   inizialmente su TestFlight; nessuna distribuzione automatica.

## Rilascio: criteri di superamento

- Build iOS Simulator riuscita **e** prova iPhone reale documentata;
- login stabile, nessuna schermata bianca, dati esistenti preservati;
- ricerca, impianti, mappe, squadra, foto, Firebase e Drive verificati;
- FATTO e WhatsApp protetti funzionanti senza regressioni;
- nessuna lettura/scrittura/listener Firestore aggiunti da questo intervento;
- App Store privacy, autorizzazioni, TestFlight, icone e sicurezza approvati;
- approvazione esplicita del proprietario prima di pubblicare.

## Riferimenti

- https://developer.apple.com/news/upcoming-requirements/\n- https://github.blog/changelog/2026-02-26-macos-26-is-now-generally-available-for-github-hosted-runners/\n- https://capacitorjs.com/docs/v7/ios
- https://developer.apple.com/app-store/review/guidelines/
- https://developer.apple.com/testflight/
- https://github.com/advisories/GHSA-rvm3-566m-v7fv
