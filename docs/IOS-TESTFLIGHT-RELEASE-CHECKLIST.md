# VARGA CANTIERI — percorso di rilascio iOS senza Mac personale

## Stato verificato (11 novembre 2026)
- Repository: `ionut290/hera-app`, PR in bozza **#1577**.
- Build di simulatore con Xcode 26 riuscita: run **43007574007**.
- Firebase Apple corretto incluso nella build di test manuale (mai negli artefatti dell'app).
- Avvio nel simulatore: superato al secondo screenshot; la schermata di accesso è visibile.
- Il test non usa credenziali, non prova il login, non prova l'invio WhatsApp e non prova l'app su iPhone fisico.
- Logo nativo: generato dalla icona già esistente nella fase di build, in asset PNG opachi per iOS.
- Nessuna modifica alle funzioni protette FATTO/WhatsApp, al database, alle regole Firestore o alla versione Android/web.

## Strategia di distribuzione
1. **Simulatore CI:** `.github/workflows/build-ios-simulator.yml` compila senza firma, fa i test e salva screenshot. Non pubblica nulla.
2. **TestFlight interno (solo autorizzazione manuale):** `.github/workflows/ios-private-testflight.yml` crea l'archivio con firma Apple su runner `macos-26`. Solo dopo la conferma esplicita `UPLOAD_INTERNAL_TESTFLIGHT_ONLY`, con SHA verificato e credenziali ufficiali, invia a **TestFlight Internal Only**. Non abilita il rilascio pubblico.
3. **App Store pubblico:** **non automatizzato**. Richiede test su dispositivo fisico, informativa privacy e dati di distribuzione, verifica App Review, e un'approvazione specifica separata da questa pipeline.

Il workflow privato è sul ramo di sviluppo. Non è eseguibile come nuovo workflow dal menu Actions finché GitHub non lo vede anche nel branch di default. **Non unire la PR e non avviarlo finché non sono completate le verifiche e non sono predisposti i controlli dell'ambiente.**

## Prerequisiti Apple (eseguiti dal titolare dell'account, non da ChatGPT)
1. Iscrizione attiva all'**Apple Developer Program**, con accesso ad App Store Connect.
2. Registrare l'**App ID esplicito** `it.vargacantieri.hera` nell'account/team corretto.
3. Creare in **App Store Connect → Apps** il record VARGA CANTIERI con lo stesso Bundle ID. Scegliere i dati del titolare e i diritti sul nome/logo.
4. In **App Store Connect → Users and Access → Integrations**, abilitare le API e creare una chiave adatta alla firma/distribuzione. Il file `.p8` è scaricabile una sola volta.
5. Verificare che Xcode automatic signing sul runner sia consentito dal ruolo della chiave. Se l'account non può creare o usare certificati/profili, la build fallirà senza pubblicare nulla e servirà configurazione Apple del titolare.
6. Registrazione Firebase iOS nel progetto **hera-app-6cd2b**, app Apple `it.vargacantieri.hera`, con `GoogleService-Info.plist` **scaricato da Firebase Console**. Non inventare i valori dal web o dalla configurazione Android.
7. Configurare l'ambiente GitHub `ios-internal-testflight` con **required reviewers** e restrizioni ai branch approvati. Senza protezione dell'ambiente, NON inserire credenziali Apple per il rilascio.
8. Aggiungere i seguenti **GitHub Actions Secrets** (preferibilmente environment secrets dell'ambiente protetto; non nel codice e non in chat):

| Secret | Valore |
|---|---|
| `IOS_FIREBASE_CONFIG_BASE64` | `GoogleService-Info.plist` ufficiale codificato in Base64 |
| `IOS_APPLE_TEAM_ID` | Team ID Apple Developer (10 caratteri) |
| `IOS_APPSTORE_KEY_ID` | Key ID App Store Connect API (10 caratteri) |
| `IOS_APPSTORE_ISSUER_ID` | Issuer ID della chiave di team App Store Connect |
| `IOS_APPSTORE_PRIVATE_KEY_BASE64` | File ufficiale `AuthKey_*.p8` codificato in Base64 |

L'ambiente è distinto dal database Firebase. I secrets NON devono essere aggiunti a `main`, a file di configurazione web, a PR, screenshot o log. Se non si è certi dei permessi, fermarsi prima del caricamento.

### Come codificare in Base64 senza un Mac
Su Windows PowerShell, dalla cartella contenente il **proprio** file, per incollare manualmente il risultato nei GitHub Secrets:

```powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes("GoogleService-Info.plist"))
[Convert]::ToBase64String([IO.File]::ReadAllBytes("AuthKey_IDDELLACHIAVE.p8"))
```

Non condividere queste stringhe in chat. Non incollarle nei repository, nelle issue o negli screenshot.

## Uso del workflow privato (quando tutti i prerequisiti sono soddisfatti)
1. Validare la PR #1577, i test automatici e la verifica visiva della schermata di login senza avviso GPS.
2. Verificare che il codice e la configurazione Firebase corrispondano al progetto autorizzato; aggiornare iOS Capacitor in modo sicuro prima del rilascio.
3. Rendere disponibile in modo controllato il workflow sul branch di default **solo dopo approvazione**. La produzione web non deve cambiare per il solo packaging.
4. In GitHub Actions aprire **Varga Cantieri iOS Private TestFlight (manual only)**; selezionare il commit revisionato e copiarne esattamente i 40 caratteri SHA nel campo `reviewed_commit_sha`.
5. Selezionare `UPLOAD_INTERNAL_TESTFLIGHT_ONLY` soltanto quando si è autorizzati al caricamento privato. L'ambiente protetto deve chiedere la revisione prima di fornire i segreti.
6. Il runner macOS compila e verifica l'archivio firmato, poi carica la build **TestFlight Internal Only**. Non carica file `.ipa` con Firebase come artefatti pubblici.
7. Attendere l'elaborazione Apple e aggiungere soltanto tester interni autorizzati in App Store Connect. TestFlight interno richiede account del team con il ruolo idoneo.
8. Installare la build con l'app **TestFlight** sull'iPhone reale e fare la checklist seguente.

## Checklist obbligatoria su iPhone reale
- [ ] Avvio, logout, ripristino sessione e login con account di **prova** (anche Google).
- [ ] Verifica errori WebView/console, assenza di schermata bianca e caricamento infinito.
- [ ] Localizzazione consentita/rifiutata, GPS e mappa, apertura indicazioni.
- [ ] Commesse e impianti: sola lettura iniziale, filtro, ricerca, ordinamento e percorsi.
- [ ] Foto, documenti, Drive e autorizzazioni privacy effettive.
- [ ] WhatsApp/WhatsApp Business **app installata soltanto**, senza browser; verificare contenuto completo e l'errore quando l'app non è installata.
- [ ] FATTO con **impianto di prova**, verifica di data, ora, operatore, stato e assenza di duplicati.
- [ ] Push/notifiche con APNs configurato: non dichiarare funzionante finché non collaudato.
- [ ] Connessione assente, riapertura, offline, sincronizzazione e batteria/permessi.
- [ ] Verifiche privacy delle informazioni di lavoratori, account e location; nessun dato reale condiviso con revisori.
- [ ] Prove su dispositivi/versioni iOS effettivamente supportati e screenshot finali.

**Non effettuare operazioni di prova distruttive su impianti o utenti reali.**

## Solo dopo i test: App Store
Preparare descrizione, categoria, screenshot originali per le dimensioni richieste, privacy policy, support URL, policy dati, classificazione per età, eventuale accesso demo per App Review, verifica della guideline 4.2 e dichiarazioni sulle autorizzazioni. L'app pubblica richiede una procedura ulteriore, con autorizzazione esplicita del titolare. Se il prodotto è destinato esclusivamente ai dipendenti, valutare distribuzione privata/Custom App anziché App Store pubblico.

## Fonti Apple
- https://developer.apple.com/news/upcoming-requirements/
- https://developer.apple.com/documentation/xcode/distributing-your-app-for-beta-testing-and-releases
- https://developer.apple.com/help/app-store-connect/get-started/app-store-connect-api/
- https://developer.apple.com/help/app-store-connect/test-a-beta-version/add-internal-testers
