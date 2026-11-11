# VARGA CANTIERI — percorso di rilascio iOS senza Mac personale

## Stato verificato (11 novembre 2026)
- Repository: `ionut290/hera-app`, PR #1577 integrata; la modifica della firma manuale è in verifica nella **PR #1580**.
- Build di simulatore con Xcode 26 riuscita: run **43007574007**.
- Firebase Apple corretto incluso nella build di test manuale (mai negli artefatti dell'app).
- Avvio nel simulatore: superato al secondo screenshot; la schermata di accesso è visibile.
- Il test non usa credenziali, non prova il login, non prova l'invio WhatsApp e non prova l'app su iPhone fisico.
- Logo nativo: generato dalla icona già esistente nella fase di build, in asset PNG opachi per iOS.
- Nessuna modifica alle funzioni protette FATTO/WhatsApp, al database, alle regole Firestore o alla versione Android/web.
- Tentativi TestFlight #1 e #2: archive firmato non riuscito per conflitti automatic signing; #3, al secondo tentativo, ha ancora un conflitto App automatic signing. **Nessuna build caricata in App Store Connect.**

## Strategia di distribuzione
1. **CI PR senza segreto Firebase Apple:** `.github/workflows/build-ios-simulator.yml` esegue i controlli protetti, verifica l'icona e **compila soltanto** l'app simulatore non firmata; non afferma che l'app si sia aperta e non pubblica nulla. Quando il vero `IOS_FIREBASE_CONFIG_BASE64` è disponibile in una esecuzione **manuale autorizzata**, il workflow installa, avvia e salva screenshot per revisione umana. Il test non esegue il login.
2. **TestFlight interno (solo autorizzazione manuale):** `.github/workflows/ios-private-testflight.yml` crea l'archivio con firma Apple su runner `macos-26`. Solo dopo la conferma esplicita `UPLOAD_INTERNAL_TESTFLIGHT_ONLY`, con SHA verificato e credenziali ufficiali, invia a **TestFlight Internal Only**. Non abilita il rilascio pubblico.
3. **App Store pubblico:** **non automatizzato**. Richiede test su dispositivo fisico, informativa privacy e dati di distribuzione, verifica App Review, e un'approvazione specifica separata da questa pipeline.

Il workflow manuale è già nel branch `main`, ma la correzione del provisioning è in **PR #1580**. **Non avviare un nuovo TestFlight** finché PR #1580 non supera CI/review, viene integrata e sono presenti i tre nuovi secrets con certificato e profilo di distribuzione.

## Prerequisiti Apple (eseguiti dal titolare dell'account, non da ChatGPT)
1. Iscrizione attiva all'**Apple Developer Program**, con accesso ad App Store Connect.
2. Registrare l'**App ID esplicito** `it.vargacantieri.hera` nell'account/team corretto.
3. Creare in **App Store Connect → Apps** il record VARGA CANTIERI con lo stesso Bundle ID. Scegliere i dati del titolare e i diritti sul nome/logo.
4. In **App Store Connect → Users and Access → Integrations**, abilitare le API e creare una chiave adatta alla firma/distribuzione. Il file `.p8` è scaricabile una sola volta.
5. Per il runner GitHub, preparare un certificato **Apple Distribution** con relativa chiave privata esportata come `.p12` e un profilo **App Store Connect** per `it.vargacantieri.hera` in formato `.mobileprovision`, creato usando esattamente quel certificato. La chiave API `.p8` autentica l'upload, ma **non è il certificato di firma**. La preparazione è possibile anche da Windows generando una CSR locale e conservando la relativa chiave privata; non caricare la chiave sul repository.
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
| `IOS_DISTRIBUTION_CERTIFICATE_BASE64` | Certificato Apple Distribution **con chiave privata**, file `.p12`, Base64 |
| `IOS_DISTRIBUTION_CERTIFICATE_PASSWORD` | Password privata del file `.p12` |
| `IOS_APPSTORE_PROFILE_BASE64` | Profilo Apple App Store Connect `.mobileprovision` per bundle e team corretti, Base64 |

L'ambiente è distinto dal database Firebase. I secrets NON devono essere aggiunti a `main`, a file di configurazione web, a PR, screenshot o log. Se non si è certi dei permessi, fermarsi prima del caricamento.

### Come codificare in Base64 senza un Mac
Su Windows PowerShell, dalla cartella contenente il **proprio** file, per incollare manualmente il risultato nei GitHub Secrets:

```powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes("GoogleService-Info.plist"))
[Convert]::ToBase64String([IO.File]::ReadAllBytes("AuthKey_IDDELLACHIAVE.p8"))
```

Per `.p12` e `.mobileprovision` si usa la stessa conversione Base64 puntando ai rispettivi file, ma è preferibile copiare direttamente negli appunti con `| Set-Clipboard`, senza stampare la stringa in console. Non condividere queste stringhe in chat. Non incollarle nei repository, nelle issue o negli screenshot.

**Nota sicurezza:** il workflow verifica sul runner che certificato e profilo corrispondano, che il provisioning non consenta debugging e che non si tratti di profilo di sviluppo, ad hoc o enterprise. Usa un keychain temporaneo che viene eliminato alla fine.

## Uso del workflow privato (quando tutti i prerequisiti sono soddisfatti)
1. Validare la PR #1580 e i suoi nuovi test, inclusa la verifica della firma isolata di App Release.
2. Verificare che il codice e la configurazione Firebase corrispondano al progetto autorizzato; aggiornare iOS Capacitor in modo sicuro prima del rilascio.
3. Integrare la correzione nel branch `main` **solo dopo approvazione e CI**, e confermare la presenza dei **8 environment secrets**. La produzione web non deve cambiare per il solo packaging.
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
- https://developer.apple.com/help/account/provisioning-profiles/create-an-app-store-provisioning-profile/
- https://docs.github.com/en/actions/how-tos/deploy/deploy-to-third-party-platforms/sign-xcode-applications
