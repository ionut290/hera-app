# VARGA CANTIERI — App Review del 18 novembre 2026

Versione revisionata da Apple: 1.0 (21). Riferimento invio: f9a92b52-6580-4718-944f-47290a7a49af.

## Stato: modifiche proposte, NON pubblicate e NON validate su iPhone/iPad

Questo ramo risponde ai rilievi 5.1.1(v) e 4.8 nel codice e prepara il lavoro per 2.1(b).
Non dichiarare la conformità completa finché l'esame dei dati personali, i test su dispositivo e la configurazione dei provider non sono terminati.

### 5.1.1(v) — eliminazione account

- Pulsanti in menu personale, profilo e schermata di accesso in attesa.
- Doppia conferma; callable autenticata deleteOwnVargaAccount in europe-west1.
- Solo l'utente autenticato può eliminare il proprio account; non si accetta un UID scelto dal client.
- Richiede autenticazione recente (10 minuti).
- Elimina il profilo Firebase Authentication, platformUsers/{uid}, privateDocuments/{uid}, operatorPositions/{uid}, userOperatorLinks/{uid}, alcuni dati personali di conferma/notifiche e i documenti marcati personal con file nello storage documents/{uid}/.
- NON elimina automaticamente commesse, ore, rapporti o documenti aziendali condivisi; non deve interrompere i flussi di lavoro degli altri dipendenti.
- **Da verificare prima della pubblicazione:** inventario completo di tutti i dati personali per UID/email in Firestore, Storage, Google Drive, registri storici, esportazioni, backup e servizi esterni; criteri e tempi di conservazione per i dati di lavoro; gestione di percorsi Storage legacy non previsti. Se la cancellazione trova percorsi non sicuri, il server non conferma il completamento.
- Prima di provare, usare esclusivamente un account di test: l'eliminazione è irreversibile.

### 4.8 — Accesso con Apple

- Nuovo pulsante Accedi con Apple nel gate di accesso SOLO per l'app iOS nativa; non appare su Android né sul sito web (Service ID / flusso OAuth web non configurato).
- Su iOS usa il plugin Capacitor FirebaseAuthentication con skipNativeAuth=true, credenziale idToken + rawNonce, quindi Firebase Web signInWithCredential.
- In eliminazione account Apple, quando possibile, esegue reautenticazione Apple e revoca la credenziale con un nuovo authorizationCode.
- Mantiene email/password e Google; non offre accesso ai cantieri senza autorizzazione aziendale.

**Configurazione verificata sul portale il 19 novembre 2026:**

1. **COMPLETATO — Apple Developer:** il bundle ID `it.vargacantieri.hera`, team `NJ64236Q54`, ha `Sign in with Apple` abilitato; verifica effettuata dopo il salvataggio.
2. **COMPLETATO — Firebase Console:** nel progetto `hera-app-6cd2b` il provider Apple è attivo. Service ID, Team ID e chiave privata OAuth sono lasciati vuoti perché il flusso previsto nell'app è solo **iOS nativo**. L'accesso Apple web è nascosto finché non verrà configurato separatamente.
3. **COMPLETATO — Apple Developer:** rigenerato il profilo iOS App Store `VARGA CANTIERI - TestFlight` precedentemente `Invalid`; stato attuale `Active`, scadenza 11 novembre 2027, associato al certificato Apple Distribution esistente e al bundle corretto. Il profilo RDI non è stato modificato.
4. **ANCORA DA FARE — GitHub Actions:** scaricare il profilo App Store **rigenerato** da Apple Developer e sostituire la secret protetta `IOS_APPSTORE_PROFILE_BASE64` (base64 del nuovo `.mobileprovision`). La versione vecchia della secret non è stata modificata. Non incollare il valore in chat, repository o log pubblici.
5. **ANCORA DA FARE — dispositivo reale:** verificare l'accesso Apple ID con 2FA, anche con Nascondi la mia email, la creazione/approvazione account, la cancellazione end-to-end, la revoca Apple e la registrazione del video richiesto da App Review. Usare un account demo eliminabile, non dati reali.
6. **ANCORA DA FARE — audit dati:** inventario e retention di dati personali nel progetto Firebase, Storage, Drive e backup prima della messa in produzione della funzione di cancellazione.

Il generatore iOS include già la capability nel target, ma un build simulator senza configurazione Firebase segreta non verifica il login reale. Nessuna nuova build App Store è stata inviata.

### 2.1(b) — Modello commerciale CONFERMATO dal proprietario, 19 novembre 2026

**VARGA CANTIERI è completamente gratuita.** La conferma è stata fornita espressamente dal proprietario dell'app il 19 novembre 2026. Non sono previsti acquisti in-app, pagamenti per sbloccare funzionalità, abbonamenti o contenuti digitali a pagamento. La gestione di account e autorizzazioni per i cantieri non è un paywall.

Testo da inviare nella risposta App Review alla richiesta relativa alla linea guida 2.1(b):

> Gentile team App Review,
>
> grazie per la vostra richiesta di chiarimento. Il modello commerciale di VARGA CANTIERI è interamente gratuito. L'app è uno strumento di gestione operativa di cantieri e attività lavorative, destinato a utenti autorizzati.
>
> 1. **Chi utilizzerà i contenuti e le funzionalità a pagamento?** Nessuno: l'app non contiene funzionalità o contenuti a pagamento.
> 2. **Dove possono gli utenti acquistare i contenuti e le funzionalità accessibili nell'app?** Non esiste un canale di acquisto: tutte le funzionalità offerte dall'app sono gratuite.
> 3. **Quali contenuti e funzionalità acquistati in precedenza sono disponibili nell'app?** Nessuno. L'app non offre accesso a contenuti o funzionalità precedentemente acquistati.
> 4. **Quali contenuti a pagamento, abbonamenti o funzionalità vengono sbloccati senza acquisti in-app?** Nessuno. Non esistono abbonamenti, sblocchi a pagamento né acquisti esterni per usare le funzionalità.
> 5. **I servizi aziendali dell'app vengono venduti a singoli utenti, consumatori o famiglie?** No. L'app non vende servizi a singoli consumatori o famiglie; è gratuita ed è pensata per un utilizzo professionale da parte di utenti autorizzati.
>
> Le autorizzazioni degli utenti servono esclusivamente a gestire accessi e ruoli nell'ambito lavorativo e non sono legate a pagamenti.
>
> Cordiali saluti,
> Il team VARGA CANTIERI

Questa sezione può essere inviata come chiarimento commerciale **anche prima della nuova build**, senza dichiarare che la correzione tecnica ai punti 4.8 e 5.1.1(v) sia già pubblicata. Non dichiarare completati gli altri due punti finché non sono testati e distribuiti.

### Prove e pubblicazione

- Eseguire npm run check:account-privacy.
- Eseguire npm run android:aab:prepare (genera la directory www usata dal wrapper Capacitor).
- Verificare che il target iOS includa account-privacy-controls.js e il provider Apple nativo.
- Distribuire in ambiente configurato la Cloud Function: firebase deploy --only functions:deleteOwnVargaAccount.
- Preparare una build iOS successiva alla 1.0 (21) e provarla su dispositivo reale.
- Registrare una schermata reale: ingresso/registrazione demo, individuazione del pulsante Elimina account, due conferme, successo; utilizzare un account di prova eliminabile.
- Inserire il video in Note della revisione e rispondere in App Store Connect alle tre linee guida. Non inviare nuovo build finché le verifiche non sono complete.

Documentazione: https://developer.apple.com/app-store/review/guidelines/it/ e https://firebase.google.com/docs/auth/ios/apple
