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

- Nuovo pulsante Accedi con Apple nel gate di accesso per iOS e web; non appare su Android.
- Su iOS usa il plugin Capacitor FirebaseAuthentication con skipNativeAuth=true, credenziale idToken + rawNonce, quindi Firebase Web signInWithCredential.
- In eliminazione account Apple, quando possibile, esegue reautenticazione Apple e revoca la credenziale con un nuovo authorizationCode.
- Mantiene email/password e Google; non offre accesso ai cantieri senza autorizzazione aziendale.

**Configurazione ancora necessaria fuori dal repository:**

1. In Apple Developer, abilita Sign in with Apple per il bundle ID it.vargacantieri.hera e verifica che il target iOS includa la capability/entitlement corrispondente. Il progetto ios/ NON è versionato qui.
2. In Firebase Console (hera-app-6cd2b), Authentication > Sign-in method, attiva Apple e configura Service ID, Apple Team ID, Key ID e chiave privata secondo la guida Firebase. Non committare mai chiavi private.
3. Verifica su un vero iPhone/iPad Apple ID con 2FA, inclusa l'opzione Nascondi la mia email, il flusso di primo accesso, accesso successivo, autorizzazione aziendale e cancellazione.
4. Testa l'accesso con il profilo di revisione già creato dagli strumenti interni, senza esporre dati reali.

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
