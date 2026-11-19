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

### 2.1(b) — Risposta alle cinque domande commerciali

Apple richiede risposte puntuali. Completare solo dopo aver verificato la situazione reale di distribuzione/pagamento:

1. Chi usa funzionalità a pagamento? **[indicare se nessuna o quali utenti aziendali]**
2. Dove vengono acquistate? **[non esistono acquisti oppure indicare il canale reale]**
3. Quali contenuti precedentemente acquistati sono disponibili? **[nessuno oppure elenco]**
4. Quali abbonamenti o funzionalità vengono sbloccati senza In-App Purchase? **[nessuno oppure elenco + giustificazione]**
5. Servizi venduti a consumatori singoli o famiglie? **[indicare se l'app è esclusivamente B2B e come funziona l'accesso]**

Non dichiarare "gratuita" o "nessun acquisto" senza conferma.

### Prove e pubblicazione

- Eseguire npm run check:account-privacy.
- Eseguire npm run android:aab:prepare (genera la directory www usata dal wrapper Capacitor).
- Verificare che il target iOS includa account-privacy-controls.js e il provider Apple nativo.
- Distribuire in ambiente configurato la Cloud Function: firebase deploy --only functions:deleteOwnVargaAccount.
- Preparare una build iOS successiva alla 1.0 (21) e provarla su dispositivo reale.
- Registrare una schermata reale: ingresso/registrazione demo, individuazione del pulsante Elimina account, due conferme, successo; utilizzare un account di prova eliminabile.
- Inserire il video in Note della revisione e rispondere in App Store Connect alle tre linee guida. Non inviare nuovo build finché le verifiche non sono complete.

Documentazione: https://developer.apple.com/app-store/review/guidelines/it/ e https://firebase.google.com/docs/auth/ios/apple
