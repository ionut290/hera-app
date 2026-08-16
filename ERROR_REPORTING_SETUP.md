# Diagnostica automatica errori via email

Il sistema è composto da:

- `client-error-reporter.js`: intercetta errori JavaScript, Promise non gestite e mancato caricamento di script/CSS locali;
- `functions/error-reporting.js`: Cloud Function autenticata che classifica il problema e invia la diagnosi all'amministratore;
- coda locale offline e deduplicazione locale per evitare invii ripetuti;
- nessuna lettura, scrittura o listener Firestore.

## Sicurezza

La chiave del servizio email non deve essere inserita nel codice client, nel repository o nei file Netlify.
La Cloud Function usa Firebase Secret Manager.

## Configurazione richiesta una sola volta

Dalla cartella del progetto, autenticati con Firebase CLI e imposta i due secret:

```bash
firebase functions:secrets:set RESEND_API_KEY --project hera-app-6cd2b
firebase functions:secrets:set ERROR_REPORT_FROM --project hera-app-6cd2b
```

- `RESEND_API_KEY`: API key Resend.
- `ERROR_REPORT_FROM`: mittente autorizzato da Resend, per esempio `Varga Cantieri <errori@dominio-verificato.it>`.

Il destinatario amministratore configurato nel backend è `ionut29019@gmail.com`.

## Deploy iniziale

Dopo aver configurato i secret:

```bash
firebase deploy --only functions:reportClientError --project hera-app-6cd2b
```

Solo dopo un test email riuscito si può aggiungere `functions:reportClientError` alla lista delle funzioni del workflow `.github/workflows/deploy-firebase-functions.yml`.

## Test funzionale manuale

Con un utente autenticato, dalla console browser della preview eseguire:

```js
HeraClientErrorReporter.report(new Error("TEST DIAGNOSTICA VARGA CANTIERI"), {
  kind: "manual-test",
  source: "test-manuale"
});
```

Verificare:

1. arrivo di una sola email all'amministratore;
2. presenza di utente, data/ora, schermata, dispositivo, messaggio, stack e diagnosi automatica;
3. nuovo invio dello stesso errore entro 30 minuti non produce una seconda email dal medesimo dispositivo;
4. offline: il report resta nella coda locale e viene ritentato quando l'utente torna online/autenticato;
5. nessun dato GPS, password, token o chiave API compare nell'email.

## Limiti intenzionali

- la coda locale conserva al massimo 10 errori e scarta quelli oltre 24 ore;
- il backend accetta al massimo 8 report per utente in 10 minuti per istanza attiva;
- la deduplicazione locale dello stesso errore dura 30 minuti;
- la Cloud Function accetta solo utenti Firebase autenticati;
- non viene usato Firestore per la diagnostica, così non vengono introdotti costi di lettura/scrittura o listener.
