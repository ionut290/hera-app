# Recupero password autonomo — VARGA CANTIERI (20 novembre 2026)

## Percorso per l'utente

**Login → PASSWORD DIMENTICATA? → inserisci email → Invia link di recupero.**

1. La schermata di recupero è separata e non richiede di inserire la
   password attuale né rivolgersi all'amministratore. L'email digitata nel
   login, se presente, viene precompilata ma può essere cambiata.
2. Firebase Authentication invia il proprio link monouso all'indirizzo
   di accesso **solo se esiste un account di tipo email/password**.
   L'utente apre il messaggio ricevuto, segue il link Firebase sicuro,
   imposta la nuova password e ritorna all'app per accedere.
3. Gli utenti con Google o Apple usano rispettivamente **Accedi con
   Google** / **Accedi con Apple**, senza tentare di cambiare una
   inesistente password dell'app.
4. Messaggio anti-enumerazione: l'interfaccia non comunica se
   un'email è registrata o meno. Nessuna password, link OTP o codice
   di recupero è archiviato nel browser, Firestore, GitHub o Analytics.
5. La UI blocca i doppi invii, impone una pausa locale di 5 minuti
   dopo richieste accettate/non confermate, indica Spam e consiglia di
   non condividere link e codici. **La pausa locale NON sostituisce
   il controllo server-side anti-abuso gestito da Firebase.**

## Motivo del cambiamento di sicurezza

L'interfaccia precedente richiedeva di compilare il login e costruiva un
`continueUrl` dalla pagina corrente (non utilizzabile uniformemente in
WKWebView iPhone); un parametro `passwordReset=complete` poteva
generare il messaggio **"Password aggiornata" senza verifica effettiva**.

Ora si utilizza **solo** `firebase.auth().sendPasswordResetEmail(email)`
senza URL dinamici, deep link non validati né successo artificiale.
Firebase gestisce la verifica del codice monouso e la pagina di
impostazione della nuova password. L'utente ritorna all'app manualmente
e accede con le nuove credenziali. Nessuna nuova Cloud Function,
nessun URL del token esposto nel client.

### Cosa verificare prima di dichiarare il flusso pronto per iPhone

- Il provider **Email/Password** è effettivamente abilitato nella console
  Firebase e il template email di reset ha mittente, lingua e dominio
  chiaramente riconoscibili. Confermare anche **Email Enumeration
  Protection** nelle impostazioni Authentication > User actions.
- Su un **account dimostrativo di cui l'utente controlla la casella
  email**: invio, ricezione, link aperto, password realmente cambiata e
  login riuscito. Non usare email o password aziendali reali nei test.
- Provare link scaduto, connessione interrotta, casella inesistente,
  Google/Apple login e pulsante di ritorno. Verificare accessibilità su iPhone.
- L'alternativa **REC-** amministrativa già presente resta una procedura
  distinta e **non è un recupero autonomo tramite email**; va valutata
  separatamente con il responsabile della sicurezza prima della
  pubblicazione. Non condividerne il codice.

### Limiti

Questa modifica è solo codice client con test offline. Dopo CI e merge,
la web app riceverà la modifica; un nuovo pacchetto iPhone la riceverà
soltanto dopo compilazione e caricamento su TestFlight.

I test `scripts/check-self-service-password-reset-runtime.cjs` usano
Firebase Auth simulato e **non inviano messaggi**.
