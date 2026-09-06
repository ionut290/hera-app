# Contratto di stabilità del login

Identificatore: `LOGIN_STABILITY_CONTRACT_V1`.

## Obiettivo

Il login deve terminare sempre in uno stato visibile e comprensibile:

1. schermata di login;
2. caricamento o verifica autorizzazione;
3. Home autenticata;
4. errore recuperabile con possibilità di riprovare;
5. accesso negato o sessione scaduta.

Non è consentito lasciare contemporaneamente nascosti login, caricamento, verifica autorizzazione e Home.

## Regole operative

- Firebase Authentication resta l'unica autorità per la sessione.
- Il caricamento del profilo o di Firestore non deve impedire la notifica dello stato autenticato.
- Commesse e squadra del giorno sono dati iniziali prioritari.
- Meteo, distributori, documenti e altre funzioni secondarie non devono bloccare la Home.
- La persistenza locale della sessione deve essere configurata prima del login.
- Errori e timeout devono mostrare una superficie visibile; non devono produrre una schermata bianca.
- La protezione non deve aggiungere letture, scritture o listener Firebase.
- Il flusso deve restare verificato su PWA e Android.

## Modificabilità autorizzata

Questa protezione è forte ma non irrevocabile. Può essere modificata, aggiornata o sostituita esclusivamente dopo il consenso esplicito del proprietario dell'app. Ogni modifica autorizzata deve aggiornare consapevolmente il contratto e i relativi controlli automatici, superare tutti i test critici e preservare i dati esistenti.

Modifiche ad altre sezioni dell'app non possono cambiare indirettamente il comportamento del login.

## Estensione autorizzata: codice unico di recupero

- Il proprietario ha autorizzato il recupero della password mediante un codice amministrativo condiviso tra Varga Cantieri e Varga Gestionale.
- Il codice deve iniziare con `REC-`, non viene incorporato nel client e viene conservato dal backend esclusivamente come hash con salt.
- Il percorso viene riconosciuto prima del normale accesso Firebase e non salva il codice nel gestore delle credenziali.
- Il backend applica limitazione dei tentativi, challenge monouso con scadenza e registro di audit.
- La password precedente viene sostituita soltanto dopo che l’utente ha inserito due volte la nuova password e ha confermato esplicitamente l’operazione.
- Il codice è utilizzabile soltanto per profili operatore attivi; gli account amministratore continuano a usare il recupero Firebase tramite email.
- Il normale accesso email/password, Google, la persistenza della sessione e le superfici visibili del login restano invariati.
- Nessuna lettura, scrittura o listener Firebase viene aggiunto all'avvio: le chiamate avvengono soltanto su azione esplicita dell'utente o dell'amministratore.
