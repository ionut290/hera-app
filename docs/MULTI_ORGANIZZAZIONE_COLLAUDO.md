# Collaudo obbligatorio multi-organizzazione

Questa funzione non deve essere unita in `main` finché tutti i controlli applicabili non risultano completati.

## 1. Controlli automatici

Eseguire:

```bash
npm run check:multi-organization-safe
```

Il comando comprende:

- test del modello accessi;
- test del modello Super Admin;
- test del piano di migrazione;
- controllo assenza listener/polling nei nuovi moduli;
- controllo sintassi e protezioni critiche FATTO.

## 2. Accesso utenti storici

- Utente storico non amministratore: entra direttamente in Varga.
- Amministratore storico: resta amministratore di Varga.
- Nessun utente deve registrarsi nuovamente.
- Nessun utente storico deve perdere l'accesso.

## 3. Nuove organizzazioni

- Solo il Super Admin può creare un'organizzazione.
- La nuova organizzazione parte senza commesse e impianti Varga.
- L'amministratore assegnato vede solo la propria organizzazione.
- Un utente appartenente a più organizzazioni visualizza il selettore.
- Un utente appartenente a una sola organizzazione entra direttamente.

## 4. Separazione dei dati

Verificare con due account distinti:

- commesse;
- impianti;
- squadre;
- ore;
- segnalazioni;
- documenti;
- preventivi;
- consuntivi.

Nessun dato dell'organizzazione A deve essere leggibile o modificabile dall'organizzazione B.

## 5. FATTO

Eseguire dall'elenco “Da fare”:

- premere FATTO;
- verificare salvataggio stato;
- verificare data, ora e operatore;
- verificare passaggio in “Fatti”;
- ricaricare l'app e verificare persistenza.

La logica e le funzioni esistenti non devono essere rinominate o riscritte.

## 6. WhatsApp/WHAZZUP

- Provare da un impianto “Da fare”.
- Provare da un impianto “Fatto”.
- Verificare messaggio precompilato completo.
- Verificare assenza del messaggio “Non puoi inviare un messaggio vuoto”.
- Verificare apertura corretta su telefono.

## 7. Firestore e prestazioni

Controllare:

- letture all'apertura dell'app;
- letture entrando in una commessa;
- scritture durante un salvataggio;
- richieste duplicate;
- listener registrati più volte;
- letture causate dal solo rendering;
- scritture senza variazioni;
- rimozione dei listener;
- tempi di apertura rispetto a `main`.

Il selettore organizzazione non deve creare listener realtime. I dati organizzazione devono essere caricati una volta e riutilizzati finché validi.

## 8. Migrazione utenti

Prima eseguire sempre il dry-run.

Il rapporto deve mostrare:

- utenti analizzati;
- utenti da aggiornare;
- utenti già compatibili;
- numero di scritture previste;
- nessuna scrittura duplicata.

La migrazione reale deve usare batch limitati e aggiornamenti `merge`.

## 9. Dispositivi

Verificare almeno:

- browser desktop;
- iPhone/Safari o PWA;
- Android/Chrome o wrapper Capacitor;
- orientamento verticale;
- connessione lenta;
- riapertura dopo logout/login.

## 10. Pubblicazione

Prima del merge:

- controllare `git diff` o diff PR;
- verificare che i file FATTO/WhatsApp non siano modificati;
- verificare CI e deploy preview;
- distribuire le regole Firestore soltanto dopo test con account separati;
- non eseguire migrazione e cambio regole nello stesso passaggio senza possibilità di rollback.
