# Stato finale branch multi-organizzazione

## Implementato

- modello organizzazioni e ruoli;
- compatibilità automatica utenti storici con Varga;
- selettore organizzazione isolato;
- pannello Super Admin autonomo;
- creazione organizzazioni tramite batch Firestore;
- assegnazione facoltativa dell'amministratore iniziale;
- dry-run migrazione utenti;
- applicazione migrazione solo dopo conferma esplicita;
- scritture batch in gruppi limitati;
- nessun listener Firestore realtime;
- nessun polling;
- bozza regole Firestore multi-organizzazione;
- controlli statici dedicati.

## Non attivato intenzionalmente

- nessun merge in `main`;
- nessuna distribuzione delle nuove regole Firestore;
- nessuna migrazione reale degli utenti;
- nessuno spostamento fisico delle collezioni operative;
- nessun collegamento del selettore ai flussi FATTO o WhatsApp/WHAZZUP.

Questi punti restano disattivati finché non vengono completati i test con un deploy preview funzionante e con regole Firestore provate in ambiente controllato.

## Costi Firestore

Il pannello usa letture una tantum attivate manualmente. Non usa `onSnapshot`, `setInterval` o polling. La creazione di una nuova organizzazione genera una scrittura per l'organizzazione e, solo quando viene indicato un amministratore, due scritture aggiuntive. La migrazione mostra prima il numero previsto di scritture e parte soltanto dopo conferma.

## Protezioni

I file operativi esistenti non sono stati modificati. In particolare non sono stati modificati `app.js`, `fatto-button-immediate.js`, i generatori dei messaggi WhatsApp/WHAZZUP o le funzioni di completamento impianti.
