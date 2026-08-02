# Hera App — Piano multi-organizzazione

## Stato

Questa implementazione viene sviluppata esclusivamente sul branch `agent/multi-organizzazione`.
Il branch `main` e l'app pubblicata non devono essere modificati finché migrazione e test non sono completati.

## Organizzazione predefinita

- ID: `varga`
- Nome: `Varga`
- Tutti gli utenti e tutti i dati esistenti devono continuare a essere considerati appartenenti a Varga durante la migrazione.
- I record storici privi di `organizationId` devono essere letti come record Varga durante la fase di compatibilità.

## Ruoli

- `super_admin`: può creare e gestire organizzazioni.
- `organization_admin`: amministra esclusivamente le organizzazioni assegnate.
- `operator`: utilizza esclusivamente le organizzazioni assegnate.

Il ruolo Super Admin deve essere verificato tramite Firebase UID e regole Firestore, non solo tramite elementi visivi dell'interfaccia.

## Principi di separazione

Ogni dato operativo deve appartenere a una sola organizzazione tramite `organizationId` oppure tramite un percorso organizzazione equivalente.

Ambiti da migrare e verificare:

- commesse;
- impianti;
- squadre;
- ore;
- segnalazioni;
- documenti;
- notifiche;
- programmazioni;
- preventivi;
- consuntivi;
- prezziari;
- mezzi e personale, chiarendo quali dati restano condivisi e quali sono organizzativi.

## Compatibilità utenti esistenti

Durante la fase transitoria:

1. un utente esistente senza membership esplicita viene considerato membro di Varga;
2. un dato esistente senza `organizationId` viene considerato appartenente a Varga;
3. chi appartiene a una sola organizzazione entra direttamente;
4. chi appartiene a più organizzazioni deve scegliere lo spazio attivo;
5. nessun utente attuale deve registrarsi nuovamente.

## Protezioni obbligatorie

Non modificare, rinominare, spostare o riscrivere:

- logica e funzioni del pulsante FATTO;
- salvataggio dello stato completato;
- data, ora e operatore del completamento;
- spostamento impianti tra Da fare e Fatti;
- generazione e apertura del messaggio WhatsApp/WHAZZUP;
- WhatsApp degli impianti già completati.

La separazione per organizzazione deve essere introdotta a monte nella risoluzione dei dati, senza riscrivere queste funzioni.

## Firestore

Prima dell'attivazione sono obbligatori:

- inventario completo di letture, scritture, query e listener;
- verifica dei listener duplicati;
- regole che consentano accesso solo ai membri dell'organizzazione;
- regole separate per il Super Admin;
- nessuna lettura globale di tutte le organizzazioni per gli amministratori normali;
- nessuna scrittura se il dato non cambia;
- migrazione idempotente e ripetibile senza duplicare documenti.

## Strategia di migrazione

1. Creare il contesto organizzazione isolato e inattivo.
2. Inventariare tutti i punti dati dell'app.
3. Preparare regole Firestore e script di migrazione senza distribuirli.
4. Aggiungere compatibilità Varga ai moduli uno alla volta.
5. Testare ogni area prima di proseguire.
6. Creare pannello Super Admin e selettore organizzazione.
7. Eseguire test completi su un ambiente separato.
8. Integrare in `main` solo dopo esito positivo.

## Controlli prima della pubblicazione

- avvio app senza schermata bianca;
- login e autorizzazioni;
- commesse e impianti visibili;
- ricerca e ordinamento;
- FATTO da Da fare;
- salvataggio stato, data, ora e operatore;
- passaggio in Fatti;
- WhatsApp precompilato e non vuoto;
- WhatsApp negli impianti Fatti;
- desktop e telefono;
- assenza di loop e caricamenti infiniti;
- assenza di nuovi listener o accessi Firestore duplicati;
- compatibilità con tutti i dati esistenti.
