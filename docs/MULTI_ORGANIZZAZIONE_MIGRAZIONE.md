# Migrazione multi-organizzazione — procedura sicura

## Stato

Questa procedura è preparatoria. Nessuna migrazione è attiva e nessuna scrittura viene eseguita automaticamente.

## Obiettivo

- creare l'organizzazione `varga`;
- mantenere tutti gli utenti attuali con accesso a Varga;
- mantenere il ruolo amministratore degli amministratori storici;
- assegnare gli altri utenti come operatori;
- non modificare i dati operativi finché la compatibilità non è verificata;
- rendere Varga l'organizzazione predefinita per i profili storici.

## Dry-run obbligatorio

Il modulo `multi-organization-migration-plan.js` produce soltanto un piano di scritture. Non importa Firebase e non chiama Firestore.

Il report indica:

- utenti analizzati;
- utenti da migrare;
- utenti già migrati e quindi saltati;
- numero totale di scritture previste;
- percorso e contenuto di ogni scrittura pianificata.

## Scritture previste per un utente storico

Per ogni utente non ancora migrato sono previste due scritture `merge`:

1. aggiornamento del profilo `platformUsers/{uid}` con organizzazione predefinita e membership Varga;
2. creazione/aggiornamento di `organizations/varga/members/{uid}`.

L'uso di `merge` evita di cancellare campi esistenti.

## Bootstrap Varga

Sono previste tre scritture iniziali:

1. `organizations/varga`;
2. `platformSuperAdmins/{uidSuperAdmin}`;
3. `organizations/varga/members/{uidSuperAdmin}`.

Il Firebase UID del Super Admin deve essere verificato prima di qualsiasi esecuzione reale. Non deve essere dedotto dall'email.

## Dati operativi

In questa fase non vengono spostate commesse, impianti, squadre, ore, segnalazioni, documenti, preventivi o consuntivi.

La compatibilità iniziale sarà ottenuta con una modalità legacy dedicata a Varga. Solo dopo i test verrà valutata la migrazione fisica delle collezioni, una alla volta.

## Protezioni

- nessuna scrittura senza dry-run approvato;
- nessuna cancellazione;
- nessuna sovrascrittura completa dei profili;
- nessuna modifica delle funzioni FATTO;
- nessuna modifica delle funzioni WhatsApp/WHAZZUP;
- nessuna distribuzione delle nuove regole Firestore prima dei test su emulatori;
- nessun merge su `main` con controlli incompleti.

## Controlli prima dell'esecuzione reale

- verificare il Firebase UID del Super Admin;
- confrontare il numero di utenti Firestore con il report dry-run;
- verificare manualmente un amministratore e un operatore;
- confermare che gli utenti già dotati di membership siano saltati;
- eseguire un backup/esportazione del database;
- testare le regole Firestore su emulatore;
- verificare letture e scritture previste;
- interrompere l'operazione al primo errore.
