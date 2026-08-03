# AGENTS.md — REGOLE OBBLIGATORIE HERA APP

## PROGETTO

- Applicazione: `https://creative-syrniki-dddbae.netlify.app`
- Repository: `https://github.com/ionut290/hera-app`
- Branch principale: `main`

Queste istruzioni sono obbligatorie per qualsiasi agente, sviluppatore o sistema automatico che analizzi o modifichi il progetto.

## 1. PRIORITÀ ASSOLUTA: FATTO, WHATSAPP E GESTIONE IMPIANTI

Le parti più importanti dell’applicazione sono:

1. il pulsante **FATTO**;
2. il pulsante **WhatsApp/WHAZZUP**;
3. la gestione degli impianti;
4. il passaggio degli impianti tra **Da fare** e **Fatti**;
5. il mantenimento di data, ora e operatore;
6. il messaggio WhatsApp precompilato;
7. la compatibilità con i dati già presenti.

Queste funzionalità devono essere considerate **PROTETTE**.

Non devono essere modificate, riscritte, spostate, rinominate, duplicate o semplificate, salvo richiesta esplicita dell’amministratore del progetto.

## 2. PROTEZIONE ASSOLUTA DEL PULSANTE FATTO

È vietato modificare direttamente o indirettamente:

- le funzioni JavaScript collegate al pulsante FATTO;
- gli eventi e i listener del pulsante FATTO;
- i controlli che permettono o impediscono il completamento;
- il salvataggio dello stato FATTO;
- la registrazione della data;
- la registrazione dell’ora;
- la registrazione dell’operatore;
- il salvataggio delle note collegate;
- il trasferimento dell’impianto da “Da fare” a “Fatti”;
- il caricamento degli impianti completati;
- la compatibilità con gli impianti già completati;
- eventuali controlli di posizione o distanza già esistenti;
- la sincronizzazione Firestore collegata al completamento.

Non rinominare, spostare, eliminare o duplicare le funzioni usate dal pulsante FATTO.

Non sostituire la logica esistente con una nuova implementazione apparentemente equivalente.

Non eseguire refactoring sulle funzioni FATTO durante modifiche che riguardano altre sezioni.

Se una modifica può interferire anche indirettamente con FATTO, interrompere il lavoro e segnalarlo prima di procedere.

## 3. PROTEZIONE ASSOLUTA DI WHATSAPP/WHAZZUP

È vietato modificare direttamente o indirettamente:

- la funzione che crea il messaggio WhatsApp;
- il testo e la struttura del messaggio precompilato;
- i dati inseriti nel messaggio;
- la codifica del testo;
- l’apertura di WhatsApp;
- i collegamenti `wa.me`;
- la gestione di WhatsApp da telefono;
- la gestione di WhatsApp da desktop;
- il pulsante WhatsApp degli impianti “Da fare”;
- il pulsante WhatsApp degli impianti “Fatti”;
- il comportamento successivo alla pressione di FATTO;
- gli eventuali collegamenti di navigazione presenti nel messaggio.

Non devono essere introdotte condizioni che possano produrre:

- messaggi vuoti;
- messaggi incompleti;
- perdita di impianto, comune, data, ora o operatore;
- errore “Non puoi inviare un messaggio vuoto”;
- apertura di WhatsApp senza testo;
- differenze di comportamento tra impianti da fare e impianti fatti.

Non rinominare, spostare, eliminare o duplicare le funzioni utilizzate da WhatsApp/WHAZZUP.

## 4. PROTEZIONE DELLA GESTIONE IMPIANTI

Non modificare senza richiesta esplicita:

- la struttura dei dati degli impianti;
- gli ID degli impianti;
- i nomi dei campi già utilizzati;
- la relazione tra impianto e commessa;
- lo stato “Da fare”;
- lo stato “Fatto”;
- l’elenco degli impianti completati;
- la ricerca degli impianti;
- l’ordinamento degli impianti;
- l’ordinamento per distanza;
- la visualizzazione nelle commesse;
- la visualizzazione sulla mappa;
- coordinate, indirizzi e dati di navigazione;
- filtri e raggruppamenti esistenti;
- compatibilità con i dati Firestore già salvati;
- funzionamento offline e sincronizzazione successiva.

Non eliminare campi apparentemente inutilizzati senza aver verificato tutte le letture, scritture e dipendenze.

Non cambiare il formato dei documenti Firestore esistenti senza una migrazione sicura e approvata.

Non sostituire gli identificativi esistenti con nuovi identificativi generati automaticamente.

## 4-BIS. PROTEZIONE ASSOLUTA E IRREVOCABILE DI PERSONALE, UTENTI E COLLEGAMENTI

Il **Personale**, gli **Utenti**, l’**elenco Personale**, l’**elenco Utenti** e ogni dato o collegamento associato sono dati critici e permanenti.

**NON DEVONO MAI ESSERE ELIMINATI, SVUOTATI, AZZERATI, SOVRASCRITTI IN MASSA, SCOLLEGATI O RICREATI CON NUOVI ID, IN NESSUN CASO.**

Questa protezione comprende obbligatoriamente:

- tutti i documenti e record del Personale;
- tutti i documenti e profili Utente;
- gli account Firebase Authentication;
- gli ID storici degli operatori e degli utenti;
- nomi, cognomi, email, telefoni, ruoli, permessi e abilitazioni;
- collegamenti tra Personale e Utenti;
- `LINKED_USER_ID`, `LINKED_USER_EMAIL`, `EMAIL_ACCESSO_APP` e campi equivalenti;
- appartenenza e collegamenti alle squadre;
- storico delle squadre;
- collegamenti con calendario, presenze, ore, timbrature, assenze, ferie e malattie;
- collegamenti con commesse, impianti, mezzi, notifiche, documenti e attività;
- dati storici, record inattivi e relazioni non immediatamente visibili;
- qualsiasi riferimento usato da altre collezioni, moduli o funzioni dell’app.

È vietato introdurre o eseguire:

- `delete()`, `batch.delete()` o cancellazioni equivalenti su Personale, Utenti e relativi collegamenti;
- svuotamenti di collezioni, array, mappe, cache persistenti o elenchi;
- cancellazioni massive o automatiche;
- cancellazioni a cascata;
- script “esegui una volta” che eliminano dati;
- sostituzioni dell’intero elenco con un elenco vuoto o incompleto;
- importazioni che eliminano i record non presenti nel file importato;
- sincronizzazioni che considerano il file esterno come sostituzione distruttiva del database;
- pulizie, deduplicazioni o ottimizzazioni che cancellano record;
- migrazioni che generano nuovi ID al posto degli ID esistenti;
- scollegamenti automatici tra personale, utenti, squadre, calendario, ore o commesse;
- rimozioni motivate dalla riduzione di letture, scritture o costi Firestore.

Questa regola vale anche durante:

- importazioni Excel o Google Sheets;
- sincronizzazioni Firestore;
- ripristini e backup;
- migrazioni;
- ottimizzazioni;
- deduplicazioni;
- test;
- correzioni urgenti;
- refactoring;
- funzioni automatiche, pianificate o temporanee.

Quando un operatore non è più attivo, il record deve essere **mantenuto** e può essere soltanto marcato come inattivo, cessato, archiviato o non abilitato. Non deve essere eliminato e non devono essere rimossi i suoi collegamenti storici.

In presenza di duplicati, non è consentita la cancellazione automatica. È obbligatorio usare una procedura non distruttiva che preservi entrambi i record, gli ID e tutti i collegamenti, oppure unisca esclusivamente i dati mancanti senza eliminare alcun documento.

Ogni importazione o sincronizzazione deve operare in modalità **aggiunta o aggiornamento non distruttivo**. I record assenti dal file importato devono rimanere invariati nel database.

Prima di qualsiasi modifica che coinvolga questi dati è obbligatorio:

1. verificare il numero di record prima dell’intervento;
2. verificare tutti gli ID e le relazioni;
3. creare o verificare un backup utilizzabile;
4. assicurarsi che non esista alcuna operazione di eliminazione;
5. controllare squadre, calendario, ore e collegamenti utente;
6. verificare che il numero dei record non diminuisca dopo l’intervento.

Se una richiesta o una modifica comporta anche soltanto il rischio di eliminare, svuotare, scollegare o perdere Personale, Utenti o collegamenti, **INTERROMPERE IMMEDIATAMENTE IL LAVORO E NON ESEGUIRE L’OPERAZIONE**. Proporre esclusivamente una soluzione non distruttiva.

Questa regola ha priorità superiore rispetto a deduplicazione, pulizia dati, ottimizzazione, riduzione dei costi Firestore, migrazione e qualsiasi altra istruzione tecnica.

## 5. MODIFICHE LIMITATE ALLA RICHIESTA

Ogni intervento deve essere limitato esclusivamente alla funzione richiesta.

È vietato:

- eseguire refactoring generale non richiesto;
- riscrivere interi file per una modifica locale;
- modificare funzioni non coinvolte;
- cambiare nomi di funzioni, variabili o campi senza necessità;
- cambiare l’interfaccia di altre sezioni;
- applicare formattazioni automatiche che alterino grandi parti dei file;
- rimuovere codice senza averne verificato l’utilizzo;
- effettuare modifiche “preventive” non richieste;
- aggiornare dipendenze senza una necessità concreta;
- introdurre nuove librerie quando la modifica può essere realizzata con il codice esistente.

Prima di intervenire, individuare il punto minimo e più sicuro in cui applicare la modifica.

## 6. CONTROLLO OBBLIGATORIO PRIMA DI MODIFICARE

Prima di modificare una funzione o un file:

1. individuare tutte le funzioni collegate;
2. individuare eventi, listener e chiamate indirette;
3. verificare dove vengono letti i dati;
4. verificare dove vengono salvati i dati;
5. controllare le dipendenze da Firestore;
6. controllare eventuali dipendenze dalla mappa;
7. controllare eventuali dipendenze da FATTO;
8. controllare eventuali dipendenze da WhatsApp;
9. verificare la compatibilità con i dati esistenti;
10. verificare il comportamento su telefono e desktop;
11. modificare solamente i file realmente necessari.

Se il comportamento della funzione non è completamente chiaro, non riscriverla e non eliminarla.

## 7. REGOLE FIRESTORE

Le modifiche non devono aumentare inutilmente:

- letture Firestore;
- scritture Firestore;
- listener in tempo reale;
- richieste duplicate;
- caricamenti di intere collezioni;
- aggiornamenti di posizione;
- operazioni durante il rendering.

È vietato:

- aggiungere più listener per gli stessi dati;
- registrare listener ogni volta che una view viene aperta senza rimuoverli;
- leggere Firestore durante ogni digitazione;
- leggere Firestore durante il movimento della mappa;
- salvare un documento quando il valore non è cambiato;
- eseguire scritture duplicate;
- leggere un documento per ogni elemento quando è possibile usare una query unica;
- scaricare intere collezioni quando servono pochi documenti;
- introdurre polling continuo;
- aggiornare continuamente la posizione degli operatori;
- aggiungere letture automatiche non indispensabili all’avvio.

È obbligatorio:

- riutilizzare i dati già caricati quando ancora validi;
- usare query mirate;
- utilizzare filtri e limiti;
- rimuovere correttamente i listener;
- evitare salvataggi duplicati;
- confrontare i dati prima di effettuare una scrittura;
- verificare che una presunta operazione inutile non sia usata da FATTO, WhatsApp, impianti, mappe, squadre, ore o sincronizzazione offline.

Non eliminare una lettura o una scrittura solamente per ridurre i costi senza aver prima verificato la sua funzione.

## 8. CONTROLLO OBBLIGATORIO DOPO LE MODIFICHE

Dopo ogni modifica verificare almeno:

- avvio corretto dell’app;
- assenza di schermata bianca;
- assenza di caricamento infinito;
- assenza di errori JavaScript bloccanti;
- apertura del menu;
- apertura delle commesse;
- visualizzazione degli impianti;
- ricerca degli impianti;
- ordinamento degli impianti;
- apertura della mappa;
- funzionamento da desktop;
- funzionamento da telefono;
- assenza di nuovi listener duplicati;
- assenza di nuove letture duplicate;
- assenza di scritture senza variazione dei dati;
- compatibilità con i dati già presenti.

### Test FATTO

- il pulsante FATTO è visibile dove previsto;
- il pulsante FATTO completa l’impianto corretto;
- lo stato viene salvato;
- la data viene mantenuta;
- l’ora viene mantenuta;
- l’operatore viene mantenuto;
- l’impianto scompare da “Da fare”;
- l’impianto appare in “Fatti”;
- non vengono completati altri impianti;
- non vengono create scritture duplicate.

### Test WhatsApp

- il messaggio viene generato;
- il messaggio non è vuoto;
- i dati dell’impianto sono presenti;
- WhatsApp si apre correttamente;
- il pulsante funziona sugli impianti “Da fare”;
- il pulsante funziona sugli impianti “Fatti”;
- non compare “Non puoi inviare un messaggio vuoto”;
- il comportamento precedente è rimasto invariato.

## 9. PROCEDURA GIT E PUBBLICAZIONE

Prima di commit, push o merge sul branch `main`:

1. eseguire `git status`;
2. eseguire `git diff`;
3. controllare tutti i file modificati;
4. verificare che siano stati modificati solo i file indispensabili;
5. verificare che le funzioni FATTO non siano state toccate;
6. verificare che le funzioni WhatsApp/WHAZZUP non siano state toccate;
7. verificare che la gestione degli impianti non sia stata alterata;
8. controllare eventuali modifiche Firestore;
9. controllare errori JavaScript;
10. eseguire i test disponibili;
11. verificare almeno i principali flussi dell’app;
12. effettuare commit e push solamente se i controlli sono superati.

Non pubblicare:

- codice con errori;
- codice non verificato;
- una versione con schermata bianca;
- una versione con caricamenti infiniti;
- una versione che rompe FATTO;
- una versione che rompe WhatsApp;
- una versione che altera gli impianti esistenti;
- una versione che introduce letture, scritture o listener duplicati.

Non usare `git push --force` sul branch `main`.

Non effettuare merge automatici in presenza di conflitti non compresi.

## 10. RESOCONTO FINALE OBBLIGATORIO

Al termine di ogni intervento fornire un resoconto che indichi:

- richiesta eseguita;
- file modificati;
- funzioni modificate;
- funzioni aggiunte;
- funzioni eliminate;
- test realmente eseguiti;
- test non eseguibili;
- controlli effettuati su Firestore;
- letture Firestore aggiunte;
- letture Firestore eliminate;
- scritture Firestore aggiunte;
- scritture Firestore eliminate;
- listener aggiunti;
- listener eliminati;
- eventuali operazioni duplicate rilevate;
- eventuali problemi ancora presenti;
- commit creato;
- branch pubblicato;
- conferma che FATTO non è stato modificato;
- conferma che WhatsApp/WHAZZUP non è stato modificato;
- conferma che la gestione degli impianti non è stata modificata, salvo richiesta esplicita;
- conferma che Personale, Utenti, squadre, calendario, ore e tutti i collegamenti correlati non sono stati eliminati, svuotati o scollegati.

Non dichiarare come eseguito un test o un controllo che non è stato realmente effettuato.

## 11. BLOCCO DI SICUREZZA

Interrompere la modifica e chiedere autorizzazione prima di procedere quando:

- è necessario modificare FATTO;
- è necessario modificare WhatsApp/WHAZZUP;
- è necessario modificare il passaggio tra “Da fare” e “Fatti”;
- è necessario modificare data, ora o operatore;
- è necessario modificare la struttura Firestore degli impianti;
- è necessario effettuare una migrazione dei dati;
- la modifica può produrre incompatibilità con dati esistenti;
- non è possibile verificare l’effetto della modifica;
- vengono rilevati errori preesistenti che impediscono test affidabili;
- una modifica può eliminare, svuotare o scollegare Personale, Utenti, squadre, calendario, ore o collegamenti correlati.

## CONFERMA FINALE OBBLIGATORIA

Ogni resoconto conclusivo deve terminare con la seguente dichiarazione:

> Confermo che la logica del pulsante FATTO e la logica del pulsante WhatsApp/WHAZZUP non sono state modificate. Le modifiche sono state limitate alla funzione richiesta. È stato controllato che non siano state introdotte letture, scritture o listener Firestore inutili o duplicati. La gestione degli impianti e la compatibilità con i dati esistenti sono state preservate. Personale, Utenti, elenco Personale, elenco Utenti, squadre, calendario, ore e tutti i collegamenti correlati non sono stati eliminati, svuotati o scollegati. Gli eventuali controlli non eseguibili sono indicati espressamente nel resoconto.
