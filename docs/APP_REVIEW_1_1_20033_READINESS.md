# VARGA CANTIERI 1.1 — candidato 20033, verifica 26 novembre 2026

## Stato: BLOCKED — nessuna approvazione pubblica
- App originale: Apple ID 6833571061. V2 e produzione non devono essere modificate da questo rilascio.
- 1.1 (20032): upload interno riuscito, Apple ha confermato disponibilità TestFlight, proprietario riferisce 8/8 prove operative Piano Neve riuscite. CI della 20032 ha superato 49 verifiche routing/asset. Non sono prove di accesso reviewer, cancellazione account o revoca Apple.
- 20032 esportata con testFlightInternalTestingOnly=True: non associabile a una versione pubblica.
- Candidato 20033: workflow manuale ios-appstore-candidate.yml preparato per 1.1, mantiene i controlli del Piano Neve e il gate --require-ready; nessuna compilazione/upload 20033 confermata.
- Non rimuovere gate, non dichiarare conformità basandosi sul collaudo del Piano Neve o sulla precedente approvazione 1.0.1.

## Evidenza Firebase (sola lettura, 26 novembre)
Console progetto hera-app-6cd2b: deleteOwnVargaAccount visibile tra le funzioni distribuite; requestOwnVargaAccountDeletion, getOwnVargaAccountDeletionStatus e specifiche funzioni di revisione cancellazione non visibili. Questo non prova che una funzione distribuita sia utilizzabile da tutti gli utenti. Nessuna invocazione o modifica eseguita.
Nessun progetto dedicato di sandbox App Review trovato nell'elenco accessibile. V2 STAGING è un progetto distinto e non va usato automaticamente come sandbox per V1.1. Stato account canary non verificato.

## Regola permanente screenshot (richiesta esplicita del proprietario)
Solo catture autentiche dell'app con dati sintetici. Nessun nome reale (nemmeno quello del proprietario), mezzo reale, targa, matricola, contatto, identificativo dipendente o cliente, coordinata riservata, documento reale. Usare etichette neutre Operatore Demo, Squadra Demo, Mezzo Demo, Commessa Demo, non copiare il database aziendale. Verifica visiva di ogni immagine a piena risoluzione e di ogni localizzazione e dimensione iPhone/iPad prima dell'assegnazione; escludere immagini incerte. Non usare mockup che simulano funzioni non esistenti. Non riattivare pagine reviewer-only rimosse dopo il rifiuto 2.5.19.

## Condizioni ancora aperte
1. Ambiente di prova isolato per il normale login, con dati sintetici e tutti i principali flussi accessibili nel rispetto dei ruoli. Test emulatori non equivalenti a prova su iPhone.
2. Approvazione documentata del titolare competente e revisore indipendente per deployment privacy protetto, retention e procedura effettiva di gestione richieste; evitare modifiche a utenti reali e non inventare attestazioni o SLA.
3. Collaudo della cancellazione/richiesta, isolamento per UID e revoca Apple con utenze sintetiche autorizzate; mantenere protezioni di produzione.
4. App Review notes veritiere, credenziali demo isolate e attive per la revisione (mai nel repository); screenshot puliti; conferma dei requisiti privacy e classificazione in ASC.
5. Nuovo upload candidato solo dopo esito positivo del gate; verificare numero build ancora libero, compilazione, firma, installazione e regressione della 20033 prima dell'invio. Nessun invio o pubblicazione automatica.

Riferimenti: APP_REVIEW_FULL_ACCESS_PLAN_2026-11-20.md, ACCOUNT_DELETION_REQUEST_DEPLOY_GATE_2026-11-20.md, SOLO_OWNER_RELEASE_PROTOCOL_2026-11-20.md, IOS-APPSTORE-2519-REMEDIATION.md. Il presente documento non sostituisce le approvazioni privacy, non modifica la logica dell'audit e non autorizza distribuzioni Firebase.
