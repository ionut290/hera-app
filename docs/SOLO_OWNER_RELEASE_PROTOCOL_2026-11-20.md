# VARGA CANTIERI — gestione sicura da amministratore unico
**20 novembre 2026 — procedura preparatoria, NON autorizzazione privacy o App Store**

## Obiettivo
Il proprietario GitHub `ionut290` può verificare e rilasciare **solo le Cloud Functions ordinarie già presenti nella lista approvata**, senza chiedere a un secondo utente di approvare ogni attività. Nessuna modifica in `main` provoca più automaticamente una distribuzione Firebase da questo workflow.

Questa procedura **non** abilita la cancellazione degli account aziendali, non modifica le regole Firestore di produzione, non cambia le autorizzazioni dei lavoratori e non sostituisce l'approvazione privacy di un titolare del trattamento competente. È compatibile con un amministratore unico per i lavori ordinari, con protezioni più rigorose per operazioni ad alto rischio.

## Per eseguire una prova in completa autonomia (senza modificare produzione)
1. Aprire GitHub > Actions > **Deploy Firebase Functions** > **Run workflow** sul branch `main`.
2. Scegliere `PREVIEW_ONLY` (valore predefinito), inserire il **SHA esatto** da 40 caratteri del commit `main`, e lasciare `DO_NOT_DEPLOY`.
3. Il job `guard` con permessi di sola lettura verifica: repository e proprietario, lancio manuale su `main`, SHA esatto, assenza di modifiche di rilascio per cancellazione, configurazione dei workflow protetti, controlli privacy e sicurezza, stato App Review.
4. L'eventuale job `deploy` non parte: non legge segreti Firebase e non modifica account, regole, storage, aziende, Apple o IAM.

## Rilascio ordinario (solo se la modifica è effettivamente autorizzata)
1. Verificare il commit `main`, tutti i test pertinenti, eventuali impatti su dati e accessi, e preparare una procedura di ripristino per le modifiche non reversibili.
2. Avviare **un nuovo workflow manuale** da `main`; usare `DEPLOY_ORDINARY_ONLY`, SHA esatto corrente e `CONFIRM_ORDINARY_FIREBASE_DEPLOY`.
3. GitHub accetta il rilascio solo dal proprietario `ionut290`, su `main`, con il SHA approvato e durante **il primo tentativo**. Il riavvio del job di distribuzione già approvato viene rifiutato: avviare un nuovo workflow per rinnovare la conferma.
4. I controlli `guard` passano **prima** che il job `deploy` possa leggere il segreto Firebase. Il job distribuisce soltanto i nomi delle funzioni nella lista esplicita, con CLI fissata alla versione prevista.
5. **Nessuna regola Firestore** viene distribuita dal percorso solo-owner; nessuna policy IAM viene ampliata automaticamente. Se una funzione dipende da nuovi permessi, il rilascio ordinario deve essere fermato e valutato con la procedura appropriata.
6. Registrare il commit, l'esito e l'eventuale problema nel ticket pertinente. Non pubblicare su App Store per il solo fatto che questo workflow sia verde.

## Operazioni che NON si possono sbloccare semplicemente con un proprio secondo account
- `deleteOwnVargaAccount` (cancellazione reale/disponibile solo a demo), `requestOwnVargaAccountDeletion` e stato richieste, e le custom claims della demo di cancellazione hanno workflow dedicati protetti da `firebase-account-deletion-production` con **Prevent self-review ON**. Nessun bypass o secondo account fittizio.
- Approvazione dei tempi di conservazione, esportazioni, backup, log, geolocalizzazione e adempimento richieste utenti richiede valutazione documentata dei reali titolari dei dati; questo è diverso dall'approvazione tecnica GitHub. Se il proprietario è anche l'effettivo titolare per certi dati, va verificata la sua competenza e documentata l'approvazione; non inventarla.
- Le regole Firestore e Storage possono esporre dati dei lavoratori anche se la modifica appare piccola: non fanno parte del percorso ordinario.
- L'accesso dei revisori Apple deve usare soltanto dati demo sintetici. Il vecchio account ASC è attivo ma **non è certificato come reviewer isolato**. Le credenziali non vanno inserite nei ticket GitHub.

## Stato operativo quando questa procedura è stata scritta
- TestFlight `1.0 (20015)` era l'ultima build confermata; la versione pubblica `1.0 (21)` rimaneva rifiutata.
- I test CI/emulatori non sostituiscono il collaudo reale della cancellazione e della revoca Apple su iPhone.
- Questo protocollo **non avvia da solo alcuna distribuzione**, non modifica l'ambiente protetto delle cancellazioni e non equivale a una release di produzione.
- Se la protezione indipendente è impossibile con un solo amministratore, i workflow sensibili rimangono bloccati. L'eventuale alternativa futura richiede una valutazione specifica di rischio, ruoli privacy, meccanismi di doppia conferma, audit e recupero: non abbassare il livello di sicurezza tramite questo workflow.

Riferimenti: issue [#1609](https://github.com/ionut290/hera-app/issues/1609), [#1666](https://github.com/ionut290/hera-app/issues/1666), [#1672](https://github.com/ionut290/hera-app/issues/1672).


## Nuovo percorso TestFlight interno dell'app iPhone (non pubblico)

Il workflow `.github/workflows/ios-internal-testflight-main.yml` consente di preparare una build firmata del codice `main` per il **TestFlight interno**, anche quando l'audit per App Review pubblica rimane `BLOCKED`.

1. Verificare la build numericamente più recente su App Store Connect. **Aggiornamento 21/11/2026:** `1.0 (20016)` è già in Testing interno; il prossimo valore *proposto*, solo se ancora libero, è **`20017`**. Il workflow rifiuta i numeri ≤ 20016.
2. Verificare i controlli CI del commit `main` e assicurarsi che nessun dato/utente reale sia esposto. Non sostituire il workflow pubblico o eliminare il suo gate.
3. Da GitHub Actions > **Varga Cantieri internal TestFlight from reviewed main**, scegliere branch `main`, opzione `UPLOAD_INTERNAL_TESTFLIGHT_ONLY`, numero build non usato e SHA esatto del commit `main`.
4. Il workflow verifica due volte l'identità `ionut290`, il primo tentativo e il SHA. Richiede configurazione Apple privata nel preesistente environment `ios-internal-testflight`; genera l'app con login Apple e Google, crea un archivio firmato e carica su **App Store Connect per TestFlight**, non in App Review.
5. Dopo l'elaborazione Apple, installare la build su iPhone e controllare login Google/Apple, persistenza, menù, revoca Apple e cancellazione account *soltanto su utenza demo isolata autorizzata*. Se le funzioni privacy sono ancora disattivate, registrarlo come test NON superato, senza effettuare cancellazioni reali.
6. La build TestFlight non equivale a pubblicazione; `ios-appstore-candidate.yml` rimane separato e blocca l'upload pubblico tramite `audit-app-review-full-access.cjs --require-ready`.

**Nessuna funzione Firebase o regola Firestore viene distribuita dal workflow TestFlight.** I servizi backend restano quelli già approvati; il codice del client può essere più recente di Firebase produzione. Nessun accesso aziendale nuovo, nessun account o permesso modificato dalla sola compilazione. La capacità di inviare alla revisione Apple dipende ancora dalle prove privacy e reviewer.
