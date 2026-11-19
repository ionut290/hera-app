# VARGA CANTIERI — Proposta di decisione sulla conservazione e cancellazione dei dati

**Stato:** CRITERI DI CONSERVAZIONE APPROVATI DAL REFERENTE IL 19 NOVEMBRE 2026. **Implementazione tecnica, verifica dei ruoli GDPR e collaudo in produzione ancora da completare.**
**Responsabile della decisione dichiarato nel dialogo:** gestore autorizzato di VARGA CANTIERI. La dichiarazione di poter approvare le regole dell'app **non prova automaticamente** di essere il titolare GDPR dei dati delle aziende o dei loro dipendenti: i ruoli e gli accordi con eventuali aziende utilizzatrici devono essere verificati.
**Ambito:** Firebase Authentication, Firestore, Cloud Storage, Google Drive/Sheets, esportazioni e backup. Non costituisce un parere legale né esegue cancellazioni.
**Fonti esaminate:** `privacy.html` del repository (ultimo aggiornamento 11 ottobre 2026), `docs/account-deletion-data-inventory-2026-11-19.md`, codice `functions/account-deletion.js`, test Firebase Auth/Firestore Emulator, norme GDPR art. 5 e 17 e indicazioni del Garante per i dati GPS dei lavoratori.

## A. Decisioni proposte — DA APPROVARE (non sono policy oggi implementate)

| Archivio / dati | Proposta di cancellazione o conservazione | Controllo prima di attivarla |
| --- | --- | --- |
| Account Firebase Auth, `platformUsers`, documenti strettamente privati | Su richiesta confermata con accesso recente: cancellare credenziali di accesso, profilo e archivi privati coperti dalla procedura, senza aspettare i 12 mesi menzionati per la sola disattivazione ordinaria. | Distinguere cancellazione volontaria da disattivazione; assicurare revoca login Apple e nessun record privato residuo senza base giuridica. |
| `personale`, ore, permessi, contratti, rapporti e dati di cantieri | Conservare **solamente** le parti necessarie agli obblighi/documentazione aziendali, senza mantenerle collegate al Firebase UID eliminato. Il termine **non è automaticamente 10 anni per tutti i dati**: 10 anni solo dove la finalità amministrativo-contabile lo giustifica, altre categorie da definire con il titolare competente. | Confermare titolarità delle aziende, singoli obblighi, categorie particolari (malattia/infortunio), destinatari e tempi dei registri. |
| `chatMessages` | Preservare i messaggi operativi indispensabili a una commessa ancora in corso; **proposta indicativa:** riesame e rimozione/anonimizzazione entro 12 mesi dalla chiusura del lavoro, salvo retention documentata. Quando l'account viene eliminato, rimuovere email/UID dell'autore e destinatario quando non necessari; non alterare contenuto o conversazioni di altri utenti indiscriminatamente. | Verificare se messaggi privati e chat aziendali sono distinguibili, regole di accesso e riferimenti incrociati; test di anonimizzazione con dati fittizi. |
| `notifications`, `appNotifications` | Rimuovere le preferenze e gli stati personali di lettura dell'account eliminato. **Proposta:** 90 giorni per notifiche transitorie; conservare solo avvisi che costituiscono effettivo documento di lavoro finché richiesto dalla relativa finalità. | I testi e `metadata` possono ancora citare persone: ispezione limitata e test senza perdita di comunicazioni per gli altri utenti. |
| `notificationEmailQueue` | Non continuare a recapitare notifiche all'indirizzo dell'account cancellato salvo obbligo specifico. **Proposta:** rimuovere code concluse/errore dopo 30 giorni e gestire le code pendenti con sospensione o rimozione selettiva. Non si possono cancellare email già recapitate a servizi esterni. | Verificare schema/stati reali della coda e impatto sui destinatari aziendali. |
| `activityLogs` e log tecnici | Applicare **90 giorni**, come già previsto (ma non automatizzato) nell'informativa, salvo singola eccezione documentata per sicurezza o obblighi. Minimizzare `userId`, `userEmail`, `deviceInfo` dopo la cancellazione. | Implementare TTL/pulizia controllata, verificare che i log necessari alla sicurezza restino disponibili entro la scadenza. |
| `gpsUpdateRequests` e coordinate operatore | Conservare coordinate precise **solo durante la verifica della richiesta**; **proposta prudenziale:** eliminare/minimizzare le coordinate entro 30 giorni dalla chiusura della richiesta, o prima quando non più necessarie. Conservare soltanto esito/luogo operativo quando giustificato. | Valutazione lavoristica e GDPR su geolocalizzazione, base giuridica, informativa ai dipendenti, periodo di raccolta, permessi, accessi e tracciamenti fuori servizio. I 30 giorni NON sono un termine legale generale. |
| Foto e allegati operativi | Applicare il termine programmato di **24 mesi dalla chiusura del lavoro**, salvo giustificazione documentata di altro termine; documenti privati senza finalità aziendale vanno rimossi alla cancellazione dell'account. | Distinguerli dai documenti contabili soggetti ad altri termini, verificare riferimenti condivisi. |
| Fatture, documenti amministrativi e registrazioni di lavorazione | **Fino a 10 anni** solo nei casi in cui sia necessario per legge o altra finalità giustificata, come descritto nell'informativa attuale. Non estendere indiscriminatamente questo periodo a chat/GPS/log. | Elenco categorie e basi giuridiche da approvare separatamente. |
| Backup Firestore | Conservare nel ciclo di **98 giorni** attualmente configurato; non tentare cancellazioni manuali indiscriminate. In caso di restore, applicare un elenco affidabile di utenti già eliminati e verificare che non siano ripristinati account o accessi. | Formalizzare la procedura di riconciliazione dei dati ripristinati e gli accessi ai backup. |
| Esportazioni Google Drive/Sheets, Storage legacy, altri fornitori | Mappare copie e responsabili, applicare gli stessi criteri per quanto tecnicamente e contrattualmente possibile. La cancellazione Firebase **non** elimina automaticamente queste copie. | Inventario esportazioni e tempi dei servizi esterni, procedura manuale per richieste aggiuntive. |

## B. Informativa e comunicazione dell'utente

- Correggere la possibile ambiguità dell'informativa tra **disattivazione ordinaria fino a 12 mesi** e **richiesta esplicita di cancellazione**, che non deve mantenere l'identità di autenticazione per 12 mesi senza giustificazione.
- Indicare esplicitamente quali tipi di dati possono rimanere nei registri aziendali, per quale finalità e quale periodo o criterio di conservazione.
- Identificare chi è il titolare GDPR per i dati di aziende/dipendenti; il gestore dell'app non diventa automaticamente titolare unico per tutti i dati.
- Rendere chiara la conservazione delle copie di sicurezza fino alla scadenza e le modalità di esercizio dei diritti.
- **Non modificare automaticamente la privacy.html pubblica prima dell'approvazione e delle relative implementazioni effettive.** Non dichiarare operativi automatismi non ancora distribuiti.

## C. Verifiche e rilascio

- [x] Login Apple su iPhone con beta interna 1.0 (10003), confermato dal proprietario.
- [x] Pulsante `Elimina il mio account` visibile su iPhone.
- [x] Test reali di cancellazione su Firebase Auth e Firestore **emulatori**; due utenti sintetici isolati.
- [x] CI PR #1606: 20/20 verdi alla data dell'ultimo controllo.
- [x] Environment GitHub `firebase-account-deletion-production` configurato: reviewer obbligatorio `ionut290`, solo `main`, bypass disattivato.
- [x] **Approvazione esplicita dei criteri di questa proposta:** l'utente, dopo essersi dichiarato autorizzato a prendere le decisioni per VARGA CANTIERI, ha risposto «Approvo» il 19 novembre 2026. Nessuna modifica ai periodi proposti è stata richiesta.
- [ ] Verifica dei ruoli GDPR con le aziende e coerenza delle basi giuridiche.
- [ ] Implementazione e test delle nuove minimizzazioni/retention approvate nei dati reali, oppure processo manuale temporaneo documentato e conforme.
- [ ] Prova di cancellazione end-to-end con account dimostrativo **in produzione**, non collegato a personale o dati aziendali.
- [ ] Video iPhone/iPad richiesto da Apple e nuova presentazione App Review.

**Regola di sicurezza:** i criteri del punto A sono approvati, ma i controlli e le implementazioni del punto C restano incompleti. Lasciare PR #1606 in bozza e non distribuire `deleteOwnVargaAccount` ai lavoratori prima di una verifica dei ruoli GDPR, del collaudo e dell'approvazione di produzione.

## D. Decisione del referente (registrata)

**Esito:** APPROVAZIONE DEI CRITERI DI CONSERVAZIONE PROPOSTI (non autorizzazione alla distribuzione della funzione).

**Approvazione:** il referente di VARGA CANTIERI, che ha confermato di avere il potere di decidere sul prodotto, ha risposto **«Approvo»** in chat il **19 novembre 2026**. Non sono state comunicate modifiche ai periodi.

**Limiti:** l'approvazione non certifica la base giuridica di ogni trattamento, la posizione di titolare per dati di altre aziende, l'esecuzione di pulizie automatiche o il completamento della prova di eliminazione su iPhone. Questi punti restano aperti nel ticket #1609.

## Riferimenti ufficiali

- GDPR, Regolamento UE 2016/679, artt. 5, 13 e 17: https://eur-lex.europa.eu/eli/reg/2016/679/oj
- EDPB, FAQ sul diritto alla cancellazione: https://www.edpb.europa.eu/contact/frequently-asked-questions_mt?search=2025
- Garante Privacy, provvedimento sulla geolocalizzazione dei lavoratori: https://www.garanteprivacy.it/web/guest/home/docweb/-/docweb-display/docweb/5217175
