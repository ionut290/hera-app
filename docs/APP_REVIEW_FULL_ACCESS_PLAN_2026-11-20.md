# VARGA CANTIERI — accesso COMPLETO per collaudo e revisione Apple

**Stato 20 novembre 2026:** NON ancora pronto per dichiarare "accesso completo". Questa scheda descrive i requisiti da implementare e verificare, non certifica funzionalità già operative.

## Decisione aggiornata del proprietario — dati REALI

**Richiesta esplicita del 20 novembre 2026:** il proprietario preferisce dare ai collaudatori accesso ai **dati operativi reali** di VARGA CANTIERI, anziché usare soltanto dati inventati. Il piano di sandbox interamente sintetica più sotto è dunque una soluzione alternativa per dati non autorizzabili, NON la preferenza del proprietario. Non confondere "accesso completo alle funzioni" con un permesso illimitato su dati personali di terzi.

**Fase richiesta: account di collaudo su backend di produzione, limitato per ruolo e autorizzazione.** Prima di attivarlo:
1. Il titolare o il responsabile autorizzato del trattamento deve confermare quali aziende, commesse, clienti e categorie di dati personali possono essere viste da un revisore esterno, e su quale base giuridica/istruzione documentata. La sola disponibilità delle credenziali di un amministratore tecnico non autorizza a condividere dati GPS, email personali, documenti d'identità, stipendi, informazioni sanitarie, ore e turni dei lavoratori.
2. Le **commesse, posizioni degli impianti e dati operativi reali autorizzati** possono essere visualizzati dal collaudatore in sola lettura, con controllo server-side per raccolte, documenti e campi autorizzati. Non bastano controlli CSS/interfaccia e non va mai disabilitato `isAppleReviewAccount` alla cieca.
3. Per provare creazione/modifica/cancellazione, usare una sezione o record **esplicitamente marcati come test** con rollback e tracciabilità, evitando alterazioni su interventi reali, orari, notifiche o assegnazioni. Se il proprietario vuole la modifica dei dati reali, richiede una procedura separata con supervisione e approvazione specifica del titolare.
4. Dati identificativi di lavoratori (nomi, email, telefonini), coordinate GPS individuali, timbrature, note mediche, documenti aziendali riservati e dati di clienti **non sono automaticamente inclusi** nell'autorizzazione. Per ciascuna categoria occorre verifica di necessità, autorizzazione e controlli di accesso. Preferire redazione/mascheramento quando basta a provare l'interfaccia.
5. Account temporaneo dedicato a revisione: minimo privilegio, eventuale scadenza **solo dopo** completamento della review, log degli accessi, possibile revoca immediata, nessun ruolo `admin`, nessuna password pubblicata in GitHub. Le credenziali devono restare valide per tutta la revisione Apple.
6. Il backend attuale blocca i reviewer dalle raccolte operative (`signedIn()` esclude `apple_review`) e dal Storage: **NON È ANCORA STATO ABILITATO NESSUN ACCESSO AI DATI REALI**. Aggiungere prima regole Firebase restrittive e test su emulatori e account reale controllato; nessun rilascio automatico per questa richiesta.
7. Il percorso di cancellazione account resta obbligatorio nell'app ordinaria, non soltanto nella demo. Le cautele di GDPR/retention ancora pendenti non vanno omesse.

**Questa variante con dati reali è subordinata all'autorizzazione pertinente e ai test di sicurezza**. Non inviare ad Apple un account privilegiato che mostra informazioni riservate senza autorizzazione. Apple chiede accesso funzionale completo, non accesso indiscriminato ai dati di dipendenti/clienti.

## Requisito Apple

App Review Guideline 2.1: i revisori devono poter usare tutte le principali funzioni attraverso un **account attivo** oppure una **modalità dimostrativa completa**, con dati di test e servizi backend disponibili. Le credenziali devono funzionare durante la revisione. Documentazione: https://developer.apple.com/app-store/review/guidelines/ e https://developer.apple.com/help/app-review/before-submitting-for-review/complete-review/

## Stato reale del repository

1. `firestore.rules` esclude `apple_review=true` dalla funzione generale `signedIn()`: perciò la precedente identità review non può aprire `commesse`, `impianti`, `oreReports`, `personale` e le altre raccolte operative; permette soltanto `appleReviewDemo` in lettura.
2. `storage.rules` proibisce al reviewer l'accesso agli allegati ordinari.
3. `scripts/prepare-capacitor-web.js` esclude intenzionalmente i file `apple-review-*` dai bundle nativi e `check-appstore-2519-clean.js` impedisce logiche esclusive al revisore: **non rimuovere** queste protezioni.
4. `deleteOwnVargaAccount` attualmente è un canary **solo demo**, verificato nel backend ma non pronto per la cancellazione generalizzata degli account dei lavoratori.
5. Il vecchio account `sqadrahera+varga-demo@gmail.com` è stato **eliminato** nel test reale GitHub #44620065125; non esiste più come account Auth. NON indicare le vecchie credenziali ad Apple.
6. Per il momento manca prova di un test di tutte le funzioni su iPhone fisico con dati fittizi.

## Architettura consigliata

**Alternativa protettiva quando l'accesso ai dati reali non è autorizzato:** ambiente di prova con progetto Firebase separato da `hera-app-6cd2b`, dati completamente sintetici e account di prova attivo per tutta la finestra App Review. Non aggiungere un amministratore Apple nel database aziendale e non disabilitare `isAppleReviewAccount`/regole Storage per avere un risultato apparentemente "verde".

- Distribuire una versione di collaudo **dell'app vera**, con le stesse schermate e azioni degli utenti normali: accesso, elenco commesse, mappa/impianti, squadre, timbrature, calendario, documenti, contabilità/preventivi/consuntivi, notifiche e cancellazione account.
- Dati modello: almeno due commesse finte, tre impianti georeferenziati con località chiaramente inventate/non riconducibili a commesse reali, due lavoratori fittizi, un turno e una timbratura, un calendario, un PDF fittizio e un consuntivo con una voce di prezzario. Non copiare dati di dipendenti, geolocalizzazioni personali, firme, numeri o email reali.
- Distinguere account con permessi **utente** e **amministratore** all'interno della sandbox; se l'app ha schermate distinte Apple raccomanda credenziali per entrambi i tipi. Ogni ruolo deve mostrare le funzioni permesse per quel ruolo. Nessun ruolo sandbox deve avere accesso al Firebase aziendale.
- Tutti i moduli dovranno essere provati anche nelle operazioni di scrittura/lettura: creazione di test, modifica di test e cancellazione dei propri dati demo. Per funzioni esterne (GPS, notifiche iOS e mail) usare simulazioni sicure o numeri/account di prova, spiegando in App Review notes eventuali limiti geografici/hardware.
- La modalità demo non deve contenere percorsi nascosti esclusivi per la sola revisione, credenziali hard-coded nel client o aggiramenti di controlli. Una modalità esercitazione pubblicamente e normalmente disponibile agli utenti, se scelta, deve essere autentica e completa.
- L'app finale distribuita all'utente deve poter fare la stessa cancellazione account conforme a legge; una sola demo cancellabile **non** soddisfa Apple Guideline 5.1.1(v).

## Matrice di collaudo obbligatoria

| Area | Prova da superare | Evidenza |
|---|---|---|
| Login | Credenziali demo funzionanti, Apple Sign in separatamente verificato | Video iPhone |
| Home | Accesso a sezioni senza pagina vuota o "in attesa" | Video e screenshot |
| Commesse | Aprire, filtrare, creare e modificare **record di prova** | Log e screenshot sandbox |
| Impianti/mappa | Aprire punti demo, navigazione/mappa senza coordinate reali riservate | Video |
| Squadre e personale | Assegnazioni fra identità finte secondo il ruolo | Screenshot sandbox |
| Timbrature | Inizio/fine turno di prova, cronologia, notifiche e limiti GPS | Video su iPhone fisico |
| Calendario | Aprire e salvare evento di prova | Screenshot sandbox |
| Documenti | Caricare, aprire e cancellare documento sintetico | Log sandbox |
| Contabilità | Inserire voce prezzario, consuntivo e calcolare totale correttamente | Dati/evidenze sintetiche |
| Notifiche | Ricevere o simulare notifica autorizzata su device | Video o test |
| Profilo e account | Opzione eliminazione facilmente trovabile e prova di eliminazione dell'account fittizio | Video iPhone e log server |
| Ripristino | Utente diverso non influenzato dalla cancellazione del demo | Verifica server |

## Gate di rilascio

- Eseguire `node scripts/audit-app-review-full-access.cjs`. Il report odierno indica **BLOCKED**; NON trasformarlo in successo fittizio.
- Per validare davvero il progetto: realizzare la sandbox e i suoi permessi backend, un login temporaneo verificato su iPhone, test E2E per ogni area, prove di isolamento, video.
- Solo quando **tutti** i test sono superati, rendere disponibili credenziali demo funzionanti in App Store Connect e scrivere le note Apple.
- GDPR: prima del rilascio generale completare inventario dei dati, ruoli aziendali, conservazione di email/log/backup Google Drive e comportamento della cancellazione; non assumere che test su demo equivalga a conformità.
