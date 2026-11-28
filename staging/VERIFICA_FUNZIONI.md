# Verifica funzionale Varga Cantieri 2.0

## Obiettivo

Verificare ogni area già esistente, correggere i difetti senza cambiare l’impostazione dell’app e raccogliere evidenze ripetibili. Questa lista non introduce funzionalità nuove. Un’area è completa solo dopo test del flusso reale, salvataggio, permessi e regressione.

## Regole

- Usare solo staging, emulatori e dati di prova.
- Non pubblicare modifiche su produzione durante l’audit.
- Non cancellare o ricreare dati, account, ID o collegamenti.
- Non rifattorizzare FATTO, WhatsApp/WHAZZUP, gestione impianti, Personale/Utenti o storico.
- Verificare prima i prerequisiti e annotare errori, passaggi e log.

## Inventario delle aree da collaudare

| Area esistente | Percorso da verificare dall’inizio alla fine | Stato |
| --- | --- | --- |
| Avvio e login | Avvio a freddo, login Google/Apple previsti, sessione valida, logout e riapertura | Da collaudare su dispositivi |
| Registrazione azienda | Creazione, comparsa dell’azienda nell’account e accesso persistente | Da collaudare |
| Codice azienda e approvazioni | Richiesta di accesso, approvazione/rifiuto admin, ripresa del login e isolamento aziende | Test emulatori di isolamento verdi; percorso utente da collaudare |
| Ruoli e permessi | Admin, operatore e utente non autorizzato, su UI e Firestore | Isolamento emulatori verde; matrice completa ruoli da verificare |
| Piano prova e quote | 5 commesse, 100 impianti, 1 admin, 11 operatori, 11 mezzi; elemento successivo bloccato senza perdere dati | Integrato di staging verde; verificare ancora nel flusso UI |
| Commesse | Creazione, apertura, modifica, chiusura/archivio, filtri, ricerca e visibilità coerente | Test di staging verdi; collaudo completo UI da fare |
| Importazione commesse e impianti | Excel, righe mancanti/duplicate, colonne alternative, errori e ripetizione import | Test unitari disponibili; collaudo con file campione da fare |
| Impianti | Inserimento manuale, modifica, ricerca, scheda, collegamento commessa e mappa | Test staging di creazione/regole verdi; percorso completo da fare |
| FATTO e stato Da fare/Fatti | Conferma, data/ora/operatore/note, riapertura, sincronizzazione e storico | Test di regressione dedicati verdi; prova su iPhone/Android da fare |
| WhatsApp/WHAZZUP | Testo corretto e completo, app installata, foto/documenti e ritorno all’app | Funzione protetta; prova nativa su dispositivi da fare |
| Mappa, GPS e navigazione | Permessi, posizione, impianto selezionato, route, ritorno in app e permesso negato | Da collaudare su dispositivi reali |
| Squadre | Creazione/modifica, data, operatori, assegnazioni, storico e ricaricamento | Test mirati presenti; collaudo end-to-end da fare |
| Personale e utenti | Inserimento/modifica, collegamenti, ruoli, inattivazione e conservazione dello storico | Regole di non cancellazione applicate; verifica dati reali esclusa, test isolato da completare |
| Mezzi | Inserimento/modifica, ID/targa, disponibilità e collegamenti alle squadre/commesse | Da collaudare end-to-end |
| Ore e contabilità | Inserimento, controllo duplicati, approvazione, periodo, esportazione e riapertura dati | Moduli/test presenti; percorso completo da fare |
| Documenti, foto e archivi | Caricamento, associazione corretta, apertura/download, permessi, quota e retry | Test quota/storage di staging verdi; test dispositivi e file grandi da fare |
| Chat e segnalazioni | Invio/ricezione, allegati, conferma, visibilità per azienda e riapertura | Da collaudare end-to-end |
| Calendario e notifiche | Creazione/aggiornamento, fuso orario, destinatari, permessi e app chiusa | Test codice parziali; notifiche native da provare su dispositivi |
| Servizio Neve | Commesse, KML, tratti, modifica/reset giro, offline e sincronizzazione | Test e materiali specifici presenti; ripetere collaudo operativo su device |
| Potature/Abbattimenti, alberi e COBO | QR, apertura scheda, lavoro, foto, stati e report | Moduli e test dedicati presenti; prova end-to-end su device da fare |
| Preventivi, prezziari e documentazione cantiere | Import prezzi, calcolo, salvataggio, PDF, allegati e riapertura | Test unitari di import/proiezione disponibili; percorso completo da fare |
| Carburanti e raccomandazioni | Ricerca, filtri, posizione, risultati vuoti e rete assente | Da collaudare |
| Offline e sincronizzazione | Operazioni offline, ripristino rete, retry, conflitti e assenza di duplicati | Coda dedicata per alcuni flussi; verifica uniforme in ogni area da fare |
| Backup e ripristino | Integrità, conteggi, ID e relazioni prima/dopo, prova di ripristino isolata | Documentazione/workflow presenti; esercitazione isolata da fare |
| Errori e supporto | Errore comprensibile, registrazione diagnostica, retry e privacy dei log | Da collaudare |
| Aggiornamento app e compatibilità | Aggiornamento senza logout involontario o perdita dati; versione web e native | Test mirati presenti; matrice versioni da completare |
| iPhone/iPad | Login, layout safe-area, fotocamera, mappe, notifiche, offline e navigazione | Test browser automatizzati verdi; prova su dispositivi reali ancora necessaria |
| Android | Login, pulsanti fuori dalle barre di sistema, WhatsApp nativo, GPS/geofence, aggiornamenti | Test automatici mirati presenti; prova su dispositivi reali ancora necessaria |

## Evidenze già controllate su staging

Commit verificato: `e52d9e60e1566397bdd1f5b0a74670fda52eb21d`.

- Suite Node: 265 test superati, 0 falliti.
- Browser automation: preview e full runtime superati.
- Build del portale staging superata.
- Emulatori Firestore: test d’isolamento tenant superati; pacchetto delle regole testate creato per revisione.
- Test integrato: pacchetto, quote/storage e billing sandbox superati.
- Hosted checkout non eseguito (skipped); non è quindi certificato da questo risultato.
- La verifica automatica del runtime segnala moduli e collezioni candidati a revisione: tale inventario non prova da solo che ogni modulo sia autorizzato o verificato.
- I test CI non sostituiscono la prova pratica su iPhone/Android né attestano il comportamento di ogni funzione.

## Chiusura dell’audit

Per ogni riga registrare: dispositivo/ambiente, account e ruolo di prova, passi eseguiti, risultato, screenshot/log, difetto e commit della correzione. Dopo ogni fix ripetere il test dell’area e le regressioni che possono essere coinvolte.

La verifica è conclusa solo quando ogni area è marcata **Superata**, oppure è indicata chiaramente come bloccata da accesso, credenziali o prova su dispositivo. Un’area non provata non va dichiarata funzionante.

## Ultima simulazione UI

Il collaudo integrato esercita anche le schermate Personale v2, Mezzi v2 e Gestione commesse tramite il browser autenticato.
