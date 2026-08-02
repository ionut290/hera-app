# Inventario Firestore per la migrazione multi-organizzazione

## Stato attuale

Le regole correnti autorizzano numerose collezioni globali tramite `signedIn()` e `isAdmin()`. In questo stato, aggiungere soltanto un selettore organizzazione nell'interfaccia non garantirebbe l'isolamento dei dati.

La migrazione deve quindi introdurre il controllo organizzazione sia nelle query applicative sia nelle regole Firestore.

## Collezioni operative da rendere organizzative

Queste collezioni contengono dati che devono essere separati tra Varga, Levato e le future organizzazioni:

- `commesse` e relative sottocollezioni;
- `squadre`;
- `squadreCommesse`;
- `squadreStorico`;
- `globalCommesse`;
- `commessaResources`;
- `noteCommessa`;
- `oreReports`;
- `oreLocks`;
- `oreApprovalRequests`;
- `hoursDeadlineAlerts`;
- `programmazioni`;
- `calendarEvents` condivisi;
- `personale`;
- `mezzi`;
- `operatorCommessaHistory`;
- `documents` con visibilità aziendale;
- `posDocuments`;
- `notifications`, `appNotifications`, `userAlerts` e conferme relative;
- `chatMessages`;
- `safetyContacts`;
- `weatherAlerts` e rischi collegati agli impianti;
- tutte le collezioni del servizio neve;
- preventivi, consuntivi, prezziari e modelli eventualmente gestiti da moduli applicativi.

## Collezioni di piattaforma da mantenere globali

Queste collezioni devono restare a livello piattaforma, con accesso ristretto:

- `platformUsers`: identità globale dell'account;
- `organizations`: anagrafica delle organizzazioni;
- `organizationMemberships`: appartenenze e ruoli;
- `superAdmins`: elenco protetto dei Super Admin per UID;
- `appConfig`: configurazione generale, separando eventuali configurazioni specifiche per organizzazione;
- `userAccessAudit` e log di sicurezza della piattaforma;
- `registryImportBackups`, se rappresentano backup amministrativi globali.

## Modello dati proposto

```text
organizations/{organizationId}
  name
  status
  createdAt
  createdByUid
  branding

organizations/{organizationId}/members/{uid}
  role: owner | admin | operator
  status: active | suspended
  joinedAt

organizations/{organizationId}/commesse/{commessaId}
organizations/{organizationId}/squadre/{squadraId}
organizations/{organizationId}/oreReports/{reportId}
...

platformUsers/{uid}
  email
  displayName
  accountStatus
  defaultOrganizationId

superAdmins/{uid}
  active: true
```

## Compatibilità utenti attuali

Durante la fase transitoria:

1. l'organizzazione predefinita è `varga`;
2. tutti gli utenti attualmente approvati ricevono una membership attiva in Varga;
3. i ruoli attuali vengono conservati;
4. un utente appartenente soltanto a Varga entra direttamente senza selettore;
5. il selettore compare solo con almeno due membership attive;
6. nessun vecchio utente deve ricreare l'account.

## Strategia di migrazione sicura

### Fase A — Struttura non attiva

- creare moduli e documentazione senza collegarli all'app;
- nessuna nuova lettura, scrittura o listener;
- nessuna modifica alle regole distribuite.

### Fase B — Membership e organizzazione Varga

- creare `organizations/varga`;
- creare le membership degli utenti correnti con operazione amministrativa controllata;
- rendere la migrazione idempotente, evitando scritture quando il documento è già corretto;
- registrare il risultato in un audit amministrativo.

### Fase C — Lettura compatibile

- leggere prima il nuovo percorso organizzativo;
- durante la migrazione, usare il percorso storico solo per Varga e solo quando il nuovo dato non esiste;
- non eseguire entrambe le query in modo permanente;
- non creare listener duplicati sui percorsi vecchi e nuovi.

### Fase D — Scrittura controllata

- scrivere soltanto nel nuovo percorso organizzativo;
- bloccare la modifica se l'utente non ha una membership attiva;
- mantenere invariati i payload usati dalle funzioni FATTO e WhatsApp;
- migrare i dati storici prima di disabilitare i percorsi precedenti.

### Fase E — Regole definitive

- autorizzare l'accesso solo tramite membership;
- autorizzare il Super Admin tramite UID;
- impedire a un amministratore di creare organizzazioni o promuoversi a Super Admin;
- rimuovere l'accesso globale `signedIn()` alle collezioni organizzative.

## Controlli sui costi Firestore

La nuova architettura non deve introdurre:

- un listener aggiuntivo per determinare l'organizzazione su ogni schermata;
- doppie query permanenti su percorso storico e nuovo;
- letture dell'intera collezione membri per verificare il ruolo;
- scritture ripetute della membership a ogni accesso;
- aggiornamenti del documento organizzazione durante il rendering;
- polling per rilevare cambi organizzazione.

La membership dell'utente deve essere caricata una sola volta dopo l'autenticazione e riutilizzata in memoria finché valida.

## Protezioni FATTO e WhatsApp

Non devono essere rinominate, spostate, duplicate o riscritte le funzioni collegate a:

- completamento dell'impianto;
- salvataggio di data, ora e operatore;
- passaggio tra “Da fare” e “Fatti”;
- generazione e apertura del messaggio WhatsApp/WHAZZUP;
- WhatsApp degli impianti già completati.

La futura integrazione dovrà fornire a queste funzioni lo stesso documento impianto e lo stesso payload attuale, cambiando esclusivamente il percorso dati risolto a monte.

## Rischi rilevati

- `isAdmin()` è oggi basato anche sull'email: per il Super Admin deve essere usato l'UID;
- molte collezioni consentono lettura o scrittura a qualsiasi utente autenticato;
- `appConfig` contiene documenti leggibili e scrivibili da tutti gli utenti autenticati, salvo due eccezioni;
- il match generico finale permette accesso globale a diverse collezioni operative;
- documenti con visibilità `global` oggi significano globali per tutta la piattaforma, mentre in futuro dovranno significare globali solo nell'organizzazione.

## Stato del lavoro

Questo documento è solo un inventario tecnico. Non modifica le regole distribuite, non migra dati e non aggiunge operazioni Firestore.