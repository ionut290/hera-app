# Centro errori amministratore

## Dove si trova

- Amministratore: `Menu → Altri strumenti → ⚠️ Centro errori`.
- Tutti gli utenti autenticati: `Menu → Altri strumenti → 🐞 Segnala problema app`.

## Rilevamento automatico

`app-error-monitor.js` registra soltanto segnali tecnici:

- errori JavaScript e Promise non gestite;
- risorse locali non caricate;
- operazioni lunghe sul thread principale, quando il browser supporta `PerformanceObserver`;
- ritardo dell’interfaccia subito dopo un tocco;
- tre tocchi ravvicinati sullo stesso comando;
- contesto tecnico della pagina e ultime azioni, senza valori digitati nei moduli.

Non vengono acquisiti posizione GPS, fotografie, contenuti dei documenti, password, PIN, token, cookie o chiavi API. I valori sensibili eventualmente presenti in un messaggio tecnico vengono oscurati.

## Coda offline

Gli eventi non inviabili vengono salvati nella chiave locale:

`hera_error_center_queue_v1`

La coda contiene al massimo 30 eventi, conserva gli elementi per 7 giorni e viene svuotata quando tornano connessione e autenticazione. Gli errori identici vengono deduplicati sul dispositivo.

## Backend

La callable `recordClientErrorGroup`:

1. richiede un utente Firebase autenticato;
2. applica limite anti-abuso;
3. normalizza e oscura i dati;
4. aggrega gli eventi in `appErrorGroups/{fingerprint}`;
5. conserva soltanto gli ultimi 8 campioni tecnici;
6. crea una notifica amministrativa deduplicata nella raccolta `notifications`;
7. aggiorna il riepilogo `systemCounters/errorCenterSummary`;
8. prova a inviare una push all’amministratore per problemi alti, critici o manuali.

Le raccolte tecniche non sono leggibili direttamente dal client. Il pannello usa callable amministrative mirate:

- `getErrorCenterSummary`;
- `getErrorCenterDashboard`;
- `markErrorCenterSeen`;
- `updateErrorCenterStatus`.

## Stati disponibili

- `open`: Aperto;
- `in_verification`: In verifica;
- `resolved`: Risolto;
- `ignored`: Ignorato.

Un problema risolto o ignorato viene riaperto automaticamente se ricompare con gravità alta/critica o tramite segnalazione manuale.

## Costi e listener

- utenti normali: nessun listener Firestore aggiuntivo;
- amministratore: nessun listener specifico del Centro errori; gli avvisi usano il Centro notifiche già esistente;
- ogni evento registrato: una transazione sul gruppo, un aggiornamento del riepilogo e, solo quando necessario, una notifica amministrativa;
- apertura del Centro errori: una query server-side limitata agli ultimi 200 gruppi;
- badge menu: lettura di un solo documento riepilogativo tramite callable.

## Deploy

Il workflow `.github/workflows/deploy-firebase-functions.yml` distribuisce:

```text
recordClientErrorGroup
getErrorCenterSummary
getErrorCenterDashboard
markErrorCenterSeen
updateErrorCenterStatus
```

Non sono richieste nuove regole Firestore client perché i gruppi tecnici sono gestiti esclusivamente tramite Admin SDK nelle Cloud Functions.

## Controlli

```bash
npm run check:error-monitor
npm run check:fatto-critical
```

Il primo comando verifica monitor, privacy, backend, caricamento PWA e smoke runtime. Il secondo include anche tutte le protezioni FATTO, Whazzup, impianti e dati.
