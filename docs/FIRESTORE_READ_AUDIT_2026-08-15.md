# Audit letture Firestore — Hera App / Varga Cantieri

Data audit: 2026-08-15
Base analizzata: branch `main`
Obiettivo: ridurre letture Firestore senza compromettere caricamento commesse/squadre, FATTO, WhatsApp/WHAZZUP, ore, mappe e sincronizzazione multi-dispositivo.

## Metodo

Audit statico dell'intero repository orientato alle chiamate Firestore (`onSnapshot`, `.get()`, accessi `firebase.firestore`, query e moduli che intercettano le letture) e verifica dei meccanismi di ottimizzazione già presenti.

L'audit automatico ha censito 504 siti di lettura in 53 file, con 50 listener e 70 candidati ad alto rischio. La classificazione è volutamente conservativa: un sito segnalato non è automaticamente un problema, perché alcune letture sono dentro funzioni realmente on-demand.

## Protezioni già presenti

- `firestore-operation-diagnostics.js`: conta documenti, operazioni, listener, query, funzioni chiamanti e schermate.
- `firestore-safe-optimizer.js`: condivide listener fisici identici di `commesse`, `squadreStorico` e `userAlerts`.
- `firestore-startup-cost-optimizer.js`: blocca listener non necessari all'avvio e usa viste condivise per le squadre quando possibile.
- `activity-logs-read-disable.js`: evita letture dell'archivio attività operativo e limita le commesse caricate all'avvio.
- `shared-static-views*`: sostituisce dataset ripetitivi con viste aggregate.
- `firestore-registry-read-optimizer.js` + cache dispositivo: riduce letture anagrafiche.

## Prima riduzione reale applicata — `appNotifications`

L'audit ha confermato che il vecchio `subscribeGlobalNotifications` in `app.js` può aprire all'avvio un listener su `appNotifications` con `limit(40)`. Il Centro notifiche corrente usa già i flussi moderni `notifications` e `userAlerts`, perciò quella lettura legacy non è necessaria per la UI corrente.

È stato aggiunto `app-notifications-read-guard.js` con queste regole:
- blocca soltanto `get()` e `onSnapshot()` sulla raccolta `appNotifications`;
- restituisce snapshot vuoti compatibili al codice legacy;
- non blocca `add`, `set`, `update`, `delete` o altre scritture;
- viene caricato da `firebase-config.js` prima di `app.js` e prima della diagnostica;
- viene incluso nella cache PWA e nel bundle Android, che copia automaticamente i file JS root;
- un controllo statico impedisce di reintrodurre letture o di bloccare accidentalmente le scritture.

**Risultato atteso:** da fino a 40 documenti iniziali del listener legacy a **0 letture Firestore di rete su `appNotifications`** per sessione normale.

## Candidati successivi

Priorità runtime:
- `app.js`: listener di avvio e listener che devono esistere solo quando una schermata è aperta;
- `squadreStorico`: verificare che `staticSquadreFallbacks` resti 0;
- login/accesso: eliminare eventuali `get()` duplicati in pochi secondi;
- `documents.js`, Preventivi e Contabilità: confermare che non leggano prima dell'apertura della funzione;
- backend: verificare scansioni complete e rebuild ripetuti separatamente dal client.

## Soglie per una sessione standard

Apertura app → Home → una commessa → un impianto, senza pannelli amministrativi:
- listener duplicati identici: 0;
- `activityLogs`: 0 letture;
- `appNotifications`: 0 letture;
- `chatMessages` senza aprire chat: 0 letture;
- `commessaResources` senza aprire risorse: 0 letture;
- Preventivi/Contabilità senza aprirli: 0 letture;
- documenti privati senza aprire Documenti: 0 letture;
- fallback `squadreStorico`: 0 in condizioni normali;
- commesse disattivate: 0 letture all'avvio.

## Regola di sicurezza

Ogni riduzione va applicata una alla volta e verificata prima/dopo con la diagnostica. Non devono cambiare: FATTO, ordinario/straordinario, foto e WHAZZUP/WhatsApp, commesse, squadre, ore, mappe, autenticazione, Android e PWA.

## Stato della fase 2

La prima riduzione è stata implementata nel branch di audit insieme al relativo controllo di regressione. Prima del merge va verificata la CI completa e poi confrontata una diagnostica runtime nuova con quella precedente per confermare che `appNotifications` resti a zero senza regressioni funzionali.
