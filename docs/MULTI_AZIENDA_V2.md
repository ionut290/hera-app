# VARGA CANTIERI 2.0 — Multi-azienda (fase 0)
Stato: progettazione e prototipo isolato, NON pronto per produzione.
Base main: b5a69aff73a0f143cd3eb05a051d2786eb66cb59.

## Riscontri dal codice esistente
- `capacitor.config.json`: VARGA CANTIERI, wrapper Capacitor; preservare appId e continuità dello store.
- `firestore.rules`: `signedIn()` e `isAdmin()` sono globali. Commesse e molte sottocollezioni sono leggibili da utenti approvati; non sono confini aziendali.
- `storage.rules`: proprietà UID e condivisione globale/condivisa; non è sufficiente per isolare aziende.
- `README.md`: Drive centralizzato e bridge amministrativo; evitare qualsiasi riutilizzo del Drive storico per nuove aziende.
- `AUTH_STABILITY_CONTRACT.md`, `AGENTS.md`: proteggere login, FATTO/WHAZZUP/RESET, impianti, ID, personale, utenti, collegamenti, cache, offline, storico. Non cambiare queste funzioni né aggirare i controlli.
- `sharedStaticViews`, notifiche globali, backend Admin SDK, Netlify, Varga Gestionale, esportazioni e cache offline sono ulteriori superfici da censire.

## Architettura proposta
Ambiente Firebase di sviluppo separato (nuovo progetto di staging da autorizzare/provisionare; nessuna connessione a produzione). Il prototipo in tests/multi-tenant è esclusivamente una base sperimentale, non modifica le regole di produzione.
- `tenants/{tenantId}`: metadati aziendali, stato attivo/sospeso.
- `tenants/{tenantId}/members/{uid}`: ruolo, stato, eventuale riferimento al personale storico. Una persona può avere più aziende. Solo backend fidato gestisce appartenenze e ruoli.
- `users/{uid}/memberships/{tenantId}`: indice personale per il selettore, scritto esclusivamente dal backend e verificato contro membership autorevole. Nessuna autorizzazione basata sul selettore o su localStorage.
- `tenants/{tenantId}/commesse/{id}/...`, `personale`, `squadre`, `mezzi`, `rapportini`, `sharedStaticViews`, `notifications`: namespace proposto; ogni collezione richiede specifiche regole di autorizzazione e migrazione non distruttiva. Non implementare fallback a dati globali.
- Storage: prefisso per tenant, metadati e accesso con membership verificata; nessun link pubblico automatico; audit dei token URL e condivisioni.
- Ruoli minimi: owner/admin, manager, team lead, worker, client. Nessuna concessione implicita a tenant admin di privilegi di piattaforma. Super admin gestisce metadati/provisioning; accessi ai dati aziendali eccezionali, autorizzati e tracciati.
- Backend: autorizzazione indipendente per ogni operazione, non fidarsi di tenantId dal client. Inviti monouso, scadenza, indirizzo verificato, transazione server, revoca effettiva. App Review solo demo sintetica isolata.
- Cambio azienda: annullare listener, azzerare solo cache temporanee dell'interfaccia (non dati persistenti storici), segregare cache e coda offline per tenant+UID, impedire invii pendenti verso altra azienda; nessuna modifica al flusso protetto FATTO/WHAZZUP.
- Drive: disattivato per tenant nuovi finché non esiste integrazione isolata per azienda e verifica server dell'appartenenza di file/cartelle. Non condividere il Drive originale.

## Fasi e gate
1. Inventario completo di path/query, indici, functions, triggers, condivisioni, export, cache, notifiche e integratori. Identificare dipendenze con Varga Gestionale e funzioni sigillate. Non toccare percorsi protetti.
2. Provisioning staging indipendente, dati fittizi A/B, regole deny by default, test emulator. Prototipo iniziale consente solo lettura di metadati e appartenenze autorizzate; nessuna scrittura client.
3. Backend di registrazione azienda/inviti/ruoli con audit e test. UI selezione azienda dopo login; eventuale modifica al contratto login richiede autorizzazione specifica. Separare branding e configurazioni geografiche aziendali.
4. Definire accessi funzionali collezione per collezione. Bloccare l'integrazione se interferisce con FATTO/WHAZZUP: nessun adattamento o bypass del componente sigillato. Verificare strategia compatibile prima di procedere.
5. Prova con due aziende fittizie su PWA/iOS/Android; nessun dato storico in staging prima di consenso e autorizzazione. Suite regressioni e test sicurezza.
6. Migrazione solo dopo approvazione distinta: backup verificato, manifest sorgente/destinazione, conteggi e relazioni/ID invariati, copia idempotente senza cancellazioni, riconciliazione delle variazioni durante cutover, controllo file e permessi, rollback provato. Non modificare dati originali. Nessun collegamento automatico a un tenant per email o nome.
7. Release separatamente autorizzata: App Store update/TestFlight, Android, web, privacy e contratto DSA se cambia il modello commerciale. Non deployare staging su produzione; nessun merge su main prima delle verifiche.

## Test di accettazione obbligatori (non ancora eseguiti)
- A legge A; A non legge B nemmeno con UID/ID noto, query, link file o chiamate backend; B simmetrico.
- Utente anonimo, sospeso, revocato o reviewer non accede ai tenant; ruolo client limitato; nessuna auto-promozione o auto-iscrizione.
- Due membership attive e cambio azienda senza fughe da cache, notifiche, export, file, code offline o listener.
- Test regole Firestore/Storage con emulator; backend Admin SDK e integrazioni testati separatamente. Test di migrazione con dati sintetici e rollback.
- Controlli preesistenti: npm run check:fatto-critical, check:auth-stability e workflow E2E/native; test manuali su dispositivi. Nessuna regressione sulle funzioni sigillate.
- Misurare letture/scritture/listener, senza polling né duplicazioni.

## Limiti di questa fase
Nessun provisioning, migrazione, deploy o modifica ai file runtime. Nessun accesso ai dati live. Nessun esito dei test emulator dichiarato. Non usare il prototipo come regole definitive. La scelta di staging separato è una proposta prudenziale, non una decisione commerciale definitiva.
