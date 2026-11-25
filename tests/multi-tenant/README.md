 # Prototipo isolamento aziende (solo emulator)

Non integrare né distribuire queste regole su Firebase reale. Il progetto demo è fittizio, l'unico host consentito è 127.0.0.1:8180. Usare un checkout temporaneo del branch, Node 22, Java compatibile con Firestore Emulator, Firebase CLI e @firebase/rules-unit-testing compatibile con Firebase 11.10.0. Installare dipendenze di test solo nel checkout temporaneo, senza aggiornare package.json o lockfile di produzione. Dal root del checkout, eseguire:

```sh
npm ci
npm install --no-save --package-lock=false @firebase/rules-unit-testing@4 firebase-tools@14 firebase-admin@13
npx firebase emulators:exec --only firestore --project demo-varga-multi-tenant --config tests/multi-tenant/firebase.emulator.json 'node tests/multi-tenant/isolation.test.cjs && node tests/multi-tenant/tenant-service.emulator.test.cjs'
```

Il test cancella soltanto i dati sintetici dell'emulatore locale dedicato. Non utilizzare un emulatore condiviso. Le 17 asserzioni riguardano il prototipo di appartenenza aziendale; non certificano l'app completa. Prima suite di 17 asserzioni superata su GitHub Actions (run 45600354265). La suite estesa va verificata separatamente. Per il piano di sviluppo vedere docs/MULTI_AZIENDA_V2.md.

## Provisioning e inviti (prototipo backend)
La logica in tenant-service.prototype.cjs non è esportata dalle Functions esistenti. Il super amministratore può creare un tenant e un owner; l'owner invita ruoli operativi e admin, un admin soltanto ruoli operativi. Gli inviti sono legati a tenant ed email verificata, con token casuale monouso memorizzato come hash, validità di 7 giorni e accettazione atomica. Nessuna membership esistente è sovrascritta. Il backend futuro dovrà verificare i token Firebase e la revoca, risolvere l'identità del proprietario tramite Admin Auth, applicare rate limiting e idempotenza alle richieste di provisioning/invito, gestire revoca e sospensione con aggiornamento dell'indice personale, trasmettere gli inviti in modo sicuro, e registrare audit senza token. Non esiste ancora un endpoint, UI di registrazione reale, invio email o progetto staging online. I sette giorni e la creazione aziende riservata al super admin sono scelte iniziali modificabili. La verifica di sicurezza completa richiede anche Storage, integrazioni e cache.
