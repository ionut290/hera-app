 # Prototipo isolamento aziende (solo emulator)

Non integrare né distribuire queste regole su Firebase reale. Il progetto demo è fittizio, l'unico host consentito è 127.0.0.1:8180. Usare un checkout temporaneo del branch, Node 22, Java compatibile con Firestore Emulator, Firebase CLI e @firebase/rules-unit-testing compatibile con Firebase 11.10.0. Installare dipendenze di test solo nel checkout temporaneo, senza aggiornare package.json o lockfile di produzione. Dal root del checkout, eseguire:

```sh
npm ci
npm install --no-save --package-lock=false @firebase/rules-unit-testing@4 firebase-tools
npx firebase emulators:exec --only firestore --project demo-varga-multi-tenant --config tests/multi-tenant/firebase.emulator.json 'node tests/multi-tenant/isolation.test.cjs'
```

Il test cancella soltanto i dati sintetici dell'emulatore locale dedicato. Non utilizzare un emulatore condiviso. Le 17 asserzioni riguardano il prototipo di appartenenza aziendale; non certificano l'app completa. Non è ancora stata eseguita la suite. Per il piano di sviluppo vedere docs/MULTI_AZIENDA_V2.md.
