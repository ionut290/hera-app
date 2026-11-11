# VARGA CANTIERI — modalità demo isolata per App Review

**Stato:** ramo di sviluppo; non distribuire, non usare come account Firebase reale, non dichiarare ancora pronta per la revisione Apple.

## Perché serve
Nelle regole Firestore correnti, `signedIn()` permette la lettura di collezioni operative (commesse, squadre, personale e altre). Creare e approvare un normale account Firebase per Apple senza isolamento sarebbe rischioso: potrebbe esporre persone e dati aziendali.

## Soluzione predisposta nel ramo
- Da `index.html`, nella schermata **Login richiesto**, link **Esplora demo · dati fittizi**.
- `apple-review-demo.html` è una pagina autonoma, senza il bundle Firebase e senza SDK esterni.
- CSP `default-src 'none'; connect-src 'none'; form-action 'none'`; solo asset locali.
- `apple-review-demo.js` conserva dati **esclusivamente nella memoria della pagina**. Non usa cookie, IndexedDB, localStorage, sessionStorage, Firebase, GPS, invii a WhatsApp, fotografica o upload.
- Panoramica, due commesse fittizie, quattro impianti, ricerca, filtri, dettagli, simulazione di intervento, squadre fittizie, inserimento ore dimostrative e documenti-segnaposto.
- Il completamento fittizio è una simulazione indipendente e **non è** il componente FATTO: non modifica né aggira la logica protetta.
- Link **Torna al login**; il normale accesso email/Google e la sessione reale restano invariati.

## Limiti dichiarati
Questa è una **anteprima dimostrativa isolata**, non un account Firebase e non un ambiente di staging completo. Alcune operazioni reali (GPS, foto, notifiche push, WhatsApp, caricamento file e flussi d'amministrazione) sono rappresentate soltanto come esempio o non disponibili. Non inserire credenziali fittizie nella schermata Sign-In Information di Apple: non funzionerebbero.

Le linee guida Apple **2.1(a)** richiedono un account demo attivo oppure una **modalità demo pienamente funzionale**; nei casi in cui non è possibile fornire un account per ragioni legali/di sicurezza, la modalità demo alternativa richiede **approvazione preventiva di Apple**.
- https://developer.apple.com/app-store/review/guidelines/
- https://developer.apple.com/videos/play/tech-talks/10885/

**Prima di presentare questa modalità come alternativa all'account Apple:**
1. Ottenere chiarimento/approvazione preliminare dal team App Review e spiegare che l'accesso a dati di personale e cantieri reali è soggetto a vincoli privacy e contrattuali.
2. Per soddisfare il requisito delle funzionalità complete, completare o sostituire i flussi mancanti con un vero ambiente di staging separato prima dell'invio.
3. Alternativa preferibile per test integrali: creare un **progetto Firebase QA separato** con database sintetico, regole validate, account reviewer attivo e servizi di backend indipendenti. Questo richiede configurazioni e autorizzazioni che non possono essere inventate dal codice repository.
4. Verificare la nuova build su iPhone, eseguire tutti i controlli, poi aggiornare versione e numero build; **non** riutilizzare build 5 perché il suo bundle non contiene la demo.
5. Non unire la PR e non distribuire senza approvazione. Non cambiare segreti e regole Firebase di produzione.

## Test di verifica
- `node scripts/check-ios-review-demo.js`: CSP, isolamento, dati fittizi, sintassi JS, e packaging (se presente `www/`).
- `npm run android:aab:prepare` per verificare l'inclusione degli asset (non installa o pubblica l'app).
- `npx playwright test tests/e2e/apple-review-demo.spec.js --browser=chromium --workers=1` dopo installazione dipendenze in ambiente di test.
- Controlli esistenti login/FATTO/Whazzup/Firestore e smoke E2E della PR.

## Perimetro delle modifiche
Questo ramo aggiunge i file `apple-review-demo.html`, `apple-review-demo.js`, `apple-review-demo.css`, i test e la documentazione; aggiunge **un solo link** alla schermata login e il relativo stile in `login-retry-fix.css`. Non cambia il JavaScript dell'autenticazione, i token, gli utenti, le regole Firestore, le credenziali o il sistema di distribuzione iOS.

**Importante:** questa preparazione non crea un account dimostrativo reale, non carica su Apple e non garantisce l'approvazione App Review.
