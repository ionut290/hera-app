# Account Apple Review — ruolo a sola lettura e dati sintetici

## Scopo

Creare in futuro un **account Firebase Authentication dedicato** ai revisori Apple, senza concedere la possibilità di premere FATTO, creare/modificare impianti, ore, utenti, personale o registrazioni aziendali. Non usare un normale account `user` approvato: le attuali regole permettono agli account attivi di leggere alcune collezioni contenenti dati reali e, in determinati moduli, di scrivere. Nascondere pulsanti nell'interfaccia non protegge il database.

**Questo ramo è preparatorio. Non attivare l'account, non distribuire regole Firebase e non caricare una build iOS prima delle verifiche elencate.**

## Matrice dei privilegi richiesta

| Funzione | Apple Reviewer | Utente attuale |
| --- | --- | --- |
| Accedere con email/password dedicata | Sì, dopo configurazione | Invariato |
| Vedere commesse/impianti reali | **No** | Invariato |
| Vedere personale/utenti/ore reali | **No** | Invariato |
| Leggere dati fittizi in `appleReviewDemo` | **Sì** | No (salvo accesso server autorizzato) |
| FATTO, reset, forzature, invio WhatsApp | **No** | Invariato |
| Creare, modificare o eliminare impianti | **No** | Invariato |
| Creare, modificare o eliminare ore | **No** | Invariato |
| Creare, modificare o eliminare utenti e personale | **No** | Invariato |
| Caricare o leggere file Firebase Storage | **No** | Invariato |
| Gestione amministrativa | **No** | Invariato |

### Cosa è implementato in questa PR

- `firestore.rules`: funzione `isAppleReviewAccount()` basata esclusivamente sulla **custom claim firmata** `apple_review: true`. Le funzioni comuni `signedIn()`, `isAdmin()` e `isSelf()` escludono tale account. Non cambiano gli utenti normali. Solo la raccolta dedicata `appleReviewDemo`, popolata **esclusivamente con dati fittizi**, è leggibile dal revisore. Nessun client può scriverci.
- `storage.rules`: stessa esclusione dalle letture e dalle scritture, compresi i file condivisi.
- Test isolati sull'emulatore Firebase per letture/scritture vietate e regressioni dei normali utenti. I test **non** si collegano al progetto reale.

### BLOCCANTE: funzioni server

Le Cloud Functions che usano Firebase Admin SDK possono scrivere dati ignorando le regole Firestore. Un account con una custom claim di sola lettura può comunque effettuare una chiamata ad alcune API del backend se queste accettano genericamente qualunque utente Firebase autenticato.

Pertanto, **questa PR non è sufficiente per dichiarare sicuro un account Apple attivo**. Prima di ogni distribuzione bisogna censire tutte le funzioni callable e gli endpoint HTTP, verificare i token e inserire restrizioni coerenti sul backend per `apple_review: true`, incluso l'accesso a Google Drive, eventuali invii, modifiche operative e sistemi collegati. Non modificare il flusso protetto FATTO/WhatsApp; negare l'azione lato API e tramite policy indipendenti. Verificare anche eventuali altri servizi di archiviazione e API non coperte dalle Firebase Rules.

### BLOCCANTE: esperienza di accesso Apple

Serve un percorso **autenticato** che riconosca la custom claim e mostri soltanto la modalità dimostrativa sintetica già creata. L'attuale `apple-review-demo.html` funziona senza account e non legge la raccolta Firebase `appleReviewDemo`. È possibile usare direttamente la demo statica, ma non va presentata come un account attivo con accesso reale finché non è stato integrato e collaudato il percorso dedicato.

### BLOCCANTE: Apple Review

Apple App Review guideline 2.1(a) chiede credenziali demo che funzionino e accesso sufficiente per verificare tutte le funzioni rilevanti. Un account **solo lettura** potrebbe non essere sufficiente per revisionare funzioni di inserimento/modifica. Fornire in un ambiente separato funzioni **di prova** sui dati sintetici, oppure descrivere con precisione il limite e concordare con Apple l'eventuale approccio alternativo. Non garantire l'approvazione.

Fonti:
- https://developer.apple.com/app-store/review/guidelines/
- https://developer.apple.com/help/app-review/before-submitting-for-review/complete-review
- https://firebase.google.com/docs/auth/admin/custom-claims
- https://firebase.google.com/docs/firestore/security/rules-conditions
- https://firebase.google.com/docs/rules/unit-tests

## Procedura futura (dopo le verifiche)

1. Creare una casella email dedicata sotto il controllo del proprietario. **Niente password o token nei commit, nella chat o nelle issue.**
2. Creare l'account in Firebase Authentication senza approvarlo come operatore normale e senza aggiungerlo all'elenco amministratori; per questo non usare in modo indiscriminato strumenti che creano tester generici.
3. Assegnare tramite Firebase Admin SDK (solo credenziali amministrative sicure) la custom claim `apple_review: true` al **UID esatto**; verificare che la modifica sia visibile nel token aggiornato. Nessun ruolo client modificabile dall'utente può attribuire quel privilegio.
4. Verificare in staging/emulator le regole Firestore **e** Storage, e tutte le Cloud Functions. Testare espressamente FATTO, creazione/modifica impianti, ore, utenti e personale, accesso documenti e collegamenti ad altri servizi.
5. Popolare `appleReviewDemo` **solo** con record sintetici tramite uno script/server autorizzato, senza copiare dati reali.
6. Collegare il login dedicato alla vista sintetica senza alterare il flusso normale e collaudare sul dispositivo iOS.
7. Preparare e caricare una nuova build (non riutilizzare build 5), e solo allora inserire username/password realmente funzionanti in **App Store Connect → App Review Information**.
8. Impostare il rilascio manuale e completare le altre verifiche privacy, screenshot, politica dati e classificazione; non inviare prima dell'esito dei test.

## Nessun rilascio

Questa PR deve rimanere in **bozza**, senza deploy Firestore/Storage e senza unire a `main`, finché non sono gestiti i blocchi backend, autenticazione e verifica Apple. Non sono stati creati account o password.
