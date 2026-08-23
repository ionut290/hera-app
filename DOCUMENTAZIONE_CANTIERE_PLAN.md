# Documentazione cantiere — implementazione sicura

## Obiettivo
Aggiungere una sezione documentale per ogni impianto/cantiere senza modificare i flussi protetti FATTO e Whazzup e senza alterare le cache globali degli impianti.

## UI richiesta
- Pulsante amministratore sotto l'ingranaggio: `📁 DOCUMENTAZIONE CANTIERE`.
- Se esiste almeno un documento, mostrare nella card impianto il pulsante `📎 LEGGI DOCUMENTAZIONE` nel punto sotto i pulsanti principali.
- Gli operatori possono solo leggere; caricamento, modifica ed eliminazione restano amministratore.
- Al click su NAVIGA: se non esiste documentazione, comportamento identico a oggi; se esiste, mostrare una schermata non bloccante con `VISUALIZZA DOCUMENTI` e `CONTINUA A NAVIGARE`.

## Tipi documento
- Preventivo
- Foto prima lavori
- Foto dopo lavori
- Planimetria
- POS / sicurezza
- Ordine di lavoro
- Autorizzazioni
- Verbali
- Consuntivo / fattura
- Altro

## Persistenza
Riutilizzare Firebase Storage + Firestore già presenti.

Percorso Storage suggerito:
`documents/{ownerUid}/{documentId}/{fileName}`

Raccolta Firestore: `documents`.

Campi minimi:
- `source: "documentazione-cantiere"`
- `commessaId`
- `impiantoKey`
- `impiantoId`
- `impiantoNome`
- `category`
- `title`
- `fileName`
- `mimeType`
- `fileSize`
- `storagePath`
- `downloadUrl`
- `visibility: "global"`
- `sharedToAll: true`
- `showBeforeNavigation`
- `requiredBeforeWork`
- `note`
- `createdBy`
- `createdByEmail`
- `createdByName`
- `createdAt`
- `updatedAt`

## Requisiti di sicurezza
1. Non modificare funzioni, listener o salvataggi del pulsante FATTO.
2. Non modificare funzioni Whazzup/WhatsApp.
3. Non scrivere in `currentImpianti` o `impiantiByCommessaId`.
4. Nessun listener Firestore realtime aggiuntivo per la documentazione.
5. Caricare i documenti solo on-demand quando si apre la sezione o quando si preme NAVIGA.
6. Nessuna migrazione dei documenti impianto esistenti.
7. Il modulo deve fallire in modo isolato: se Firestore/Storage non risponde, NAVIGA continua a funzionare normalmente mostrando un messaggio non invasivo.

## Nota importante
Il tentativo precedente di estendere il modulo PDF a tutte le commesse è stato revertito perché provocava regressioni sui cantieri. Questa nuova implementazione deve quindi essere sviluppata come modulo indipendente e non deve riutilizzare il codice che decora o muta le card tramite cache globali.
