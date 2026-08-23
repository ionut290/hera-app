# PDF nei Lavori occasionali

I PDF dei cantieri occasionali vengono caricati in Firebase Storage sotto:

`lavori-occasionali/{plantId}/{ownerUserId}/{documentId}/{fileName}`

I metadati sono salvati nella sotto-collezione Firestore:

`commesse/lavori-occasionali/impianti/{plantId}/documentiPdf/{documentId}`

Regole principali:
- solo PDF;
- massimo 15 MB per file;
- lettura disponibile agli utenti autenticati dell'app;
- caricamento/cancellazione Storage consentiti al proprietario del file;
- nessuna scadenza automatica;
- la funzione è confinata alla commessa Lavori occasionali.
