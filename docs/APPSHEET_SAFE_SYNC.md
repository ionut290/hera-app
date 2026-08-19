# Sincronizzazione sicura Varga Cantieri → AppSheet

## Stato

Preparata ma **NON attivata**. Questo modulo non viene importato da `functions/main.js`, quindi non cambia il runtime, il pulsante FATTO, i listener client, Netlify o i dati correnti.

AppSheet target:

- App ID: `a33fc9cd-0a18-4aa8-b70a-c067c0c6c278`
- App URL: `https://www.appsheet.com/start/a33fc9cd-0a18-4aa8-b70a-c067c0c6c278`

## Obiettivo

1. Copiare in AppSheet tutte le commesse e tutti gli impianti già presenti in Firestore.
2. Dopo l'attivazione, ogni nuova commessa/impianto viene aggiunta in AppSheet.
3. Ogni modifica Firestore viene inviata ad AppSheet.
4. Quando Varga Cantieri salva lo stato FATTO, AppSheet riceve l'aggiornamento **dopo** il salvataggio Firestore, senza modificare il codice FATTO.
5. Se AppSheet non risponde, Varga Cantieri continua a funzionare e il dato Varga non viene annullato.
6. Nessuna cancellazione automatica viene inviata ad AppSheet.

## Perché è sicuro

La sincronizzazione è lato Cloud Functions e osserva i documenti Firestore dopo la scrittura. Non modifica:

- `fatto-button-immediate.js`;
- listener o funzioni FATTO;
- WhatsApp/WHAZZUP;
- struttura dei documenti `commesse/{commessaId}/impianti/{impiantoId}`;
- ID esistenti;
- coordinate;
- funzionamento offline del client.

## File preparati

- `functions/appsheet-sync-core.js`: mapping, validazione e payload AppSheet, senza dipendenze Firebase.
- `functions/appsheet-sync.js`: trigger Firestore + backfill amministrativo.
- `scripts/check-appsheet-sync-core.js`: test permanente delle regole base.

## Configurazione richiesta prima dell'attivazione

L'API AppSheet deve essere disponibile sul piano utilizzato e deve essere generata una `ApplicationAccessKey`.

Salvare la chiave soltanto in Firebase Secret Manager:

```bash
firebase functions:secrets:set APPSHEET_ACCESS_KEY
```

Creare il documento Firestore `appConfig/appsheetSync` con `enabled: false` mentre si prepara AppSheet. Esempio di struttura (i nomi delle colonne AppSheet vanno adattati alla tabella reale):

```json
{
  "enabled": false,
  "appId": "a33fc9cd-0a18-4aa8-b70a-c067c0c6c278",
  "region": "www.appsheet.com",
  "locale": "it-IT",
  "timezone": "Europe/Rome",
  "tables": {
    "commesse": {
      "tableName": "COMMESSE",
      "keyColumn": "VARGA_ID",
      "fields": {
        "VARGA_ID": "$id",
        "NOME": "nome",
        "CODICE": "codice"
      }
    },
    "impianti": {
      "tableName": "IMPIANTI",
      "keyColumn": "VARGA_ID",
      "fields": {
        "VARGA_ID": "$id",
        "COMMESSA_ID": "$parentId",
        "ID_SAP": "idSap",
        "DENOMINAZIONE": "denominazione",
        "COMUNE": "comune",
        "STATO": "stato",
        "DATA_FATTO": "doneAt",
        "OPERATORE": "doneBy"
      }
    }
  }
}
```

I campi sopra sono solo un esempio di mapping: prima dell'attivazione vanno confrontati con i campi Firestore realmente usati e con le colonne reali delle tabelle AppSheet.

## Attivazione controllata

Solo dopo aver verificato AppSheet e il mapping:

1. importare `./appsheet-sync` da `functions/main.js` e aggiungerlo a `Object.assign(exports, ...)`;
2. distribuire le sole Cloud Functions coinvolte;
3. mantenere `enabled: false`;
4. eseguire un test su una commessa/impianto di prova;
5. impostare `enabled: true`;
6. eseguire `backfillAppSheetFromVarga` come admin per il caricamento iniziale;
7. verificare conteggi e ID prima di usare la sincronizzazione ordinaria.

Il backfill usa `Find` per capire se la chiave Varga esiste già e poi esegue `Add` o `Edit`. Non elimina righe AppSheet.

## Regola di rollback

Per fermare immediatamente la sincronizzazione senza modificare Varga Cantieri impostare:

```text
enabled = false
```

nel documento `appConfig/appsheetSync`.
