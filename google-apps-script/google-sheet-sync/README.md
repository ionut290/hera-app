# Sincronizzazione bidirezionale Google Sheet

Questo Apps Script è il ponte di **scrittura** usato da Varga Cantieri:

`modifica nell'app → Netlify Function autenticata → Apps Script doPost → Google Sheet`

La direzione opposta usa il proxy GViz già presente:

`modifica nel Google Sheet → GViz CSV → aggiornamento app`

## 1. Creare e pubblicare lo script

1. Apri Google Apps Script con lo stesso account che può modificare il Google Sheet.
2. Crea un progetto e sostituisci `Code.gs` con il file presente in questa cartella.
3. In **Impostazioni progetto → Proprietà script** aggiungi:
   - `SYNC_SECRET`: una stringa lunga e casuale;
   - `ALLOWED_SPREADSHEET_IDS`: facoltativo, uno o più ID separati da virgola.
4. Pubblica come **Applicazione web**:
   - Esegui come: **Me**;
   - Chi ha accesso: **Chiunque**.
5. Copia l'URL terminante in `/exec`.

Il foglio deve essere modificabile dall'account proprietario dello script. Per la lettura GViz deve inoltre essere condiviso come **Chiunque abbia il link → Visualizzatore**.

## 2. Configurare Netlify

Aggiungi nelle variabili d'ambiente del sito:

- `GOOGLE_SHEET_APPS_SCRIPT_URL`: URL `/exec` dell'applicazione web;
- `GOOGLE_SHEET_SYNC_SECRET`: lo stesso valore impostato in `SYNC_SECRET`;
- `FIREBASE_PROJECT_ID`: `hera-app-6cd2b`;
- `GOOGLE_SHEET_SYNC_ALLOWED_EMAILS`: facoltativo, email amministratori separate da virgola.

Esegui poi un nuovo deploy Netlify.

## 3. Attivare la sincronizzazione nell'app

Nella schermata **Gestione impianti e contabilità**:

1. incolla il link del Google Sheet;
2. premi **Salva collegamento**;
3. scegli la prima direzione con **Aggiorna app dal foglio** oppure **Invia app al foglio**;
4. attiva **Sync automatica** e scegli 5, 15, 30 o 60 minuti.

La prima direzione è sempre manuale per evitare che un foglio vuoto cancelli dati o che l'app sovrascriva per errore un foglio esistente.

Le prime due colonne tecniche (`SYNC_KEY` e `IMPIANTO_KEY`) vengono nascoste dallo script: non eliminarle, perché consentono aggiornamenti e cancellazioni senza duplicare le righe.

## 4. Matrici Personale e Mezzi

Le sezioni **Personale** e **Mezzi** usano lo stesso endpoint autenticato e possono creare o collegare un file con i fogli `PERSONALE`, `MEZZI`, `COMMESSE_PERSONALE`, `COMMESSE_MEZZI` e `LOG_SINCRONIZZAZIONE`. Non è necessario rendere pubblico questo file: la lettura e la scrittura avvengono tramite Apps Script.

Conservare in ciascun foglio le colonne `RECORD_ID`, `UPDATED_AT`, `UPDATED_BY`, `SYNC_VERSION`, `SYNC_SOURCE` e `ROW_STATUS`. La sincronizzazione è incrementale, risolve i conflitti in base alla versione e alla data di aggiornamento e non elimina automaticamente righe né record. Dopo aver aggiornato `Code.gs`, creare una nuova versione del deployment Apps Script e verificare che `GOOGLE_SHEET_APPS_SCRIPT_URL` punti al deployment aggiornato.
