# Sincronizzazione Varga Cantieri → Varga Gestionale

## Obiettivo

Varga Cantieri resta la sorgente operativa. Varga Gestionale importa una fotografia strutturata dei dati aziendali senza interrogare continuamente tutte le collezioni Firestore.

## Flusso

1. Un amministratore autenticato richiama `rebuildVargaGestionaleSnapshot`.
2. La Cloud Function legge le collezioni aziendali autorizzate e crea uno snapshot compatto a blocchi.
3. `gestionaleSyncState/current` espone soltanto il manifest dello snapshot corrente.
4. Varga Gestionale legge il manifest tramite `getVargaGestionaleSnapshotManifest` e scarica i blocchi con `getVargaGestionaleSnapshotChunk`.
5. Ogni record conserva `sourcePath`, `rootCollection` e l'ID originale di Varga Cantieri, così le sincronizzazioni successive possono aggiornare i record senza duplicarli.

## Dati inclusi

Lo snapshot include le principali aree operative: commesse e relative sotto-collezioni, impianti, squadre, ore, personale, mezzi, documenti condivisi/globali, POS, calendario/programmazioni, Verde Levato e Servizio Neve.

Le sotto-collezioni di `commesse` vengono percorse ricorsivamente, quindi restano disponibili anche moduli gestionali ospitati sotto la commessa, inclusi i dati dei preventivi quando presenti nel percorso `commesse/__preventivi__/...`.

## Dati esclusi per sicurezza

Non vengono esportati documenti privati, log/audit tecnici, posizioni operatori, token push, notifiche, richieste password o configurazioni/segreti Drive. I documenti con `visibility=personal` non entrano nello snapshot.

## Autorizzazioni

- Ricostruzione snapshot: solo amministratore.
- Lettura snapshot: amministratore oppure utente presente in `appConfig/gestionaleUsers`.

Esempio configurazione:

```json
{
  "emails": ["utente@azienda.it"],
  "uids": ["firebase-auth-uid"]
}
```

## Consumo Firestore

Le letture sulle collezioni operative avvengono quando viene ricostruito lo snapshot, non ogni volta che un utente apre Varga Gestionale. Gli utenti del Gestionale leggono poi solo il manifest e i blocchi dello snapshot corrente. Lo snapshot precedente viene eliminato dopo che il nuovo è pronto.

## Cloud Functions

- `rebuildVargaGestionaleSnapshot`
- `getVargaGestionaleSnapshotManifest`
- `getVargaGestionaleSnapshotChunk`

Regione: `europe-west1`.
