# Generazione Android App Bundle per Google Play

Il workflow `.github/workflows/build-android-aab.yml` crea automaticamente un file `.aab` firmato senza richiedere Android Studio.

## Secret GitHub obbligatori

Nel repository aprire:

`Settings → Secrets and variables → Actions → New repository secret`

Creare esattamente questi quattro secret:

- `ANDROID_KEYSTORE_BASE64`
- `ANDROID_KEYSTORE_PASSWORD`
- `ANDROID_KEY_ALIAS`
- `ANDROID_KEY_PASSWORD`

I valori sono contenuti nel file `GITHUB-SECRETS.txt` consegnato separatamente insieme alla chiave `varga-cantieri-upload.jks`.

## Avvio della compilazione

1. Aprire la scheda `Actions` del repository.
2. Selezionare `Build Android App Bundle`.
3. Premere `Run workflow`.
4. Attendere il completamento verde.
5. Aprire l'esecuzione e scaricare l'artefatto `VARGA-CANTIERI-AAB-...`.
6. Estrarre il file `.aab` e caricarlo nella release di test interno della Play Console.

## Importante

Conservare per sempre la chiave `varga-cantieri-upload.jks` e le password. La stessa chiave deve essere utilizzata per tutti gli aggiornamenti futuri dell'app.

Il workflow assegna automaticamente un `versionCode` crescente usando il numero dell'esecuzione GitHub Actions, evitando il rifiuto degli aggiornamenti per numero di versione già utilizzato.
