# Verifica cache locale personale e mezzi

La modifica resta limitata alle collezioni `personale` e `mezzi` e non modifica calendario, squadre, ore, FATTO o WhatsApp.

## Comportamenti da verificare

- La cache IndexedDB viene caricata prima di `app.js`.
- Un `get()` sulla collezione completa `personale` usa la cache locale quando contiene record validi e recenti.
- Un `get()` sulla collezione completa `mezzi` usa la cache locale quando contiene record validi e recenti.
- La risposta locale mantiene le proprietà usate di un `QuerySnapshot`: `docs`, `size`, `empty`, `forEach()` e `docChanges()`.
- Ogni documento locale mantiene `id`, `ref`, `exists`, `data()` e `get()`.
- Se la cache manca, è scaduta o contiene record senza ID, viene usata la query Firestore originale.
- Il listener Firestore continua a essere la fonte di aggiornamento e riscrive IndexedDB soltanto quando l’elenco cambia.
- Tre `get()` contemporanei sulla stessa collezione condividono una sola Promise.
- Le query con `source` esplicita e tutte le altre collezioni continuano a usare direttamente Firestore.
- Le scritture normali su personale e mezzi invalidano la copia recente dopo il successo.
- Le patch automatiche limitate a nome, email e foto non svuotano l’intera cache: il listener aggiorna la copia.
- Patch automatiche identiche sullo stesso documento entro 60 secondi vengono eseguite una sola volta.

## Controllo dopo il deploy

Aprire la console ed eseguire:

```js
HeraFirestoreRegistryOptimizer.getState()
```

Contatori principali:

- `reusedDeviceCache`: caricamenti completi evitati grazie a IndexedDB;
- `reusedInFlight`: richieste contemporanee accorpate;
- `reusedRecent`: risultati recenti riutilizzati;
- `profileWritesSkipped`: scritture profilo identiche evitate;
- `networkGets`: caricamenti realmente inoltrati a Firestore.
