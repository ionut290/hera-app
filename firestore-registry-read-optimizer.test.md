# Verifica manuale ottimizzatore Firestore

Questa modifica è limitata alle collezioni complete `personale` e `mezzi`.

## Controlli eseguiti

- Tre `get()` contemporanei sulla stessa collezione producono una sola richiesta reale.
- Un `get()` ripetuto entro la finestra breve riusa lo stesso `QuerySnapshot`.
- Il primo `QuerySnapshot` ricevuto da `onSnapshot()` può soddisfare i `get()` successivi.
- `set()`, `update()`, `delete()` e `add()` invalidano immediatamente il dato recente.
- Le query con `source` esplicita continuano a usare direttamente Firestore.
- Le altre collezioni e le query non identificabili con sicurezza non vengono intercettate.

## Controllo dopo il deploy

Aprire la console ed eseguire:

```js
HeraFirestoreRegistryOptimizer.getState()
```

I contatori `reusedInFlight` e `reusedRecent` indicano quante richieste duplicate sono state evitate.
