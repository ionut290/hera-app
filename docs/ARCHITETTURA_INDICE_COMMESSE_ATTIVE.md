# Architettura: indice leggero delle commesse attive

## Obiettivo

Ridurre drasticamente letture e listener Firestore senza nascondere o perdere lo storico ore.

## Soluzione scelta

Usare un solo documento leggero:

`appConfig/activeCommesse`

Esempio:

```json
{
  "ids": ["commessaA", "commessaB", "commessaC"],
  "updatedAt": "serverTimestamp",
  "updatedBy": "uid-admin"
}
```

All'avvio l'app legge questo documento una sola volta o mantiene un solo listener su questo documento. Solo gli ID presenti sono considerati operativi.

## Regole obbligatorie

1. Una commessa disattivata non viene eliminata e mantiene lo stesso ID.
2. Non si cancellano impianti, ore, squadre, note, documenti o storico.
3. Gli impianti vengono caricati e ascoltati solo per le commesse presenti in `activeCommesse.ids`.
4. Il calendario personale continua a leggere lo storico ore senza filtrare per commessa attiva.
5. Le ore già registrate devono conservare almeno `commessaId`, `commessaNome` e, se presente, `commessaCodice`.
6. Se una vecchia registrazione contiene solo `commessaId`, il calendario può risolvere il nome tramite una cache metadati, ma non deve aprire listener sugli impianti.
7. La riattivazione reinserisce l'ID nell'indice e riusa tutti i dati esistenti.
8. FATTO e WhatsApp non devono essere modificati.

## Interfaccia amministratore

In Gestione commesse mostrare:

- `Attiva` con pulsante `Disattiva`;
- `Archiviata` con pulsante `Riattiva`;
- filtro `Tutte / Attive / Archiviate`.

Prima della disattivazione mostrare:

> La commessa verrà nascosta dalle sezioni operative e i suoi listener impianti verranno chiusi. Nessun dato sarà eliminato. Le ore già registrate resteranno nel calendario personale.

## Flusso di disattivazione

1. Verificare ruolo amministratore.
2. Leggere `appConfig/activeCommesse` in transazione.
3. Rimuovere soltanto l'ID selezionato dall'array.
4. Aggiornare `updatedAt` e `updatedBy`.
5. Chiudere immediatamente i listener impianti associati a quell'ID.
6. Rimuovere la commessa dalle viste operative senza modificare lo storico.

## Flusso di riattivazione

1. Verificare ruolo amministratore.
2. Aggiungere l'ID all'array con `arrayUnion` o transazione equivalente.
3. Non creare una nuova commessa e non generare un nuovo ID.
4. Ricaricare metadati e impianti della sola commessa riattivata.

## Compatibilità iniziale

Se `appConfig/activeCommesse` non esiste ancora:

- considerare temporaneamente tutte le commesse attive;
- creare l'indice una sola volta lato amministratore;
- non scrivere automaticamente a ogni apertura;
- non modificare i documenti storici.

## Calendario personale

Il calendario personale deve essere indipendente dall'indice delle commesse attive.

Sono vietati filtri del tipo:

```js
ore.filter((riga) => activeCommesseIds.includes(riga.commessaId))
```

La visualizzazione deve includere tutte le ore storiche dell'utente, anche quando la commessa è archiviata.

## Risparmio atteso

Con 17 commesse attualmente ascoltate e solo 5 operative, i listener impianti possono scendere indicativamente da 17 a 5, più un solo listener leggero sull'indice. Il risparmio effettivo dipende dal numero di documenti impianto e dalle riaperture dell'app.

## Test obbligatori

- disattivare una commessa con ore storiche;
- verificare che scompaia dalla home operativa;
- verificare che non parta il listener `commesse/{id}/impianti`;
- verificare che le ore rimangano nel calendario personale;
- riattivare e verificare stesso ID e stessi impianti;
- verificare squadre, ore, FATTO e WhatsApp sulle commesse rimaste attive;
- verificare che nessun documento venga eliminato;
- confrontare diagnostica Firestore prima e dopo.
