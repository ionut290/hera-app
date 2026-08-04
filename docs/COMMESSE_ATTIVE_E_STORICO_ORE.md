# Commesse attive/disattivate con conservazione dello storico ore

## Obiettivo

Consentire all'amministratore, nella schermata **Gestione commesse**, di disattivare temporaneamente una commessa non utilizzata e di riattivarla in seguito.

La disattivazione serve esclusivamente a ridurre letture Firestore, listener e caricamenti operativi. Non deve eliminare né nascondere dati storici.

## Campo dati

Ogni documento della collezione `commesse` può contenere:

```js
attiva: true
```

Compatibilità obbligatoria con i dati esistenti:

```js
const isCommessaAttiva = commessa.attiva !== false;
```

Quindi una commessa priva del campo `attiva` deve essere considerata attiva. Non eseguire migrazioni massive e non riscrivere tutte le commesse esistenti.

## Interfaccia Gestione commesse

Ogni scheda della schermata `Gestione commesse` deve mostrare:

- badge `ATTIVA` oppure `DISATTIVATA`;
- pulsante/interruttore `Disattiva commessa` quando attiva;
- pulsante/interruttore `Riattiva commessa` quando disattivata.

La modifica dello stato deve aggiornare esclusivamente il documento della singola commessa e solo quando il valore cambia.

Prima della disattivazione mostrare una conferma chiara:

> La commessa verrà nascosta dalle attività operative e i suoi impianti non saranno caricati automaticamente. Le ore già registrate resteranno visibili nel calendario personale. Nessun dato verrà eliminato.

## Regole operative

Quando `attiva === false`, la commessa non deve:

- comparire nell'elenco operativo principale;
- essere proposta per nuove composizioni squadre;
- essere proposta per nuove programmazioni;
- aprire listener automatici sulle sottocollezioni `impianti`;
- caricare automaticamente impianti, statistiche e risorse operative all'avvio;
- generare letture Firestore automatiche non necessarie.

La commessa disattivata deve invece:

- rimanere visibile in `Gestione commesse`;
- poter essere cercata nella gestione amministrativa;
- poter essere riattivata in qualsiasi momento;
- conservare tutti i documenti, impianti, collegamenti, squadre storiche e ore;
- mantenere invariato lo stesso ID Firestore.

## Protezione assoluta dello storico ore

La disattivazione di una commessa non deve mai modificare, eliminare, nascondere o scollegare:

- `oreReports`;
- `oreApprovalRequests`;
- calendario personale;
- riepiloghi mensili già pubblicati;
- viste condivise del calendario;
- ore approvate, rifiutate, annullate o in attesa;
- associazione storica tra ora, operatore e commessa.

Le query e i filtri del calendario personale non devono usare `commessa.attiva === true` come condizione.

Il calendario deve continuare a mostrare le ore di tutte le commesse, comprese quelle disattivate.

Esempio obbligatorio:

```text
Commessa HERA MODENA: attiva = false
04/08/2026 · HERA MODENA · 8 ore
```

La voce deve restare visibile nel calendario personale.

Quando possibile, ogni record ore deve visualizzare il nome già salvato nel record o nella vista condivisa. Se serve risolvere il nome tramite ID, il resolver deve poter usare anche commesse disattivate, senza riattivare i listener degli impianti.

## Riduzione delle letture

Il caricamento iniziale deve applicare il filtro di attività prima di registrare i listener degli impianti:

```js
const commesseOperative = commesse.filter((commessa) => commessa.attiva !== false);
```

Aprire listener `commesse/{id}/impianti` esclusivamente per `commesseOperative` realmente necessarie alla schermata corrente.

Non aprire listener degli impianti per le commesse disattivate.

La schermata amministrativa può leggere l'elenco essenziale delle commesse, ma non deve aprire automaticamente tutte le sottocollezioni degli impianti.

## Riattivazione

Quando l'amministratore riattiva una commessa:

- aggiornare `attiva: true` sul documento esistente;
- non creare una nuova commessa;
- non cambiare ID;
- non duplicare impianti;
- non ricreare squadre o ore;
- renderla nuovamente disponibile nelle sezioni operative;
- aprire i listener operativi soltanto quando richiesti dalla schermata.

## Parti protette

Non modificare direttamente o indirettamente:

- pulsante e logica `FATTO`;
- WhatsApp/WHAZZUP;
- salvataggio di data, ora e operatore del completamento;
- struttura o ID degli impianti;
- struttura e ID del personale;
- struttura e ID dei mezzi;
- storico squadre;
- storico ore;
- funzionamento offline esistente.

## Test obbligatori

1. Disattivare una commessa con impianti e ore storiche.
2. Riaprire completamente l'app.
3. Verificare che la commessa non compaia nell'elenco operativo.
4. Verificare che non venga aperto alcun listener `commesse/{id}/impianti` per la commessa disattivata.
5. Aprire il calendario personale e verificare che tutte le ore storiche della commessa siano presenti.
6. Verificare mese precedente, mese corrente e riepilogo mensile.
7. Verificare che personale e mezzi non vengano eliminati o svuotati.
8. Riattivare la stessa commessa.
9. Verificare che ritorni operativa con lo stesso ID e gli stessi impianti.
10. Verificare FATTO e WhatsApp su un impianto della commessa riattivata.
11. Confrontare la diagnostica Firestore prima e dopo: i listener degli impianti delle commesse disattivate devono risultare assenti.

## Criterio di accettazione

La funzione è accettata soltanto se riduce le letture e i listener delle commesse non utilizzate senza modificare o nascondere nemmeno una singola ora storica nel calendario personale.
