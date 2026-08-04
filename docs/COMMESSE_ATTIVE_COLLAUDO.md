# Collaudo commesse attive

## Prima del merge

1. Aprire Gestione commesse come amministratore.
2. Verificare la nuova sezione "Commesse caricate all'avvio".
3. Disattivare una commessa di prova che contiene impianti e ore storiche.
4. Attendere il ricaricamento automatico.
5. Verificare che la commessa non compaia nella home operativa.
6. Esportare la diagnostica Firestore e verificare che non venga aperto il listener `commesse/{id}/impianti` della commessa disattivata.
7. Aprire il calendario personale e verificare che le ore storiche della commessa disattivata siano ancora visibili.
8. Verificare che nessun documento della commessa, impianto, ora, squadra, personale o mezzo sia stato eliminato o spostato.
9. Riattivare la commessa.
10. Verificare che torni con lo stesso ID e gli stessi dati.

## Controlli obbligatori

- FATTO invariato.
- WHAZZUP invariato.
- Calendario personale completo.
- Ore storiche complete.
- Squadre storiche complete.
- Personale e mezzi invariati.
- Nessuna cancellazione Firestore.
- Una sola lettura del documento `appConfig/activeCommesse` all'avvio.

## Comportamento di sicurezza

Se il documento `appConfig/activeCommesse` manca, è illeggibile o non contiene un array `ids`, tutte le commesse restano attive. L'ottimizzazione non deve mai nascondere commesse per errore in caso di problemi di rete o permessi.
