# Varga Cantieri iOS — Ricorda di timbrare anche con l'app chiusa

## Funzione

La sezione Personale → **Ricorda di timbrare** ora contiene due opzioni indipendenti:
- **GPS mentre l'app è aperta**: comportamento originale, confronto posizione magazzino e fascia oraria.
- **Notifiche settimanali native iPhone**: suonano all'ora iniziale impostata per entrata e uscita nei giorni selezionati, senza GPS e senza tenere aperta Varga Cantieri.

Orari iniziali suggeriti, modificabili dall'utente: **06:30** e **15:15**, lunedì–venerdì. Le notifiche sono **disattivate finché l'utente non le abilita**. I permessi di iOS sono richiesti quando salva con l'opzione attivata. Modificando giorni/orari la pianificazione viene sostituita; disattivando l'opzione tutte e solo le notifiche di timbratura di Varga vengono cancellate.

Sono promemoria sonori **standard**: Modalità silenziosa, Full Immersion, disabilitazione notifiche e politiche iOS possono attenuare o sopprimere il suono. Non sono allarmi critici e non certificano la presenza in magazzino.

## Apertura Easy-Box

L'app aziendale richiesta è **Easy-Box**, ID App Store **1612276209** (ICT-GROUP).

Toccando la notifica, iOS apre Varga Cantieri, che mostra le azioni:
1. **Apri Easy-Box con Comandi**: funziona se l'utente crea nell'app iOS *Comandi Rapidi* un comando denominato esattamente **Apri Easy-Box**, con azione **Apri app → Easy-Box**.
2. **Easy-Box su App Store**: link ufficiale alla pagina dell'app, da cui l'utente può scegliere **Apri** quando l'app è installata.

Nessuna timbratura automatica, nessuna credenziale condivisa e nessun URL scheme Easy-Box inventato. Un'app iOS non può aprire liberamente altre app quando arriva una notifica senza interazione dell'utente.

## Architettura e test

- Plugin `@capacitor/local-notifications@7.0.7`, compatibile Capacitor 7.
- 14 ID riservati (`620100`–`620113`), al massimo 2 × 7 notifiche ricorrenti per settimana.
- `schedule.on.weekday` con conversione JS domenica=0 → iOS domenica=1.
- Ricezione `localNotificationActionPerformed` e pulsante «Apri timbratura».
- Le preferenze restano nel localStorage **del dispositivo**; non si creano nuovi dati lato server.
- Test statici `npm run check:timbratura-ios`; build simulator workflow su PR.
- Non modifica turni, ore reali, contabilità, GPS e sistemi di presenza Easy-Box.

La nuova funzionalità richiede la compilazione e installazione di una **nuova build nativa iOS** contenente il plugin; non viene attivata da un semplice refresh del sito. La distribuzione pubblica e TestFlight non sono automatiche.
