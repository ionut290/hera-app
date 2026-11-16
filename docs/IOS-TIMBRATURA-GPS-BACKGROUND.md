# iOS: promemoria di timbratura GPS anche con Varga Cantieri chiusa

La nuova opzione **GPS iPhone anche con app chiusa** registra sul dispositivo un unico geofence
intorno alle coordinate del magazzino. Il controllo è gestito da Core Location,
non dal JavaScript della pagina: funziona anche quando Varga Cantieri non è in
primo piano e può ricevere eventi quando iOS rilancia l'app per una transizione.
Nessun percorso, tracciamento GPS continuo, coordinate o presenze sono inviati a un server.

## Attivazione e funzionamento
1. Installare la nuova build iOS e aprire «Ricorda di timbrare».
2. Inserire latitudine, longitudine, raggio GPS (100–2000 m, consigliati 150–300 m),
   giorni e fasce Entrata e Uscita.
3. Attivare **GPS iPhone anche con app chiusa** e premere «Salva e attiva».
4. Consentire notifiche e localizzazione **Sempre** nelle Impostazioni iPhone.
5. Quando il telefono *entra* nella regione, iOS consegna un evento nativo: l'app
   verifica il giorno e la fascia oraria locale prima di inviare una notifica sonora.
6. Toccando la notifica si apre Varga Cantieri con i comandi per raggiungere Easy-Box.

Non vengono eseguite timbrature automatiche. Ogni tipologia di promemoria viene
notificata al massimo una volta al giorno. Disattivare la modalità GPS rimuove
il geofence e la configurazione dal dispositivo. I promemoria ad orario fisso
sono un'alternativa **senza GPS** e vengono disattivati scegliendo il GPS.

## Limiti importanti iOS

- Il geofence segnala *l'attraversamento* dell'area, non fa controlli GPS continui
  al minuto: se si è già dentro l'area quando comincia la fascia, la sola
  variazione dell'orario non provoca una notifica.
- Eventi GPS non sono garantiti immediati; il raggio ha una tolleranza stabilita
  dal sistema e il permesso Sempre è necessario. Eventi possono non essere
  consegnati dopo chiusura forzata, con localizzazione o aggiornamento app in
  background disattivati e in condizioni di risparmio energia.
- Allarmi sono **notifiche locali standard**, non Critical Alerts; Silenzioso o
  Full Immersion possono sopprimerne il suono.
- Il plugin native è compilato **solo** sulle app iOS tramite script
  `scripts/install-ios-native-geofence.rb`; web e Android restano invariati.
- Per verificare l'effettivo funzionamento con app chiusa serve una prova
  sul dispositivo reale con un attraversamento fisico del confine GPS.
- Test statici: `npm run check:ios-geofence`; test di compilazione iOS su PR.

## Privacy e permessi

- Il consenso è opt-in; il geofence non viene registrato se la posizione non
  è autorizzata Sempre. Gli orari e coordinate restano sul dispositivo.
- AppDelegate crea il CLLocationManager prima della WebView in caso di
  rilancio Core Location; la configurazione è salvata in UserDefaults.
- Info.plist dichiara NSLocationAlwaysAndWhenInUseUsageDescription con il
  motivo preciso. Non è necessaria una sessione di navigazione GPS continua.

La build non si distribuisce automaticamente: serve processo manuale TestFlight
interno autorizzato con SHA esatto dopo esito verde CI.


## Apertura Easy-Box al tocco della notifica (opt-in)

**Il collegamento diretto all'app Easy-Box di ICT-GROUP non risulta documentato pubblicamente.**
Non inventare URL `easybox://` o schemi equivalenti. Varga Cantieri usa il metodo
documentato da Apple per eseguire un Comando Rapido tramite
`shortcuts://run-shortcut?name=Apri%20Easy-Box`.

Preparazione **una volta sola sull'iPhone**:
1. Installa Easy-Box e l'app Apple **Comandi Rapidi**.
2. Comandi Rapidi → **+** → aggiungi azione **Apri app**, seleziona **Easy-Box**.
3. Rinomina il comando esattamente **Apri Easy-Box**.
4. In Varga Cantieri → Ricorda di timbrare → **Prova apertura Easy-Box**:
   verifica che il comando apra davvero Easy-Box. Se manca o ha un altro nome,
   iOS può soltanto aprire Comandi Rapidi senza aprire Easy-Box.
5. Attiva la casella **Apri subito Easy-Box quando tocco la notifica** e salva.

Quando il proprietario **tocca la notifica** (o il pulsante Apri timbratura),
il plugin Capacitor LocalNotifications consegna l'azione; Varga Cantieri
richiede al nativo iOS di aprire il link Shortcuts. L'app Varga Cantieri deve
essere attivata da iOS per ricevere il tocco, anche se l'obiettivo è
passare subito a Easy-Box. Questo passaggio o una conferma possono essere
visibili brevemente; non è possibile garantire che non si vedano mai.

Se il collegamento non può aprirsi viene mostrata la schermata già esistente
con i pulsanti manuali Easy-Box / App Store. L'opzione è disattivata di default
e non cambia le timbrature: ogni registrazione resta manuale.

**Sicurezza:** notifiche di altre funzionalità, notifiche eliminate/dismiss,
semplice ingresso GPS e riapertura ordinaria di Varga Cantieri non possono
causare l'apertura automatica di Easy-Box. È necessario un vero tocco
dell'utente su una notifica di timbratura e il consenso nell'interfaccia.
Verificare su iPhone reale perché l'app terza, Comandi Rapidi e la gestione
delle notifiche dipendono da iOS e dal dispositivo.
