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
