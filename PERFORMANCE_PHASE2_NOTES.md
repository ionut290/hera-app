# Performance phase 2

Intervento conservativo lato client.

- Riduce il ticker Home da un render completo ogni minuto a un aggiornamento ai soli cambi di finestra 12:00/19:00.
- Esegue un refresh Home al ritorno in primo piano se la lista commesse è visibile.
- Sospende l'animazione radar mentre la pagina è nascosta.
- Mantiene invariati dati, Firestore, FATTO, WHAZZUP, commesse, impianti, ore, squadre e calendario.
