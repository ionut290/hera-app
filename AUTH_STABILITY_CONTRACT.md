# Contratto di stabilità del login

Identificatore: `LOGIN_STABILITY_CONTRACT_V1`.

## Obiettivo

Il login deve terminare sempre in uno stato visibile e comprensibile:

1. schermata di login;
2. caricamento o verifica autorizzazione;
3. Home autenticata;
4. errore recuperabile con possibilità di riprovare;
5. accesso negato o sessione scaduta.

Non è consentito lasciare contemporaneamente nascosti login, caricamento, verifica autorizzazione e Home.

## Regole operative

- Firebase Authentication resta l'unica autorità per la sessione.
- Il caricamento del profilo o di Firestore non deve impedire la notifica dello stato autenticato.
- Commesse e squadra del giorno sono dati iniziali prioritari.
- Meteo, distributori, documenti e altre funzioni secondarie non devono bloccare la Home.
- La persistenza locale della sessione deve essere configurata prima del login.
- Errori e timeout devono mostrare una superficie visibile; non devono produrre una schermata bianca.
- La protezione non deve aggiungere letture, scritture o listener Firebase.
- Il flusso deve restare verificato su PWA e Android.

## Modificabilità autorizzata

Questa protezione è forte ma non irrevocabile. Può essere modificata, aggiornata o sostituita esclusivamente dopo il consenso esplicito del proprietario dell'app. Ogni modifica autorizzata deve aggiornare consapevolmente il contratto e i relativi controlli automatici, superare tutti i test critici e preservare i dati esistenti.

Modifiche ad altre sezioni dell'app non possono cambiare indirettamente il comportamento del login.
