# Verifica matrice prezziario

Controlli eseguiti prima della pull request:

- `node --check preventivi-price-list-matrix.js`;
- controllo sintattico del loader `preventivi-storage-config.js`;
- test unitario del calcolo ribasso: 100,00 € con ribasso 15% = 85,00 €;
- test del parsing italiano: 12,50 € con ribasso 20% = 10,00 €;
- test rilevamento codici prezzo duplicati senza distinzione tra maiuscole/minuscole;
- test blocco ribasso fuori dall'intervallo 0–100;
- verifica della matrice `.xlsx` con intestazioni, area compilabile, formati e scheda istruzioni;
- verifica assenza di errori formula nel file Excel generato.

Il file è una nota di collaudo della funzione e non contiene dati operativi dell'app.
