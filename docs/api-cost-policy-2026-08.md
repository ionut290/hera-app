# Policy controllo costi API

Aggiornata al 23/08/2026.

Il Centro di Controllo distingue tra contatori affidabili per il blocco e metriche indicative. Street View 360° usa un contatore applicativo condiviso e resta bloccato a 4.800 aperture/mese, sotto il free usage cap Google di 5.000 eventi/mese. Per le altre API Google, il monitoraggio Service Runtime è usato per avvisi e confronto con i free usage cap, ma non come hard-stop automatico perché una richiesta API non coincide sempre con un evento fatturabile/SKU.

Soglie di riferimento Google Maps Platform (global pricing, agosto 2026): Dynamic Maps 10.000 eventi gratuiti/mese; Static Street View 10.000; Dynamic Street View 5.000; Places Autocomplete Requests 10.000; Place Details Essentials 10.000; Place Details Pro 5.000. L'app che usa Places richiede anche displayName, quindi può attivare Place Details Pro: il Centro di Controllo usa prudenzialmente 5.000 come riferimento per Places.
