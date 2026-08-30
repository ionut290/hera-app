# Assistenti Verde e Mezzi — configurazione gratuita

Le integrazioni sono inattive finché le chiavi non vengono aggiunte come variabili protette del sito Netlify. Le chiavi non devono essere inserite in `index.html`, `green-assistant.js` o in altri file pubblici.

## Variabili Netlify

| Variabile | Servizio | Uso |
| --- | --- | --- |
| `PLANTNET_API_KEY` | Pl@ntNet | Identificazione piante e patologie coperte |
| `TREFLE_API_TOKEN` | Trefle | Ricerca e schede botaniche |
| `GEMINI_API_KEY` | Gemini Developer API | Schede prudenti di mezzi e utensili |
| `GEMINI_MODEL` | Gemini Developer API | Facoltativa; predefinito `gemini-3.5-flash-lite` |
| `BRAVE_SEARCH_API_KEY` | Brave Search API | Ricerca di manuali, schede tecniche, manutenzione e ricambi |

Procedura:

1. Creare gli account gratuiti sui portali ufficiali Pl@ntNet, Trefle e Google AI Studio.
2. Per Gemini usare un progetto senza fatturazione collegata. Il codice non abilita Grounding, Search o altri strumenti a pagamento.
3. In Netlify aprire **Site configuration → Environment variables** e aggiungere le variabili dei servizi che si vogliono utilizzare.
4. Eseguire un nuovo deploy del sito.
5. Aprire uno dei due assistenti e controllare la riga di stato dei servizi.

## Garanzie e limiti

- Non esiste un passaggio automatico a servizi a pagamento.
- Quando termina una quota gratuita, l’utente riceve un errore e deve attendere il ripristino della quota.
- La gratuità di Gemini dipende anche dal progetto Google associato alla chiave: mantenere la fatturazione disattivata.
- Senza testo del manuale ufficiale, miscela, olio, capacità, ricambi e intervalli sono marcati **Da verificare**.
- Le fotografie inviate vengono compresse nel browser e inoltrate al solo servizio scelto; non vengono salvate in Firestore dall’assistente.
- Le schede salvate dall’utente restano nel `localStorage` del dispositivo.
- La ricerca Brave parte soltanto quando l'utente richiede la scheda del mezzo; i risultati vengono conservati sul dispositivo per evitare richieste ripetute.
- Il codice mezzo (per esempio `R50`) viene risolto usando `mezziRecords` già caricato dall'app e non introduce letture Firestore.
- Impostare nel pannello Brave un limite di utilizzo/spesa compatibile con il credito mensile desiderato: l'app non può modificare la fatturazione dell'account API.

Riferimenti ufficiali: [Pl@ntNet API](https://my.plantnet.org/), [Trefle API](https://docs.trefle.io/), [Gemini API](https://ai.google.dev/gemini-api/docs/).
