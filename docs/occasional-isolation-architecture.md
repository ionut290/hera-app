# Isolamento Lavori occasionali

I moduli `lavori-occasionali*.js` sono estensioni additive dell'interfaccia e non sono proprietari delle cache globali degli impianti.

Regola obbligatoria: possono leggere `currentImpianti` e `impiantiByCommessaId`, ma non possono assegnare, cancellare o mutare tali strutture. La selezione e la visibilità dei cantieri delle commesse normali restano sotto la logica principale dell'app.

Il controllo automatico `scripts/check-lavori-occasionali-isolation.js` viene eseguito da GitHub Actions per ogni pull request che modifica i moduli Lavori occasionali e blocca il merge in presenza di mutazioni pericolose.
