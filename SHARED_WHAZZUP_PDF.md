# PDF condivisi Whazzup

I PDF aggiunti dal selettore allegati degli impianti sono condivisi tra tutti gli utenti autenticati.

- Il file binario viene salvato in Firebase Storage nel percorso già autorizzato `documents/{ownerUid}/{documentId}/{fileName}`.
- I metadati vengono salvati nella raccolta Firestore `documents`, con `source = whazzup-impianto-pdf`, `impiantoKey`, eventuale `commessaId` e scadenza a 30 giorni.
- I PDF scaduti non vengono mostrati dal client.
- La funzione schedulata `cleanupExpiredWhazzupPdfs` elimina ogni giorno i file Storage e i relativi record Firestore scaduti.
- Le foto Whazzup locali restano invariate.
