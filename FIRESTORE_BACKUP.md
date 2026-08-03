# Backup Firestore — Hera App

Progetto Firebase/Google Cloud: `hera-app-6cd2b`

Questa procedura configura il backup nativo giornaliero del database Firestore `(default)` con conservazione di 14 settimane.

## Caratteristiche

- Backup completo del database Firestore.
- Esecuzione gestita da Google Cloud, separata dalla web app.
- Nessun listener, lettura o scrittura aggiunto durante l'uso dell'app.
- Nessuna modifica alle funzioni FATTO o WhatsApp/WHAZZUP.
- I backup contengono dati e configurazioni degli indici presenti nel momento del backup.
- Le regole di sicurezza Firebase e le policy TTL non sono incluse nel backup e restano versionate nel repository.

## Requisiti

1. Fatturazione Google Cloud attiva sul progetto.
2. Google Cloud CLI installata.
3. Account autenticato con permessi per gestire i backup Firestore, ad esempio `roles/datastore.owner` oppure `roles/datastore.backupSchedulesAdmin`.

## Attivazione da Windows PowerShell

Dalla cartella del repository:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\setup-firestore-backup.ps1
```

Lo script:

1. verifica `gcloud` e l'account attivo;
2. seleziona il progetto `hera-app-6cd2b`;
3. abilita l'API Firestore se necessario;
4. controlla se esiste già una pianificazione giornaliera;
5. crea il backup giornaliero solo se assente;
6. mostra pianificazioni e backup disponibili.

Lo script è idempotente: rilanciarlo non crea una seconda pianificazione giornaliera.

## Verifica manuale

```powershell
gcloud firestore backups schedules list --database="(default)" --project="hera-app-6cd2b"

gcloud firestore backups list --project="hera-app-6cd2b" --format="table(name,database,state,createTime,expireTime)"
```

Il primo backup non è necessariamente disponibile subito dopo la creazione della pianificazione. Google Cloud decide l'orario di esecuzione giornaliero.

## Ripristino

Il backup nativo Firestore viene ripristinato in un **nuovo database**, non direttamente sopra il database attivo. Prima di un ripristino bisogna:

1. individuare il backup corretto;
2. creare un nuovo ID database di destinazione;
3. avviare il ripristino;
4. verificare dati, indici e regole;
5. pianificare con attenzione l'eventuale passaggio dell'app al database ripristinato.

Esempio, da completare con posizione e ID reali:

```powershell
gcloud firestore databases restore `
  --source-backup="projects/hera-app-6cd2b/locations/LOCATION/backups/BACKUP_ID" `
  --destination-database="hera-restore-YYYYMMDD" `
  --project="hera-app-6cd2b"
```

Non eseguire un ripristino senza prima verificare dipendenze, regole Firestore, utenti, calendario, ore, agenda, squadre, commesse e impianti.

## Dati protetti

Il backup comprende le collezioni presenti nel database, incluse, se esistenti:

- personale;
- utenti e profili applicativi;
- calendario;
- ore;
- agenda;
- squadre e storico squadre;
- commesse;
- impianti;
- segnalazioni;
- altri documenti Firestore dell'app.

## Costi

Il backup nativo comporta costi di archiviazione e costi di ripristino. La creazione e conservazione dei backup non aggiunge letture o scritture al normale utilizzo dell'app.
