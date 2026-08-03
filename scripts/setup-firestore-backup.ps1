param(
  [string]$ProjectId = "hera-app-6cd2b",
  [string]$DatabaseId = "(default)",
  [string]$Retention = "14w"
)

$ErrorActionPreference = "Stop"

function Invoke-GCloud {
  param([Parameter(Mandatory = $true)][string[]]$Arguments)

  & gcloud @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "Comando gcloud non riuscito: gcloud $($Arguments -join ' ')"
  }
}

if (-not (Get-Command gcloud -ErrorAction SilentlyContinue)) {
  throw "Google Cloud CLI (gcloud) non è installato o non è disponibile nel PATH."
}

Write-Host "Verifica autenticazione Google Cloud..."
$activeAccount = (& gcloud auth list --filter=status:ACTIVE --format="value(account)")
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($activeAccount)) {
  throw "Nessun account Google Cloud autenticato. Esegui prima: gcloud auth login"
}

Write-Host "Imposto il progetto $ProjectId..."
Invoke-GCloud -Arguments @("config", "set", "project", $ProjectId)

Write-Host "Abilito l'API Firestore, se necessario..."
Invoke-GCloud -Arguments @("services", "enable", "firestore.googleapis.com", "--project", $ProjectId)

Write-Host "Controllo le pianificazioni di backup esistenti..."
$existingJson = & gcloud firestore backups schedules list --database=$DatabaseId --project=$ProjectId --format=json
if ($LASTEXITCODE -ne 0) {
  throw "Impossibile leggere le pianificazioni di backup. Verifica fatturazione e permessi IAM."
}

$existing = @()
if (-not [string]::IsNullOrWhiteSpace($existingJson)) {
  $existing = @($existingJson | ConvertFrom-Json)
}

$dailyExists = $false
foreach ($schedule in $existing) {
  if ($null -ne $schedule.dailyRecurrence) {
    $dailyExists = $true
    Write-Host "Esiste già un backup giornaliero: $($schedule.name)"
    Write-Host "Conservazione attuale: $($schedule.retention)"
  }
}

if (-not $dailyExists) {
  Write-Host "Creo il backup giornaliero con conservazione $Retention..."
  Invoke-GCloud -Arguments @(
    "firestore", "backups", "schedules", "create",
    "--database=$DatabaseId",
    "--recurrence=daily",
    "--retention=$Retention",
    "--project=$ProjectId"
  )
}

Write-Host ""
Write-Host "Pianificazioni attive:"
Invoke-GCloud -Arguments @(
  "firestore", "backups", "schedules", "list",
  "--database=$DatabaseId",
  "--project=$ProjectId"
)

Write-Host ""
Write-Host "Backup disponibili:"
Invoke-GCloud -Arguments @(
  "firestore", "backups", "list",
  "--project=$ProjectId",
  "--format=table(name,database,state,createTime,expireTime)"
)

Write-Host ""
Write-Host "Configurazione completata. I backup Firestore sono gestiti da Google Cloud e non modificano la logica dell'app."
