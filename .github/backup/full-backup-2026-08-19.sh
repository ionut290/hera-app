#!/usr/bin/env bash
set -u

PROJECT_ID="${PROJECT_ID:-hera-app-6cd2b}"
SOURCE_BUCKET="${SOURCE_BUCKET:-gs://hera-app-6cd2b.firebasestorage.app}"
BACKUP_BUCKET="${BACKUP_BUCKET:-gs://hera-app-6cd2b-dr-backups-645390631375}"
STABLE_CODE_COMMIT="${STABLE_CODE_COMMIT:-81ba299a1e4c7fdafdf56c332daa513a89045985}"
STABLE_CODE_BRANCH="${STABLE_CODE_BRANCH:-backup/stable-2026-08-19}"
PREFIX="stable-2026-08-19/run-${GITHUB_RUN_ID:-manual}"
FIRESTORE_URI="${BACKUP_BUCKET}/${PREFIX}/firestore"
STORAGE_URI="${BACKUP_BUCKET}/${PREFIX}/storage"
META_URI="${BACKUP_BUCKET}/${PREFIX}/metadata"
TMP_DIR="${RUNNER_TEMP:-/tmp}/varga-stable-backup"
mkdir -p "$TMP_DIR"

bucket_status="pending"
firestore_status="pending"
storage_status="pending"
auth_status="pending"
rtdb_status="pending"
config_status="pending"

set_output() {
  local key="$1" value="$2"
  if [ -n "${GITHUB_OUTPUT:-}" ]; then
    printf '%s=%s\n' "$key" "$value" >> "$GITHUB_OUTPUT"
  fi
}

finish_outputs() {
  set_output prefix "$PREFIX"
  set_output bucket_status "$bucket_status"
  set_output firestore_status "$firestore_status"
  set_output storage_status "$storage_status"
  set_output auth_status "$auth_status"
  set_output rtdb_status "$rtdb_status"
  set_output config_status "$config_status"
}
trap finish_outputs EXIT

echo "[backup] progetto: $PROJECT_ID"
echo "[backup] codice stabile: $STABLE_CODE_COMMIT"
gcloud config set project "$PROJECT_ID" >/dev/null

gcloud firestore databases describe --database='(default)' --format=json > "$TMP_DIR/firestore-database.json" 2>"$TMP_DIR/firestore-database.err" || true
source_location="$(gcloud storage buckets describe "$SOURCE_BUCKET" --format='value(location)' 2>/dev/null)"
if [ -z "$source_location" ]; then source_location="EU"; fi

echo "[backup] preparo bucket disaster recovery"
if gcloud storage buckets describe "$BACKUP_BUCKET" >/dev/null 2>&1; then
  bucket_status="existing"
elif gcloud storage buckets create "$BACKUP_BUCKET" --location="$source_location" --uniform-bucket-level-access --project="$PROJECT_ID" >/dev/null 2>"$TMP_DIR/bucket-create.err"; then
  bucket_status="created"
else
  bucket_status="failed"
fi

if [ "$bucket_status" = "failed" ]; then
  echo "[backup] ERRORE: bucket disaster recovery non disponibile"
  exit 1
fi

project_number="$(gcloud projects describe "$PROJECT_ID" --format='value(projectNumber)' 2>/dev/null)"
if [ -n "$project_number" ]; then
  firestore_agent="service-${project_number}@gcp-sa-firestore.iam.gserviceaccount.com"
  gcloud storage buckets add-iam-policy-binding "$BACKUP_BUCKET" \
    --member="serviceAccount:${firestore_agent}" \
    --role="roles/storage.admin" --quiet >/dev/null 2>&1 || true
fi
legacy_agent="${PROJECT_ID}@appspot.gserviceaccount.com"
if gcloud iam service-accounts describe "$legacy_agent" --project="$PROJECT_ID" >/dev/null 2>&1; then
  gcloud storage buckets add-iam-policy-binding "$BACKUP_BUCKET" \
    --member="serviceAccount:${legacy_agent}" \
    --role="roles/storage.admin" --quiet >/dev/null 2>&1 || true
fi

echo "[backup] esportazione Firestore"
if gcloud firestore export "$FIRESTORE_URI" --database='(default)' --project="$PROJECT_ID" --quiet >"$TMP_DIR/firestore-export.log" 2>"$TMP_DIR/firestore-export.err"; then
  firestore_status="success"
else
  firestore_status="failed"
fi

echo "[backup] copia Firebase Storage"
if gcloud storage rsync --recursive "$SOURCE_BUCKET" "$STORAGE_URI" >"$TMP_DIR/storage-rsync.log" 2>"$TMP_DIR/storage-rsync.err"; then
  storage_status="success"
else
  storage_status="failed"
fi

echo "[backup] esportazione Firebase Authentication"
if npx --yes firebase-tools@latest auth:export "$TMP_DIR/auth-users.json" \
  --format=json --project="$PROJECT_ID" --non-interactive >"$TMP_DIR/auth-export.log" 2>"$TMP_DIR/auth-export.err"; then
  if gcloud storage cp "$TMP_DIR/auth-users.json" "$META_URI/auth-users.json" >/dev/null 2>"$TMP_DIR/auth-copy.err"; then
    auth_status="success"
  else
    auth_status="failed-copy"
  fi
else
  auth_status="failed"
fi

echo "[backup] verifica Realtime Database"
if npx --yes firebase-tools@latest database:instances:list --project="$PROJECT_ID" --json --non-interactive > "$TMP_DIR/rtdb-instances.json" 2>"$TMP_DIR/rtdb-list.err"; then
  node - "$TMP_DIR/rtdb-instances.json" "$TMP_DIR/rtdb-instance-ids.txt" <<'NODE'
const fs = require('fs');
let parsed = {};
try { parsed = JSON.parse(fs.readFileSync(process.argv[2], 'utf8')); } catch (_) {}
const rows = Array.isArray(parsed) ? parsed : Array.isArray(parsed.result) ? parsed.result : Array.isArray(parsed.instances) ? parsed.instances : [];
const ids = rows.map((row) => {
  const raw = String(row?.instance || row?.instanceId || row?.name || row?.databaseUrl || '').trim();
  if (!raw) return '';
  if (raw.includes('/')) return raw.split('/').filter(Boolean).pop();
  if (raw.includes('.firebaseio.com')) return raw.replace(/^https?:\/\//, '').split('.')[0];
  return raw;
}).filter(Boolean);
fs.writeFileSync(process.argv[3], [...new Set(ids)].join('\n'));
NODE
  if [ -s "$TMP_DIR/rtdb-instance-ids.txt" ]; then
    rtdb_status="success"
    while IFS= read -r instance; do
      [ -z "$instance" ] && continue
      safe_instance="$(printf '%s' "$instance" | tr -cd 'A-Za-z0-9_-')"
      if ! npx --yes firebase-tools@latest database:get / --project="$PROJECT_ID" --instance="$instance" --pretty --non-interactive > "$TMP_DIR/rtdb-${safe_instance}.json" 2>"$TMP_DIR/rtdb-${safe_instance}.err"; then
        rtdb_status="partial"
        continue
      fi
      if ! gcloud storage cp "$TMP_DIR/rtdb-${safe_instance}.json" "$META_URI/realtime-database/${safe_instance}.json" >/dev/null 2>&1; then
        rtdb_status="partial"
      fi
    done < "$TMP_DIR/rtdb-instance-ids.txt"
  else
    rtdb_status="not-configured"
  fi
else
  rtdb_status="not-configured-or-inaccessible"
fi

echo "[backup] archiviazione configurazione"
config_files=(firestore.rules firestore.indexes.json storage.rules firebase.json firebase-config.js package.json .firebaserc)
present_files=()
for file in "${config_files[@]}"; do
  [ -f "$file" ] && present_files+=("$file")
done
if [ "${#present_files[@]}" -gt 0 ]; then
  tar -czf "$TMP_DIR/app-config.tar.gz" "${present_files[@]}"
  if gcloud storage cp "$TMP_DIR/app-config.tar.gz" "$META_URI/app-config.tar.gz" >/dev/null 2>"$TMP_DIR/config-copy.err"; then
    config_status="success"
  else
    config_status="failed"
  fi
else
  config_status="failed"
fi

node - "$TMP_DIR/backup-manifest.json" <<NODE
const fs = require('fs');
fs.writeFileSync(process.argv[2], JSON.stringify({
  schemaVersion: 1,
  label: 'VARGA-STABLE-2026-08-19',
  createdAt: new Date().toISOString(),
  projectId: '$PROJECT_ID',
  stableCodeBranch: '$STABLE_CODE_BRANCH',
  stableCodeCommit: '$STABLE_CODE_COMMIT',
  backupBucket: '$BACKUP_BUCKET',
  prefix: '$PREFIX',
  firestoreUri: '$FIRESTORE_URI',
  storageUri: '$STORAGE_URI',
  metadataUri: '$META_URI',
  runId: '${GITHUB_RUN_ID:-manual}',
  statuses: {
    backupBucket: '$bucket_status',
    firestore: '$firestore_status',
    storage: '$storage_status',
    authentication: '$auth_status',
    realtimeDatabase: '$rtdb_status',
    configuration: '$config_status'
  }
}, null, 2) + '\n');
NODE

gcloud storage cp "$TMP_DIR/backup-manifest.json" "$META_URI/backup-manifest.json" >/dev/null 2>&1 || true
gcloud storage cp "$TMP_DIR/firestore-database.json" "$META_URI/firestore-database.json" >/dev/null 2>&1 || true

echo "[backup] RISULTATI"
echo "bucket=$bucket_status firestore=$firestore_status storage=$storage_status auth=$auth_status rtdb=$rtdb_status config=$config_status"

if [ "$firestore_status" != "success" ] || [ "$storage_status" != "success" ] || [ "$auth_status" != "success" ] || [ "$config_status" != "success" ]; then
  exit 1
fi
