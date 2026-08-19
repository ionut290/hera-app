#!/usr/bin/env bash
set -euo pipefail

LABEL="VARGA-STABLE-2026-08-19"
BACKUP_BUCKET="${BACKUP_BUCKET:-gs://hera-app-6cd2b-dr-backups-645390631375}"
ROOT="$BACKUP_BUCKET/stable-2026-08-19/project-buckets"
DRY_RUN="${DRY_RUN:-true}"
CONFIRM_RESTORE="${CONFIRM_RESTORE:-}"
ALLOW_STORAGE_DELETE="${ALLOW_STORAGE_DELETE:-}"

if [ "$DRY_RUN" != "true" ] && [ "$DRY_RUN" != "false" ]; then
  echo "DRY_RUN deve essere true oppure false." >&2
  exit 1
fi

if [ "$DRY_RUN" = "false" ]; then
  if [ "$CONFIRM_RESTORE" != "$LABEL" ]; then
    echo "Ripristino bloccato: manca CONFIRM_RESTORE=$LABEL" >&2
    exit 1
  fi
  if [ "$ALLOW_STORAGE_DELETE" != "DELETE_AND_RESTORE_STORAGE_2026_08_19" ]; then
    echo "Ripristino esatto Storage bloccato: manca ALLOW_STORAGE_DELETE=DELETE_AND_RESTORE_STORAGE_2026_08_19" >&2
    exit 1
  fi
fi

mapfile -t urls < <(gcloud storage ls --recursive "$ROOT/**" 2>/dev/null || true)
if [ "${#urls[@]}" -eq 0 ]; then
  echo "Nessun oggetto Storage trovato in $ROOT" >&2
  exit 1
fi

buckets_file="$(mktemp)"
trap 'rm -f "$buckets_file"' EXIT
for url in "${urls[@]}"; do
  case "$url" in
    "$ROOT"/*)
      relative="${url#"$ROOT"/}"
      bucket="${relative%%/*}"
      [ -n "$bucket" ] && printf '%s\n' "$bucket" >> "$buckets_file"
      ;;
  esac
done
mapfile -t buckets < <(sort -u "$buckets_file")

if [ "${#buckets[@]}" -eq 0 ]; then
  echo "Impossibile ricostruire i nomi dei bucket dal backup." >&2
  exit 1
fi

echo "Bucket da ripristinare: ${#buckets[@]}"
for bucket in "${buckets[@]}"; do
  source="$ROOT/$bucket"
  destination="gs://$bucket"
  echo "- $bucket"
  if [ "$DRY_RUN" = "true" ]; then
    gcloud storage rsync --recursive --dry-run --delete-unmatched-destination-objects "$source" "$destination"
  else
    gcloud storage rsync --recursive --delete-unmatched-destination-objects "$source" "$destination"
  fi
done

echo "Storage restore ${DRY_RUN:+verificato} per $LABEL. DRY_RUN=$DRY_RUN"
