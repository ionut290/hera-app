#!/bin/bash
# Local, non-destructive backup. Off-host encrypted copy and restore drill required.
set -euo pipefail
umask 077
cd "$(dirname "$0")"
: "${VARGA_SECRETS_DIR:?set absolute secrets directory}"
: "${VARGA_BACKUP_DIR:?set absolute backup directory}"
[[ "$VARGA_BACKUP_DIR" = /* && "$VARGA_SECRETS_DIR" = /* ]] || exit 1
mkdir -p "$VARGA_BACKUP_DIR"
out="$VARGA_BACKUP_DIR/varga-v2-$(date -u +%Y%m%dT%H%M%SZ).dump"
tmp="$(mktemp "$VARGA_BACKUP_DIR/.varga-XXXXXX")"
trap 'rm -f "$tmp"' EXIT
docker compose -f compose.yml exec -T db pg_dump -U postgres -d varga_v2_oracle -Fc > "$tmp"
test -s "$tmp"
docker compose -f compose.yml exec -T db pg_restore --list < "$tmp" >/dev/null
ln "$tmp" "$out"
printf 'Backup verified (archive index): %s\n' "$out"
