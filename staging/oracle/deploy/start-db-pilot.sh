#!/bin/bash
# First-time, private staging database only. Do not run twice or replace any existing project.
set -euo pipefail
umask 077
cd "$(dirname "$0")"
bash shared-host-preflight.sh
command -v vmstat >/dev/null
command -v openssl >/dev/null
sudo -n docker compose version >/dev/null
# A short new observation gate; exclude the vmstat since-boot first sample.
printf '\nChecking CPU headroom for 30 seconds...\n'
LC_ALL=C vmstat 1 31 | awk 'NR>3 {n++; sum+=$15; if($15<10) busy++} END {if(n!=30) {print "STOP: incomplete CPU sample"; exit 1} printf "CPU idle average %.1f%%; near-saturation samples %d/%d\n",sum/n,busy,n; if(sum/n<35 || busy>6) {print "STOP: insufficient sustained CPU headroom"; exit 1}}'
sudo -n docker image inspect postgres:16 >/dev/null 2>&1 || { echo 'STOP: postgres:16 not cached; do not download on this shared host while load is uncertain'; exit 1; }
[[ "$(sudo -n docker image inspect postgres:16 --format '{{.Architecture}}')" = arm64 ]] || { echo 'STOP: image architecture mismatch'; exit 1; }
[[ -z "$(sudo -n docker network ls -q --filter label=com.docker.compose.project=varga-v2-oracle)" ]] || { echo 'STOP: existing project network'; exit 1; }
[[ -z "$(sudo -n docker ps -aq --filter name='^/varga-v2-oracle-db-1$')" ]] || { echo 'STOP: container name collision'; exit 1; }
[[ -z "$(sudo -n docker volume ls -q --filter name='^varga-v2-oracle_pgdata$')" ]] || { echo 'STOP: volume name collision'; exit 1; }
export VARGA_SECRETS_DIR="$HOME/.varga-v2-oracle-secrets"
[[ ! -e "$VARGA_SECRETS_DIR" ]] || { echo 'STOP: existing credentials; do not overwrite'; exit 1; }
# Obtain the actual PostgreSQL group from the cached image.
pg_gid="$(sudo -n docker run --rm --pull=never --network none --cpus .10 --memory 128m --entrypoint id postgres:16 -g postgres)"
[[ "$pg_gid" =~ ^[0-9]+$ ]] || exit 1
mkdir -m 700 "$VARGA_SECRETS_DIR"
openssl rand -hex 32 > "$VARGA_SECRETS_DIR/db-owner-password"
openssl rand -hex 32 > "$VARGA_SECRETS_DIR/app-db-password"
# Generate DB URL for the future API without displaying passwords.
printf 'postgresql://varga_app:%s@db:5432/varga_v2_oracle' "$(cat "$VARGA_SECRETS_DIR/app-db-password")" > "$VARGA_SECRETS_DIR/app-database-url"
sudo -n chown root:"$pg_gid" "$VARGA_SECRETS_DIR/db-owner-password" "$VARGA_SECRETS_DIR/app-db-password"
sudo -n chmod 440 "$VARGA_SECRETS_DIR/db-owner-password" "$VARGA_SECRETS_DIR/app-db-password"
# Files containing SQL must be readable in container; no secrets are stored with them.
chmod 755 ../migrations
chmod 644 init-db.sh ../migrations/001_initial.sql ../migrations/002_plan_management.sql
c=(sudo -n --preserve-env=VARGA_SECRETS_DIR docker compose -p varga-v2-oracle -f compose.db-pilot.yml)
"${c[@]}" config --quiet
"${c[@]}" up -d --no-build --pull never --wait --wait-timeout 180 db
tables="$("${c[@]}" exec -T db psql -U postgres -d varga_v2_oracle -Atc "SELECT count(*) FROM information_schema.tables WHERE table_schema='public' AND table_name IN ('tenants','members','join_requests','jobs','plants','plan_audit');")"
[[ "$tables" = 6 ]] || { echo 'STOP: incomplete schema; inspect without deleting volumes'; exit 1; }
role="$("${c[@]}" exec -T db psql -U postgres -d varga_v2_oracle -Atc "SELECT rolname,rolsuper,rolbypassrls FROM pg_roles WHERE rolname='varga_app';")"
[[ "$role" = 'varga_app|f|f' ]] || { echo "STOP: unsafe database role"; exit 1; }
printf '%s\n' "$role"
"${c[@]}" ps
printf '\nPASS: six tables initialized, private database only. No API, DNS, firewall or radio changes.\n'
