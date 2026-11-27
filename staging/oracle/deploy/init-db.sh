#!/bin/bash
# Executed/sourced only by official PostgreSQL image on an EMPTY data volume.
# Do not use against existing databases.
(
set -euo pipefail
[ "${POSTGRES_DB:-}" = varga_v2_oracle ] || { echo "Wrong database"; exit 1; }
export VARGA_APP_PASSWORD="$(cat /run/secrets/app_db_password)"
[ "${#VARGA_APP_PASSWORD}" -ge 32 ] || { echo "App password too short"; exit 1; }
psql --username postgres --dbname varga_v2_oracle -v ON_ERROR_STOP=1 <<'SQL'
\getenv app_password VARGA_APP_PASSWORD
CREATE ROLE varga_app LOGIN PASSWORD :'app_password' NOSUPERUSER NOBYPASSRLS NOCREATEDB NOCREATEROLE;
SQL
unset VARGA_APP_PASSWORD
for file in /opt/varga-migrations/001_initial.sql /opt/varga-migrations/002_plan_management.sql; do
  psql --username postgres --dbname varga_v2_oracle -v ON_ERROR_STOP=1 -f "$file"
done
)
