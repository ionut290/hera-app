#!/bin/bash
# Read-only; does not install, stop, restart, or change any service.
set -euo pipefail
port="${VARGA_API_PORT:-18084}"
[[ "$port" =~ ^[0-9]{1,5}$ ]] && (( 10#$port >= 1024 && 10#$port <= 65535 )) || { echo "Invalid port"; exit 1; }
[[ "$(hostname -s | tr '[:upper:]' '[:lower:]')" = radio-romanii-italia ]] || { echo "STOP: unexpected host"; exit 1; }
printf 'Host: '; hostname -s
printf 'Architecture: '; uname -m
printf 'CPU: '; nproc
uptime
free -h
df -h /
printf '\nDocker containers (names, state, ports only):\n'
if docker info >/dev/null 2>&1; then d=(docker); elif sudo -n docker info >/dev/null 2>&1; then d=(sudo -n docker); else echo 'STOP: Docker unavailable'; exit 1; fi
"${d[@]}" ps -a --format '{{.Names}} | {{.Status}} | {{.Ports}}'
printf '\nListening TCP ports:\n'
ss -H -ltn
if ss -H -ltn | awk -v p="$port" '$4 ~ (":" p "$"){found=1} END{exit !found}'; then echo "STOP: proposed port already in use"; exit 1; fi
if [[ -n "$("${d[@]}" ps -aq --filter label=com.docker.compose.project=varga-v2-oracle)" ]]; then echo "STOP: existing project containers; inspect without replacing"; exit 1; fi
if [[ -n "$("${d[@]}" volume ls -q --filter label=com.docker.compose.project=varga-v2-oracle)" ]]; then echo "STOP: existing project volumes; inspect without replacing"; exit 1; fi
mem="$(awk '/MemAvailable:/{print $2}' /proc/meminfo)"
disk="$(df -Pk / | awk 'NR==2{print $4}')"
(( mem >= 4194304 && disk >= 20971520 )) || { echo "STOP: less than 4 GiB available RAM or 20 GiB root free space"; exit 1; }
printf '\nPASS: basic capacity and port checks only. Check sustained radio performance, Docker disk location and proxy config before deployment.\n'
