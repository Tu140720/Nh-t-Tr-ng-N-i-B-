#!/bin/sh
set -eu

BOOTSTRAP_MARKER="/app/data/.bootstrap-seeded"

if [ -d "/app/bootstrap-data" ] && [ ! -f "$BOOTSTRAP_MARKER" ]; then
  mkdir -p /app/data
  cp -a /app/bootstrap-data/. /app/data/ 2>/dev/null || true
  date -u +"%Y-%m-%dT%H:%M:%SZ" > "$BOOTSTRAP_MARKER"
fi

exec "$@"
