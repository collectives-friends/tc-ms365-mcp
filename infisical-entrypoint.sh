#!/bin/sh
set -e
export INFISICAL_TOKEN="$(infisical login --method=universal-auth \
  --client-id="$INFISICAL_CLIENT_ID" --client-secret="$INFISICAL_CLIENT_SECRET" \
  --domain="$INFISICAL_API_URL" --silent --plain)"
exec infisical run --projectId="$INFISICAL_PROJECT_ID" --env="${INFISICAL_ENV:-prod}" \
  --domain="$INFISICAL_API_URL" --path="${INFISICAL_PATH:-/}" -- node dist/index.js "$@"
