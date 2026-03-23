#!/usr/bin/env bash
set -euo pipefail

HOST_CERT_ROOT="/home/web/certs"
LE_ROOT="/etc/letsencrypt/live"
NGINX_CONTAINER="nginx"

sync_cert_dir() {
  local domain="$1"
  local target_dir="$HOST_CERT_ROOT/$domain"

  if [[ ! -d "$LE_ROOT/$domain" ]]; then
    echo "skip: missing letsencrypt dir for $domain"
    return 0
  fi

  mkdir -p "$target_dir"
  install -m 0644 "$LE_ROOT/$domain/fullchain.pem" "$target_dir/fullchain.pem"
  install -m 0600 "$LE_ROOT/$domain/privkey.pem" "$target_dir/privkey.pem"
  echo "synced: $domain -> $target_dir"
}

sync_cert_file_pair() {
  local domain="$1"
  local target_cert="$2"
  local target_key="$3"

  if [[ ! -d "$LE_ROOT/$domain" ]]; then
    echo "skip: missing letsencrypt dir for $domain"
    return 0
  fi

  install -m 0644 "$LE_ROOT/$domain/fullchain.pem" "$target_cert"
  install -m 0600 "$LE_ROOT/$domain/privkey.pem" "$target_key"
  echo "synced: $domain -> $target_cert / $target_key"
}

sync_cert_dir "excel.freegarden.dpdns.org"
sync_cert_file_pair \
  "n8n.freegarden.dpdns.org" \
  "$HOST_CERT_ROOT/n8n.freegarden.dpdns.org_cert.pem" \
  "$HOST_CERT_ROOT/n8n.freegarden.dpdns.org_key.pem"

docker exec "$NGINX_CONTAINER" nginx -s reload
echo "reloaded: $NGINX_CONTAINER"
