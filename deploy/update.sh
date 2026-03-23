#!/usr/bin/env bash
set -euo pipefail

APP_DIR="/opt/excel_webdatamerger"
SERVICE_NAME="excel_webdatamerger.service"
BRANCH="${1:-main}"

cd "$APP_DIR"

echo ">>> Fetching latest code for $BRANCH..."
git fetch origin "$BRANCH"
git reset --hard "origin/$BRANCH"

if [[ -d venv ]]; then
  echo ">>> Updating dependencies..."
  # shellcheck disable=SC1091
  source venv/bin/activate
  pip install -r requirements.txt
else
  echo ">>> Skip dependency update: venv/ not found"
fi

echo ">>> Restarting service: $SERVICE_NAME"
systemctl restart "$SERVICE_NAME"

echo ">>> Done"
