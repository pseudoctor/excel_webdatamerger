#!/usr/bin/env bash
set -euo pipefail

APP_DIR="/opt/excel_webdatamerger"
SERVICE_NAME="excel_webdatamerger.service"
SERVICE_USER="www-data"
UPLOAD_ROOT="/tmp/excel_webdatamerger"
LOG_DIR="$APP_DIR/logs"
BRANCH="${1:-main}"

cd "$APP_DIR"

echo ">>> Fetching latest code for $BRANCH..."
git fetch origin "$BRANCH"
git reset --hard "origin/$BRANCH"

if [[ ! -d venv ]]; then
  echo ">>> Creating virtualenv..."
  python3 -m venv venv
fi

echo ">>> Updating dependencies..."
# shellcheck disable=SC1091
source venv/bin/activate
pip install -r requirements.txt

echo ">>> Ensuring runtime directories exist..."
install -d -m 0755 "$UPLOAD_ROOT" "$LOG_DIR"
chown "$SERVICE_USER:$SERVICE_USER" "$UPLOAD_ROOT" "$LOG_DIR"

echo ">>> Restarting service: $SERVICE_NAME"
systemctl restart "$SERVICE_NAME"

echo ">>> Done"
