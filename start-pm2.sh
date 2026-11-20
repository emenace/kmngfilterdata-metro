#!/usr/bin/env bash
set -euo pipefail

APP_NAME="kmngfilterdata-metro"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

export PORT="${PORT:-3000}"
export MAX_FILE_SIZE_MB="${MAX_FILE_SIZE_MB:-100}"
export JOB_TTL_MINUTES="${JOB_TTL_MINUTES:-60}"

if ! command -v node >/dev/null 2>&1; then
  echo "Error: Node.js is not installed. Please install Node.js >= 22.5.0." >&2
  exit 1
fi

if ! command -v pm2 >/dev/null 2>&1; then
  echo "Error: pm2 is not installed. Install it globally with: npm install -g pm2" >&2
  exit 1
fi

if [ ! -d "node_modules" ]; then
  echo "Installing dependencies..."
  npm install
fi

# Run as a single instance (fork mode) because JobStore tracks active jobs in memory.
if pm2 describe "$APP_NAME" >/dev/null 2>&1; then
  echo "Restarting $APP_NAME on PM2 (port $PORT)..."
  pm2 restart "$APP_NAME" --update-env
else
  echo "Starting $APP_NAME on PM2 (port $PORT)..."
  pm2 start src/server.js --name "$APP_NAME" --instances 1
fi

pm2 status "$APP_NAME"
echo ""
echo "App running at http://localhost:$PORT"
echo "Logs: pm2 logs $APP_NAME"

