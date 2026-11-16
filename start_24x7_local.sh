#!/usr/bin/env bash
set -euo pipefail

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
export PYTHONPATH="/Users/sparshmehta/.gemini/antigravity/scratch/python_packages:$APP_DIR:${PYTHONPATH:-}"
CLOUDFLARED_BIN="/Users/sparshmehta/.gemini/antigravity/scratch/bin/cloudflared"

echo "Starting GGL-Recon persistent local + Cloudflare Tunnel runner..."
echo "Press Ctrl+C to stop both server and tunnel."

cleanup() {
  echo "Stopping GGL-Recon services..."
  pkill -P $$ 2>/dev/null || true
}
trap cleanup EXIT INT TERM

# Keep server alive in a restart loop
(
  while true; do
    echo "[$(date)] Launching server.py on port 8085..."
    python3 "$APP_DIR/server.py" || true
    echo "[$(date)] server.py exited, restarting in 2s..."
    sleep 2
  done
) &

sleep 2

# Keep Cloudflare quick tunnel alive in a restart loop
if [ -x "$CLOUDFLARED_BIN" ]; then
  while true; do
    echo "[$(date)] Launching Cloudflare tunnel..."
    "$CLOUDFLARED_BIN" tunnel --url http://127.0.0.1:8085 || true
    echo "[$(date)] cloudflared exited, reconnecting in 3s..."
    sleep 3
  done
else
  echo "cloudflared binary not found at $CLOUDFLARED_BIN"
  wait
fi
