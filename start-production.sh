#!/bin/bash
# Run AstuteCorp CRM in production mode (for Cloudflare Tunnel or VPS)
CRM_DIR="/Users/kelvinchia/AstuteCorpCRM"
LOG_FILE="$CRM_DIR/instance/crm-server.log"
PID_FILE="$CRM_DIR/instance/crm-server.pid"
PORT="${PORT:-5001}"

cd "$CRM_DIR" || exit 1
mkdir -p instance
source venv/bin/activate

if [ -f "$PID_FILE" ]; then
    OLD_PID="$(cat "$PID_FILE")"
    if [ -n "$OLD_PID" ] && kill -0 "$OLD_PID" 2>/dev/null; then
        echo "CRM already running (PID $OLD_PID) on port $PORT"
        exit 0
    fi
fi

export FLASK_DEBUG=0
nohup gunicorn -w 1 --timeout 120 -b "127.0.0.1:$PORT" app:app >> "$LOG_FILE" 2>&1 &
echo $! > "$PID_FILE"
echo "AstuteCorp CRM started on http://127.0.0.1:$PORT (PID $(cat "$PID_FILE"))"
echo "Log: $LOG_FILE"