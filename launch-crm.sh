#!/bin/bash
CRM_DIR="/Users/kelvinchia/AstuteCorpCRM"
LOG_FILE="$CRM_DIR/instance/crm-server.log"
PID_FILE="$CRM_DIR/instance/crm-server.pid"
URL="http://127.0.0.1:5001"

mkdir -p "$CRM_DIR/instance"

is_running() {
    curl -fsS -o /dev/null "$URL/login" 2>/dev/null
}

open_browser() {
    open "$URL"
}

if is_running; then
    open_browser
    exit 0
fi

if [ -f "$PID_FILE" ]; then
    OLD_PID="$(cat "$PID_FILE")"
    if [ -n "$OLD_PID" ] && kill -0 "$OLD_PID" 2>/dev/null; then
        for _ in {1..20}; do
            if is_running; then
                open_browser
                exit 0
            fi
            sleep 0.5
        done
    fi
fi

cd "$CRM_DIR" || exit 1
source venv/bin/activate
nohup python3 app.py >> "$LOG_FILE" 2>&1 &
echo $! > "$PID_FILE"

for _ in {1..30}; do
    if is_running; then
        open_browser
        exit 0
    fi
    sleep 0.5
done

osascript -e 'display alert "AstuteCorp CRM" message "The server could not start. Check the log file in AstuteCorpCRM/instance/crm-server.log." as critical'
exit 1