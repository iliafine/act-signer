#!/bin/bash
# Запуск сервиса подписания актов на Mac. Можно запускать двойным кликом.
set -e
cd "$(dirname "$0")/backend"

PORT=8743
PIDFILE="../.service.pid"

if [ -f "$PIDFILE" ] && kill -0 "$(cat "$PIDFILE")" 2>/dev/null; then
  echo "Сервис уже запущен (PID $(cat "$PIDFILE")). Открываю браузер..."
  open "http://127.0.0.1:$PORT"
  exit 0
fi

if [ ! -d ".venv" ]; then
  echo "Создаю виртуальное окружение..."
  python3 -m venv .venv
fi

source .venv/bin/activate
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt

nohup python -m uvicorn app.main:app --host 127.0.0.1 --port "$PORT" > ../service.log 2>&1 &
echo $! > "$PIDFILE"

echo "Сервис запущен (PID $(cat "$PIDFILE")). Открываю браузер..."
sleep 2
open "http://127.0.0.1:$PORT"
