#!/bin/bash
# Остановка сервиса подписания актов на Mac.
cd "$(dirname "$0")"
PIDFILE=".service.pid"

if [ -f "$PIDFILE" ] && kill -0 "$(cat "$PIDFILE")" 2>/dev/null; then
  kill "$(cat "$PIDFILE")"
  rm -f "$PIDFILE"
  echo "Сервис остановлен."
else
  echo "Сервис не запущен."
fi
