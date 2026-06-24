@echo off
REM Запуск сервиса подписания актов на Windows. Запускать двойным кликом.
setlocal
set PORT=8743
cd /d "%~dp0backend"

if not exist .venv (
  echo Создаю виртуальное окружение...
  py -3 -m venv .venv
)

call .venv\Scripts\activate.bat
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt

start "Сервис подписания актов" /min cmd /c "python -m uvicorn app.main:app --host 127.0.0.1 --port %PORT% > ..\service.log 2>&1"

timeout /t 2 /nobreak >nul
start http://127.0.0.1:%PORT%
echo Сервис запущен. Окно можно закрыть, сервис продолжит работать в фоне.
echo Для остановки используйте stop.bat
pause
