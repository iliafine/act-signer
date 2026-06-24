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
REM Апгрейд pip необязателен и может падать за прокси — не считаем это ошибкой.
python -m pip install --upgrade pip >nul 2>&1
echo Устанавливаю зависимости (первый раз — пару минут)...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo [ОШИБКА] Не удалось установить зависимости.
  echo Скорее всего мешает прокси/сеть. Попробуйте:
  echo   1^) Раздать интернет с телефона ^(мобильная точка^) и запустить заново.
  echo   2^) Либо открыть PowerShell в этой папке и выполнить:
  echo      $env:HTTP_PROXY=""; $env:HTTPS_PROXY=""; backend\.venv\Scripts\python -m pip install -r backend\requirements.txt
  echo.
  pause
  exit /b 1
)

start "Сервис подписания актов" /min cmd /c "python -m uvicorn app.main:app --host 127.0.0.1 --port %PORT% > ..\service.log 2>&1"

timeout /t 2 /nobreak >nul
start http://127.0.0.1:%PORT%
echo Сервис запущен. Окно можно закрыть, сервис продолжит работать в фоне.
echo Для остановки используйте stop.bat
pause
