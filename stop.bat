@echo off
REM Остановка сервиса подписания актов на Windows.
setlocal enabledelayedexpansion
set FOUND=0

for /f "tokens=2 delims==" %%P in ('wmic process where "CommandLine like '%%uvicorn app.main%%'" get ProcessId /value 2^>nul ^| find "ProcessId"') do (
  taskkill /PID %%P /F >nul 2>&1
  set FOUND=1
)

if "%FOUND%"=="1" (
  echo Сервис остановлен.
) else (
  echo Сервис не запущен.
)
pause
