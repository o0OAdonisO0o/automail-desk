@echo off
setlocal enabledelayedexpansion
title Automail Desk
cd /d "%~dp0"

set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY (
  where py >nul 2>&1 && set "PY=py -3"
)
if not defined PY (
  echo [ERROR] Python 3.9+ was not found on PATH.
  echo         Install from https://www.python.org/downloads/
  echo         and check "Add python.exe to PATH" during setup.
  echo.
  pause
  exit /b 1
)

set "PORT=8765"
if exist "config.json" (
  %PY% -c "import json;print(int(json.load(open('config.json',encoding='utf-8-sig')).get('port') or 8765))" > "%TEMP%\automail_port.txt" 2>nul
  if exist "%TEMP%\automail_port.txt" set /p PORT=<"%TEMP%\automail_port.txt"
  del /q "%TEMP%\automail_port.txt" >nul 2>&1
)
set "URL=http://127.0.0.1:!PORT!"

%PY% -c "import socket,sys;s=socket.socket();s.settimeout(1);sys.exit(0 if s.connect_ex(('127.0.0.1',!PORT!))==0 else 1)" >nul 2>&1
if !errorlevel!==0 (
  echo [i] Panel already running, opening browser...
  start "" "!URL!"
  exit /b 0
)

echo ============================================
echo   Automail Desk
echo   Panel : !URL!
echo   Log   : %~dp0logs\server.log
echo   Keep this window open. Ctrl+C stops it.
echo ============================================
echo.

start "" "!URL!"
%PY% "%~dp0server.py"

echo.
echo [server exited] press any key to close
pause >nul
