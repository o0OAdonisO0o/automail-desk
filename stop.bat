@echo off
setlocal
cd /d "%~dp0"

set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY (
  where py >nul 2>&1 && set "PY=py -3"
)
if not defined PY (
  echo [ERROR] Python 3.9+ was not found on PATH.
  pause
  exit /b 1
)

set "PORT=8765"
if exist "config.json" (
  %PY% -c "import json;print(int(json.load(open('config.json',encoding='utf-8-sig')).get('port') or 8765))" > "%TEMP%\automail_port.txt" 2>nul
  if exist "%TEMP%\automail_port.txt" set /p PORT=<"%TEMP%\automail_port.txt"
  del /q "%TEMP%\automail_port.txt" >nul 2>&1
)

echo Stopping Automail Desk on port %PORT% ...
%PY% -c "import os,signal,subprocess;out=subprocess.run(['netstat','-ano'],capture_output=True).stdout.decode('gbk','ignore');pids={l.split()[-1] for l in out.splitlines() if ':%PORT%' in l and 'LISTENING' in l};[os.kill(int(p),signal.SIGTERM) for p in pids];print('stopped' if pids else 'no listener', sorted(pids))"
timeout /t 2 >nul
echo Done.
timeout /t 2 >nul
