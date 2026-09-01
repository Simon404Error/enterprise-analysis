@echo off
setlocal
cd /d "%~dp0"
set "PY=%~dp0python\python.exe"

if not exist "%PY%" (
    echo [ERROR] Python runtime not found at python\python.exe
    echo         Keep the whole extracted folder intact.
    pause
    exit /b 1
)

"%PY%" "%~dp0make_report.py" %*
pause