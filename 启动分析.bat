@echo off
setlocal
cd /d "%~dp0"
set "PY=%~dp0python\python.exe"

if not exist "%PY%" (
    echo [ERROR] Python runtime not found at: python\python.exe
    echo         Keep the whole extracted folder intact. Do not rename or move the python folder.
    echo.
    pause
    exit /b 1
)

echo ==============================================
echo   Enterprise Operation Analysis - Offline
echo   URL: http://localhost:8501
echo   Close this window to stop the program.
echo ==============================================
echo.

"%PY%" -m streamlit run app.py

echo.
echo Program stopped. You may close this window.
pause