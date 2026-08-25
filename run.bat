@echo off
cd /d "%~dp0"
echo ==============================================
echo   Enterprise Analysis - Local Dashboard
echo   URL: http://localhost:8501
echo   First run installs dependencies automatically.
echo   Close this window to stop the server.
echo ==============================================
where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found. Install Python 3.8+ and check "Add Python to PATH".
    pause
    exit /b 1
)
python -c "import streamlit, plotly, pandas, openpyxl, numpy" >nul 2>nul
if errorlevel 1 (
    echo First run: installing dependencies via pip ...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Dependency install failed. Try manually:
        echo   python -m pip install -r requirements.txt
        pause
        exit /b 1
    )
)
python -m streamlit run app.py
pause
