@echo off
chcp 65001 >nul
echo ============================================
echo    InventoryHouse Pro - Backend Server
echo ============================================
echo.

cd /d "%~dp0\backend"

if not exist "venv" (
    echo [1/3] Creating virtual environment...
    python -m venv venv
    echo.
)

echo [2/3] Activating virtual environment...
call venv\Scripts\activate.bat

echo [3/3] Installing dependencies...
pip install -q -r requirements.txt

echo.
echo ============================================
echo    Starting FastAPI Server
echo ============================================
echo.
echo API URL: http://127.0.0.1:8000
echo Docs:    http://127.0.0.1:8000/api/v1/docs
echo.
echo Press Ctrl+C to stop the server
echo.

uvicorn main:app --host 127.0.0.1 --port 8000 --reload
