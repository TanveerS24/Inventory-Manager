@echo off
echo ============================================
echo    InventoryHouse Pro - Startup Script
echo ============================================
echo.

REM Check Python
echo Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH
    echo Please install Python 3.9+ from https://python.org
    pause
    exit /b 1
)

REM Check Node.js
echo Checking Node.js installation...
node --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Node.js is not installed or not in PATH
    echo Please install Node.js from https://nodejs.org
    pause
    exit /b 1
)

echo.
echo ============================================
echo    Setting up Backend
echo ============================================
echo.

cd backend

REM Create virtual environment if not exists
if not exist "venv" (
    echo Creating Python virtual environment...
    python -m venv venv
)

REM Activate virtual environment
call venv\Scripts\activate.bat

REM Install dependencies
echo Installing Python dependencies...
pip install -r requirements.txt

REM Start backend in background
echo Starting FastAPI backend on http://127.0.0.1:8000...
start "InventoryHouse Backend" cmd /c "uvicorn main:app --host 127.0.0.1 --port 8000 --reload"

cd ..

echo.
echo ============================================
echo    Setting up Frontend
echo ============================================
echo.

cd frontend

REM Install dependencies if node_modules doesn't exist
if not exist "node_modules" (
    echo Installing Node.js dependencies...
    npm install
)

echo.
echo ============================================
echo    Starting Application
echo ============================================
echo.
echo Backend is running at: http://127.0.0.1:8000
echo Frontend will start shortly...
echo.
echo Press Ctrl+C in the backend window to stop the server
echo.

REM Start Electron app
npm start

REM Cleanup on exit
echo.
echo Stopping backend server...
taskkill /FI "WINDOWTITLE eq InventoryHouse Backend" /F >nul 2>&1

echo.
echo Thank you for using InventoryHouse Pro!
pause
