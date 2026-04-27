@echo off
chcp 65001 >nul
echo ============================================
echo    InventoryHouse Pro - Frontend
echo ============================================
echo.

cd /d "%~dp0\frontend"

if not exist "node_modules" (
    echo [1/2] Installing Node.js dependencies...
    npm install
    echo.
) else (
    echo [1/2] Node modules already installed
echo.
)

echo [2/2] Starting Electron...
echo.
echo Make sure the backend is running first!
echo Backend: http://127.0.0.1:8000
echo.

npm start
