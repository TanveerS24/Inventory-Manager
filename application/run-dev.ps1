# InventoryHouse Pro - Development Runner
# This script starts both backend and frontend

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   InventoryHouse Pro - Development Mode" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$ErrorActionPreference = "Stop"

# Get the directory where this script is located
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

# Check Python
Write-Host "[1/6] Checking Python..." -ForegroundColor Yellow
$pyVersion = python --version 2>&1
Write-Host "      $pyVersion" -ForegroundColor Green

# Check Node.js
Write-Host "[2/6] Checking Node.js..." -ForegroundColor Yellow
$nodeVersion = node --version 2>&1
Write-Host "      Node.js $nodeVersion" -ForegroundColor Green

# Setup Backend
Write-Host ""
Write-Host "[3/6] Setting up Backend..." -ForegroundColor Yellow
Set-Location "$ScriptDir\backend"

# Create virtual environment if it doesn't exist
if (-not (Test-Path "venv")) {
    Write-Host "      Creating virtual environment..." -ForegroundColor Gray
    python -m venv venv
}

# Activate and install
Write-Host "      Installing Python dependencies..." -ForegroundColor Gray
& .\venv\Scripts\Activate.ps1
pip install -q -r requirements.txt

Write-Host "[4/6] Starting Backend Server..." -ForegroundColor Yellow
Write-Host "      API will be available at: http://127.0.0.1:8000" -ForegroundColor Cyan

# Start backend in a new window
$BackendProcess = Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$ScriptDir\backend'; & .\venv\Scripts\Activate.ps1; uvicorn main:app --host 127.0.0.1 --port 8000 --reload" -PassThru

# Wait for backend to start
Write-Host "      Waiting for backend to start..." -ForegroundColor Gray
Start-Sleep -Seconds 4

# Test if backend is running
try {
    $response = Invoke-RestMethod -Uri "http://127.0.0.1:8000/health" -Method GET -TimeoutSec 5
    Write-Host "      Backend is running! ($($response.app) v$($response.version))" -ForegroundColor Green
} catch {
    Write-Host "      Warning: Backend may not be ready yet" -ForegroundColor Yellow
}

# Setup Frontend
Write-Host ""
Write-Host "[5/6] Setting up Frontend..." -ForegroundColor Yellow
Set-Location "$ScriptDir\frontend"

# Install dependencies if needed
if (-not (Test-Path "node_modules")) {
    Write-Host "      Installing Node.js dependencies..." -ForegroundColor Gray
    npm install
} else {
    Write-Host "      Node modules already installed" -ForegroundColor Gray
}

# Start Frontend
Write-Host "[6/6] Starting Electron Frontend..." -ForegroundColor Yellow
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "   Application Starting!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Backend:  http://127.0.0.1:8000" -ForegroundColor Cyan
Write-Host "Frontend: Electron Desktop App" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press Ctrl+C in the backend window to stop the server" -ForegroundColor Yellow
Write-Host "Close the Electron app to stop the frontend" -ForegroundColor Yellow
Write-Host ""

# Start Electron
npm start

# Cleanup
Write-Host ""
Write-Host "Stopping backend server..." -ForegroundColor Yellow
if ($BackendProcess -and -not $BackendProcess.HasExited) {
    Stop-Process -Id $BackendProcess.Id -Force -ErrorAction SilentlyContinue
}

Write-Host "Development session ended." -ForegroundColor Green
