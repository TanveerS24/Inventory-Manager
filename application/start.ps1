# InventoryHouse Pro - PowerShell Startup Script

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   InventoryHouse Pro - Startup Script" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Check Python
Write-Host "Checking Python installation..." -ForegroundColor Yellow
$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    Write-Host "[ERROR] Python is not installed or not in PATH" -ForegroundColor Red
    Write-Host "Please install Python 3.9+ from https://python.org" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "Python found: $(python --version)" -ForegroundColor Green

# Check Node.js
Write-Host ""
Write-Host "Checking Node.js installation..." -ForegroundColor Yellow
$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
    Write-Host "[ERROR] Node.js is not installed or not in PATH" -ForegroundColor Red
    Write-Host "Please install Node.js from https://nodejs.org" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "Node.js found: $(node --version)" -ForegroundColor Green

# Setup Backend
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   Setting up Backend" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

Set-Location backend

# Create virtual environment
if (-not (Test-Path "venv")) {
    Write-Host "Creating Python virtual environment..." -ForegroundColor Yellow
    python -m venv venv
}

# Activate and install
Write-Host "Installing Python dependencies..." -ForegroundColor Yellow
& .\venv\Scripts\Activate.ps1
pip install -r requirements.txt

# Start backend
Write-Host "Starting FastAPI backend on http://127.0.0.1:8000..." -ForegroundColor Green
$backendJob = Start-Job -ScriptBlock {
    Set-Location $using:PWD
    & .\venv\Scripts\python.exe -m uvicorn main:app --host 127.0.0.1 --port 8000
}

Set-Location ..

# Wait for backend to start
Write-Host "Waiting for backend to start..." -ForegroundColor Yellow
Start-Sleep -Seconds 3

# Setup Frontend
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   Setting up Frontend" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

Set-Location frontend

# Install dependencies
if (-not (Test-Path "node_modules")) {
    Write-Host "Installing Node.js dependencies..." -ForegroundColor Yellow
    npm install
}

# Start Electron
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   Starting Application" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Backend running at: http://127.0.0.1:8000" -ForegroundColor Green
Write-Host "Starting Electron frontend..." -ForegroundColor Green
Write-Host ""

npm start

# Cleanup
Write-Host ""
Write-Host "Stopping backend server..." -ForegroundColor Yellow
Stop-Job $backendJob
Remove-Job $backendJob

Write-Host ""
Write-Host "Thank you for using InventoryHouse Pro!" -ForegroundColor Green
Read-Host "Press Enter to exit"
