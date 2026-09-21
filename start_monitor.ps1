# =====================================================================
# PlacementMonitor PowerShell Launcher
# Activates the virtual environment and starts continuous monitoring
# =====================================================================

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location -Path $ScriptDir

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "               Starting PlacementMonitor                    " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Project Directory: $ScriptDir" -ForegroundColor DarkGray

# Verify virtual environment
$VenvPython = Join-Path $ScriptDir ".venv\Scripts\python.exe"
$VenvActivate = Join-Path $ScriptDir ".venv\Scripts\Activate.ps1"

if (-not (Test-Path $VenvPython)) {
    Write-Host "`n[ERROR] Virtual environment not found at: $VenvPython" -ForegroundColor Red
    Write-Host "Please create the virtual environment first using: python -m venv .venv" -ForegroundColor Yellow
    exit 1
}

# Activate virtual environment
if (Test-Path $VenvActivate) {
    Write-Host "[Launcher] Activating virtual environment..." -ForegroundColor DarkGray
    & $VenvActivate
}

# Run the monitoring engine
Write-Host "[Launcher] Starting PlacementMonitor..." -ForegroundColor Green
Write-Host "[Launcher] Press Ctrl+C at any time to stop monitoring cleanly.`n" -ForegroundColor DarkGray

& $VenvPython "src/main.py"
