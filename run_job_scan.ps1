$ErrorActionPreference = "Stop"

$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
$BundledPython = "C:\Users\lkmee\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if (Test-Path $BundledPython) {
    $Python = $BundledPython
} else {
    $Python = (Get-Command python -ErrorAction SilentlyContinue).Source
}

if (-not $Python) {
    Write-Host "Python was not found."
    Write-Host "Install Python 3.11+ from https://www.python.org/downloads/ and then run:"
    Write-Host "  pip install -r requirements.txt"
    exit 1
}

& $Python "$Here\job_radar.py"
