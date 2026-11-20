# setup.ps1 - one-time setup on a Windows laptop (run from this folder).
# Needs Python 3.11+ and Node.js LTS installed first.
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
if (-not (Test-Path .venv)) { python -m venv .venv }
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m playwright install chromium
npm install
Write-Host "Setup done. Benchmark run:  .\.venv\Scripts\python.exe pipeline.py jobs\crown-villa --project `"Sample Villa`""
