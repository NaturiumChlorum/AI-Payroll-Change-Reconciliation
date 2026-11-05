param([int]$Port = 8000)
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
if (!(Test-Path '.venv312/Scripts/python.exe')) { throw 'Please follow README to install dependencies first.' }
if (!(Test-Path 'frontend/dist/index.html')) { throw 'Run npm.cmd run build in frontend first.' }
& '.venv312/Scripts/python.exe' -m uvicorn backend.main:app --host 127.0.0.1 --port $Port
