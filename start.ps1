# Sobe backend (FastAPI) + frontend (Vite) em janelas separadas.
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot

Write-Host "Iniciando backend em http://127.0.0.1:8077 ..." -ForegroundColor Cyan
Start-Process powershell -ArgumentList @(
  "-NoExit", "-Command",
  "cd '$root\backend'; if (-not (Test-Path data\catalog.json)) { python build_catalog.py }; uvicorn app:app --host 127.0.0.1 --port 8077"
)

Start-Sleep -Seconds 2

Write-Host "Iniciando frontend em http://localhost:5180 ..." -ForegroundColor Cyan
Start-Process powershell -ArgumentList @(
  "-NoExit", "-Command",
  "cd '$root\frontend'; if (-not (Test-Path node_modules)) { npm install }; npm run dev"
)

Write-Host ""
Write-Host "Pronto! Abra http://localhost:5180 no navegador." -ForegroundColor Green
