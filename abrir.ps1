$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$back = Join-Path $root "backend"
$front = Join-Path $root "frontend"

Write-Host "=== Gerador de TDTs - ADMS ===" -ForegroundColor Cyan

Write-Host "Dependencias do backend..."
python -m pip install -q -r (Join-Path $back "requirements.txt")

# Excel e' obrigatorio: o ADMS so aceita arquivos salvos pelo MS Excel.
$excelReg = $null
try { $excelReg = Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\Excel.Application\CurVer" -ErrorAction Stop } catch {}
if (-not $excelReg) {
  Write-Host "AVISO: MS Excel nao detectado nesta maquina." -ForegroundColor Yellow
  Write-Host "       A TDT gerada pode ser RECUSADA pelo ADMS ('Invalid TDI file format')." -ForegroundColor Yellow
  Write-Host "       Instale o MS Excel para a conversao para o formato oficial." -ForegroundColor Yellow
}

if (-not (Test-Path (Join-Path $back "data\catalog.json"))) {
  Write-Host "Gerando catalogo..."
  Push-Location $back; python build_catalog.py; Pop-Location
}

Write-Host "Iniciando backend (8077)..."
Start-Process powershell -ArgumentList @(
  "-NoExit","-NoProfile","-Command",
  "Set-Location '$back'; python -m uvicorn app:app --host 127.0.0.1 --port 8077"
)

if (-not (Test-Path (Join-Path $front "node_modules"))) {
  Write-Host "Instalando dependencias do frontend (pode demorar)..."
  Push-Location $front; npm install; Pop-Location
}

Write-Host "Iniciando frontend (5180)..."
Start-Process powershell -ArgumentList @(
  "-NoExit","-NoProfile","-Command",
  "Set-Location '$front'; npm run dev"
)

Start-Sleep -Seconds 8

# Abre em MODO APP (janela sem barra do navegador). Tenta Chrome, depois Edge.
$url = "http://localhost:5180"
$chrome = @(
  "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
  "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
  "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
$edge = @(
  "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
  "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($chrome) {
  Start-Process $chrome -ArgumentList "--app=$url","--window-size=1280,860"
  Write-Host "Aberto como app (Chrome)." -ForegroundColor Green
} elseif ($edge) {
  Start-Process $edge -ArgumentList "--app=$url","--window-size=1280,860"
  Write-Host "Aberto como app (Edge)." -ForegroundColor Green
} else {
  Start-Process $url
  Write-Host "Aberto no navegador padrão." -ForegroundColor Green
}
Write-Host "Dica: na janela do app, menu > Instalar para ter um atalho fixo." -ForegroundColor DarkGray
