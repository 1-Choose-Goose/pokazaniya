$ErrorActionPreference = "Stop"

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
    $python = "python"
}

& $python -m PyInstaller --noconfirm --clean (Join-Path $PSScriptRoot "Pokazaniya.spec")
if ($LASTEXITCODE -ne 0) {
    throw "Application build failed with exit code $LASTEXITCODE"
}

& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name PokazaniyaUpdater `
    --icon (Join-Path $PSScriptRoot "app_icon.ico") `
    --distpath (Join-Path $PSScriptRoot "build\updater-dist") `
    --workpath (Join-Path $PSScriptRoot "build\updater-build") `
    --specpath (Join-Path $PSScriptRoot "build") `
    (Join-Path $PSScriptRoot "updater.py")
if ($LASTEXITCODE -ne 0) {
    throw "Updater build failed with exit code $LASTEXITCODE"
}

$distribution = Join-Path $PSScriptRoot "dist\Pokazaniya"
$updater = Join-Path $PSScriptRoot "build\updater-dist\PokazaniyaUpdater.exe"
if (-not (Test-Path -LiteralPath $updater)) {
    throw "Build completed without PokazaniyaUpdater.exe"
}
Copy-Item -LiteralPath $updater -Destination (Join-Path $distribution "_internal\PokazaniyaUpdater.exe") -Force

Write-Host "Ready: $distribution"
