# Genera el reporte en PDF desde SIRT (según .env) usando la plantilla
# DEMO_TEMPLATE_ESTRUCTURA_ACEBEDO_v2.xlsx y abre el PDF en Google Chrome.
#
# Uso: .\generar_pdf_y_abrir_chrome.ps1
# Opcional (PowerShell): $env:DB_FECHA_PLAN = "2026-04-15"; .\generar_pdf_y_abrir_chrome.ps1

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

Remove-Item Env:DEMO_XLSX -ErrorAction SilentlyContinue
if (-not $env:OPEN_PDF_IN_CHROME) {
    $env:OPEN_PDF_IN_CHROME = "1"
}

$py = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
    $py = "python"
}

& $py (Join-Path $ProjectRoot "bot_reportes.py")
exit $LASTEXITCODE
