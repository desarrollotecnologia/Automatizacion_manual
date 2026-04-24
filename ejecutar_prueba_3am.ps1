# PRUEBA PROGRAMADA 3 AM — 1 cliente, 4 destinatarios fijos.
# NO activa EJECUCION_NOCTURNA (que enviaria a los 453 clientes).
# Envia el reporte del cliente de prueba SOLO a los correos indicados abajo.
#
# Cuando valides que llega bien y el formato esta OK:
#   - Desregistra esta tarea: Unregister-ScheduledTask -TaskName RendimientosBot_Prueba_03AM -Confirm:$false
#   - Activa la tarea real: .\Registrar_Tarea_3AM.ps1 (ya registrada -> RendimientosBot_03AM)
#   - Cuando corra la real, el modo nocturno IGNORA TEST_EMAIL_TO y envia a cada cliente

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

# Cliente de prueba: se envia UNA copia con este cliente y su data reciente.
# Cambia estos valores si quieres probar otro cliente.
$env:SOLO_CLIENTE_NOMBRE = "CALIXTO ARDILA JAIME"
$env:DB_PROPIETARIO_LIKE = "%CALIXTO ARDILA%"
# Fecha = "ayer" automatico (cuando corra a las 3 AM del 23, buscara datos del 22).
# Si NO hay datos del 22 para este cliente, cambia a fecha fija:
# $env:DB_FECHA_PLAN = "2026-04-21"
$env:DB_FECHA_PLAN_AUTO = "ayer"

# Destinatarios de la prueba (todos los correos van al TEST_EMAIL_TO).
$env:TEST_EMAIL_TO = "coordinacion.linea@colbeef.com,planillaje@frigorificoriofrio.com,desarrollo.tecnologia@colbeef.com,analista.tic@colbeef.com"
$env:TEST_EMAIL_SUBJECT = "Prueba Sistema Rendimientos Beef — $($env:SOLO_CLIENTE_NOMBRE)"

# Envio real (pero SOLO al TEST_EMAIL_TO, no al cliente).
$env:DRY_RUN = "0"

# IMPORTANTE: NO activar EJECUCION_NOCTURNA. Solo este cliente.
Remove-Item Env:EJECUCION_NOCTURNA -ErrorAction SilentlyContinue

$py = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
    Write-Error "No existe el entorno virtual: $py"
}

# Log con fecha para revisar despues.
$logDir = Join-Path $ProjectRoot "logs_prueba"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$logFile = Join-Path $logDir ("prueba_3am_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".log")

try {
    & $py (Join-Path $ProjectRoot "bot_reportes.py") *>&1 | Tee-Object -FilePath $logFile
    $exit = $LASTEXITCODE
    if ($exit -ne 0) {
        Write-Host "ERROR: bot termino con codigo $exit. Revisar log: $logFile"
    } else {
        Write-Host "OK: prueba completada. Log: $logFile"
    }
    exit $exit
} catch {
    Write-Host "EXCEPCION: $_"
    $_ | Out-File -FilePath $logFile -Append
    exit 1
}
