# Ejecución nocturna / por lotes (misma carpeta que el bot: .env, RendimientosBot.exe o .venv).
# En el servidor: copie aquí la carpeta portable (build_portable.ps1 -> dist\RendimientosBot), su .env,
# registre la tarea con Registrar_Tarea_3AM.ps1 y use Ejecutar_Rendimientos_Manual.bat solo para corridas manuales
# (ese .bat NO activa EJECUCION_NOCTURNA; este script sí cuando $EjecutarConEnvioReal = $true).
#
# FASE PRUEBAS (sin Gmail ni SMTP aún): deja $EjecutarConEnvioReal = $false
#   -> usa solo tu .env (p. ej. DRY_RUN=1: genera Excel, NO envía correos).
# FASE PRODUCCIÓN (cuando ya tengas SMTP/Gmail u otro servidor): pon $EjecutarConEnvioReal = $true
#   -> activa EJECUCION_NOCTURNA (todos los clientes, envío real). Requiere SMTP_* correctos en .env.
#
# Programado a las 03:00: Registrar_Tarea_3AM.ps1 (solo cuando vayas a producción con envío).

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

$EjecutarConEnvioReal = $true

if ($EjecutarConEnvioReal) {
    $env:EJECUCION_NOCTURNA = "1"
} else {
    Remove-Item Env:EJECUCION_NOCTURNA -ErrorAction SilentlyContinue
}
# Opcional: fuerza la fecha de faena en la consulta (si no, el bot usa "ayer" por defecto en modo noche)
# $env:DB_FECHA_PLAN_AUTO = "ayer"   # o "hoy"

$exe = Join-Path $ProjectRoot "RendimientosBot.exe"
if (Test-Path $exe) {
    & $exe
    exit $LASTEXITCODE
}

$py = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
    Write-Error "No existe RendimientosBot.exe ni el entorno virtual: $py"
}

& $py (Join-Path $ProjectRoot "bot_reportes.py")
exit $LASTEXITCODE
