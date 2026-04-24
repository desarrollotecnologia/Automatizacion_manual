# Registra una tarea en Windows para ejecutar el bot todos los días a las 03:00.
# Ejecutar desde la carpeta donde están ejecutar_noche.ps1, .env y RendimientosBot.exe (portable) o el .venv.
# Úsalo solo cuando tengas SMTP configurado y ejecutar_noche.ps1 con envío real ($EjecutarConEnvioReal = $true).
# En fase de pruebas (sin Gmail) NO registres esta tarea o seguirá en modo prueba según tu .env.
#
# Ejecutar PowerShell COMO ADMINISTRADOR si Windows lo pide.
# Para cambiar la hora, edita -At "03:00".

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptNoche = Join-Path $ProjectRoot "ejecutar_noche.ps1"

if (-not (Test-Path $scriptNoche)) {
    Write-Error "No se encuentra: $scriptNoche"
}

$taskName = "RendimientosBot_03AM"
$arg = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptNoche`""

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $arg
$trigger = New-ScheduledTaskTrigger -Daily -At "03:00"
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force -Description "Automatizacion rendimientos: consulta BD, genera Excel por cliente y envia correos."

Write-Host "Tarea registrada: $taskName (diaria 03:00). Revisa Programador de tareas de Windows."
