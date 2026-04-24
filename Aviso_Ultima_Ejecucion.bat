@echo off
REM Muestra una ventana con el resumen de la ultima corrida exitosa del bot (archivo ultima_corrida.txt).
REM Ese archivo se crea al terminar bien el bot (.exe portable o bot_reportes.py / tarea 03:00).
setlocal
cd /d "%~dp0"
set "ULTIMA=%~dp0ultima_corrida.txt"
if exist "%ULTIMA%" (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.Windows.Forms; $p = [Environment]::GetEnvironmentVariable('ULTIMA','Process'); $t = Get-Content -LiteralPath $p -Raw -Encoding UTF8; if ([string]::IsNullOrWhiteSpace($t)) { $t = '(archivo vacio)' }; [void][System.Windows.Forms.MessageBox]::Show($t, 'Ultima ejecucion - Bot rendimientos',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.Windows.Forms; [void][System.Windows.Forms.MessageBox]::Show('No hay registro aun. Cuando el bot termine sin error global, se creara ultima_corrida.txt en esta carpeta.','Ultima ejecucion - Bot rendimientos',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)"
)
