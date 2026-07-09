@echo off
REM Abre el sistema SIRT y ejecuta el bot de rendimientos.
REM Debe ejecutarse desde la carpeta del proyecto/portable en el servidor.
setlocal
cd /d "%~dp0"

start "" "http://192.168.20.205:8090/"

if exist "%~dp0RendimientosBot.exe" goto runexe

set "PY=%~dp0.venv\Scripts\python.exe"
if exist "%PY%" goto runpy

echo No se encontro RendimientosBot.exe ni "%PY%" en esta carpeta.
pause
exit /b 1

:runexe
"%~dp0RendimientosBot.exe"
set "EC=%ERRORLEVEL%"
goto avisar

:runpy
"%PY%" "%~dp0bot_reportes.py"
set "EC=%ERRORLEVEL%"
goto avisar

:avisar
if "%EC%"=="0" (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.Windows.Forms; [void][System.Windows.Forms.MessageBox]::Show('Ejecucion terminada correctamente. Revise el correo y la carpeta de salida configurada (REPORT_OUTPUT_DIR).','Rendimientos - Bot manual',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.Windows.Forms; [void][System.Windows.Forms.MessageBox]::Show('La ejecucion termino con errores (codigo %EC%). Revise bot_reportes.log en esta carpeta.','Rendimientos - Bot manual',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)"
)
pause
exit /b %EC%
