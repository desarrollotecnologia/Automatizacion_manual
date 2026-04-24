@echo off
REM Manual del bot (no activa EJECUCION_NOCTURNA; respeta el .env). La tarea de las 03:00 usa ejecutar_noche.ps1.
REM Carpeta portable: RendimientosBot.exe. En servidor: ejecute este .bat en sesion del servidor (RDP/local), no solo abriendo desde red en otra PC.
REM Modo desarrollo: sin .exe usa .venv\Scripts\python.exe + bot_reportes.py. Requiere Windows y Excel si genera con .xlsm.
setlocal
cd /d "%~dp0"
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
exit /b %EC%
