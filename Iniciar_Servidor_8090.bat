@echo off
REM Inicia este equipo como servidor web en el puerto 8090.
REM Comparte la URL que muestra la consola: http://IP_DE_ESTE_EQUIPO:8090/

setlocal
cd /d "%~dp0"

set "PORT=8090"
set "PY=%~dp0.venv\Scripts\python.exe"

if exist "%PY%" goto run
set "PY=python"

:run
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /c:"IPv4"') do (
  set "LOCAL_IP=%%A"
  goto gotip
)

:gotip
set "LOCAL_IP=%LOCAL_IP: =%"

echo.
echo Servidor iniciado en este equipo.
echo URL local:      http://localhost:%PORT%/
if defined LOCAL_IP echo URL para la red: http://%LOCAL_IP%:%PORT%/
echo.
echo Para detenerlo, cierra esta ventana o presiona Ctrl+C.
echo.

start "" "http://localhost:%PORT%/"
"%PY%" -m http.server %PORT% --bind 0.0.0.0

pause
