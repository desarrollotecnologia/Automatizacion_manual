@echo off
REM Reintenta el envio SOLO para los clientes que fallaron.
REM Uso: doble clic en el servidor (misma carpeta del bot).
REM Requiere: .env configurado en produccion (envio real, SIRT, correos_opcionales).

setlocal
cd /d "%~dp0"

REM Nombres separados por | (pipe). El bot los busca por coincidencia parcial normalizada.
set "SOLO_CLIENTE_NOMBRE=CRUZ LEONIDAS|INVERSIONES ZULUAGA RUEDA|JAIMES BERMUDEZ JOSE MARIA"

REM Fecha del plan: por defecto usa "ayer" (misma logica de la tarea 03:00).
REM Si vas a reintentar despues de mas de 1 dia, comenta esta linea y
REM descomenta DB_FECHA_PLAN con la fecha exacta de la faena original (YYYY-MM-DD).
set "DB_FECHA_PLAN_AUTO=ayer"
REM set "DB_FECHA_PLAN=2026-07-06"

REM Envio real (no dry run). Comentar si solo quieres probar sin enviar.
set "DRY_RUN=0"

REM No activar EJECUCION_NOCTURNA (esa modalidad limpia SOLO_CLIENTE_NOMBRE).
set "EJECUCION_NOCTURNA="

if exist "%~dp0RendimientosBot.exe" (
    "%~dp0RendimientosBot.exe"
    set "EC=%ERRORLEVEL%"
    goto avisar
)

set "PY=%~dp0.venv\Scripts\python.exe"
if not exist "%PY%" (
    echo No se encontro RendimientosBot.exe ni "%PY%".
    pause
    exit /b 1
)

"%PY%" "%~dp0bot_reportes.py"
set "EC=%ERRORLEVEL%"

:avisar
if "%EC%"=="0" (
    echo.
    echo OK: reintento terminado. Revisa bot_reportes.log y el correo de resumen.
) else (
    echo.
    echo ERROR: reintento termino con codigo %EC%. Revisa bot_reportes.log.
)
pause
exit /b %EC%
