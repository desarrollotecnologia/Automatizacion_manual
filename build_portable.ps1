# Genera una carpeta portable: los usuarios solo copian esa carpeta + su .env (sin Python ni repo).
# Uso (en ESTE proyecto, una vez):  .\.venv\Scripts\pip install -r requirements.txt -r requirements-build.txt
#                                    .\build_portable.ps1
# Salida: dist\RendimientosBot\  (zip esa carpeta para repartir; junto al .exe van los .bat y Colbeef.png si existe).

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$py = Join-Path $Root ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
    Write-Error "No existe $py. Cree el venv e instale: python -m venv .venv ; .\.venv\Scripts\pip install -r requirements.txt -r requirements-build.txt"
}

Write-Host "Instalando PyInstaller si falta..."
& $py -m pip install -q -r (Join-Path $Root "requirements-build.txt")

$piArgs = @(
    "-m", "PyInstaller",
    "--noconfirm", "--clean",
    "--name", "RendimientosBot",
    "--onedir",
    "--console",
    (Join-Path $Root "bot_reportes.py"),
    "--collect-all", "pywin32",
    "--hidden-import", "win32timezone",
    "--hidden-import", "pythoncom",
    "--hidden-import", "pywintypes",
    "--hidden-import", "win32com.client",
    "--hidden-import", "pyodbc",
    "--hidden-import", "psycopg",
    "--collect-all", "psycopg"
)

Write-Host "PyInstaller: $($piArgs -join ' ')"
& $py @piArgs
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$outDir = Join-Path $Root "dist\RendimientosBot"
if (-not (Test-Path $outDir)) {
    Write-Error "No se creo la carpeta esperada: $outDir"
}

Copy-Item -Force (Join-Path $Root "Ejecutar_Rendimientos_Manual.bat") $outDir
Copy-Item -Force (Join-Path $Root "Aviso_Ultima_Ejecucion.bat") $outDir
Copy-Item -Force (Join-Path $Root "ejecutar_noche.ps1") $outDir
Copy-Item -Force (Join-Path $Root "Registrar_Tarea_3AM.ps1") $outDir
$logo = Join-Path $Root "Colbeef.png"
if (Test-Path $logo) {
    Copy-Item -Force $logo $outDir
}

Write-Host ""
Write-Host "Listo. Carpeta lista para servidor o reparto: $outDir"
Write-Host "Servidor: copie la carpeta, ponga .env, ejecute Registrar_Tarea_3AM.ps1 (admin) para las 03:00, y manual con Ejecutar_Rendimientos_Manual.bat."
Write-Host "Manual en el servidor (Escritorio remoto o sesion local); si abre el .bat desde red en su PC, corre en su PC, no en el servidor."
Write-Host 'Nota: Windows y Excel si usa plantilla .xlsm; red/VPN segun el .env.'
