# Copia el codigo del proyecto a la carpeta de produccion.
#
# Uso (PowerShell, desde la carpeta del proyecto):
#   .\scripts\publicar_a_produccion.ps1 -Destino 'C:\MACROS_COMPRAS\MACRO_VENTAS\Generacion Nuevo Proceso'
#
# Copia SOLO codigo y configuracion. No toca .env, .venv, logs ni archivos de datos.
# Al final ejecuta "uv sync" en el destino para instalar dependencias nuevas.

param(
    [Parameter(Mandatory = $true)][string]$Destino
)

$ErrorActionPreference = 'Stop'
$Origen = Split-Path -Parent $PSScriptRoot

if (-not (Test-Path $Destino)) { throw "No existe la carpeta destino: $Destino" }

# Carpetas completas (agrega y reemplaza; no borra lo que haya de mas en destino)
foreach ($carpeta in @('src', 'tasks', 'core', 'config', 'scripts', 'docs\adr', 'tests')) {
    $o = Join-Path $Origen $carpeta
    $d = Join-Path $Destino $carpeta
    if (Test-Path $o) {
        robocopy $o $d /E /XD __pycache__ .mypy_cache .pytest_cache /XF *.pyc /NFL /NDL /NJH /NJS /NP | Out-Null
        if ($LASTEXITCODE -ge 8) { throw "robocopy fallo copiando $carpeta (codigo $LASTEXITCODE)" }
        Write-Host "OK  $carpeta"
    }
}

# Archivos sueltos de la raiz
foreach ($archivo in @('run.py', 'orchestrator.py', 'pyproject.toml', 'uv.lock', 'requirements.txt',
                       '.importlinter', '.env.example', 'README.md')) {
    $o = Join-Path $Origen $archivo
    if (Test-Path $o) {
        Copy-Item $o (Join-Path $Destino $archivo) -Force
        Write-Host "OK  $archivo"
    }
}

# Las descargas por Selenium se retiraron (ADR 0008): si quedaron en destino, se borran.
foreach ($viejo in @('tasks\descarga_valorizados.py', 'tasks\descarga_inventario_general.py',
                     'tasks\descarga_informe_ventas.py', 'core\browser.py', 'core\erp_navigation.py',
                     'config\erp_selectors.py')) {
    $p = Join-Path $Destino $viejo
    if (Test-Path $p) { Remove-Item $p -Force; Write-Host "BORRADO  $viejo" }
}

Push-Location $Destino
try {
    uv sync
} finally {
    Pop-Location
}
Write-Host ""
Write-Host "Listo. Revisa que el .env de destino tenga DB_EXPORT_DIR."
