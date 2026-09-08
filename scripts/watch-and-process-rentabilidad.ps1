<#
Busca en el Escritorio y en Descargas el export mas reciente de "Detallado de
ventas extendido" (el que Gisela genera a mano desde la web de Gescom -- el
navegador a veces lo guarda en Descargas en vez de en el Escritorio), lo
procesa con pull-venta-rentabilidad-csv.ps1 y sube el resultado a GitHub --
todo sin intervencion manual, mas alla del export en si (que Gescom no
expone por API).

Pensado para correr en una Tarea Programada de Windows cada cierto tiempo.
Si no hay un archivo mas nuevo que la ultima vez procesada, no hace nada.
#>
param(
    [string[]]$Carpetas = @("C:\Users\gcaceres\Desktop", "C:\Users\gcaceres\Downloads"),
    [string]$Patron = "ventas-Detallado de ventas extendido-*.csv",
    [string]$RepoDir = "C:\Users\gcaceres\gescom-dashboards",
    [string]$MarcadorPath = (Join-Path $PSScriptRoot "ultimo-csv-procesado.txt")
)
$ErrorActionPreference = "Stop"
function Write-Log($m) { Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $m" }

# --- elegir el candidato mas reciente (de cualquiera de las carpetas) que
# tenga el formato de detalle (no el pivot chico) ---
$candidatos = Get-ChildItem -Path $Carpetas -Filter $Patron -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending
$elegido = $null
foreach ($c in $candidatos) {
    $enc = [System.Text.Encoding]::GetEncoding(1252)
    $fs = [System.IO.File]::Open($c.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $sr = New-Object System.IO.StreamReader($fs, $enc)
        $primeraLinea = $sr.ReadLine()
    } finally {
        $fs.Dispose()
    }
    if ($primeraLinea -match "ImporteNetoItem" -and $primeraLinea -match "PesoKgReal") {
        $elegido = $c
        break
    }
}
if (-not $elegido) {
    Write-Log "No se encontro ningun archivo de detalle valido en $($Carpetas -join ', ') con el patron '$Patron'."
    exit 0
}

$marcaAnterior = if (Test-Path $MarcadorPath) { Get-Content $MarcadorPath -Raw } else { "" }
$marcaActual = "$($elegido.FullName)|$($elegido.LastWriteTimeUtc.Ticks)"
if ($marcaActual -eq $marcaAnterior.Trim()) {
    Write-Log "Sin cambios: '$($elegido.Name)' ya fue procesado (misma fecha de modificacion)."
    exit 0
}

Write-Log "Procesando '$($elegido.Name)' (modificado $($elegido.LastWriteTime))..."
& powershell -NoProfile -Command "& '$RepoDir\scripts\pull-venta-rentabilidad-csv.ps1' -CsvPath '$($elegido.FullName)' -DocsDir '$RepoDir\docs\rentabilidad' -TemplatePath '$RepoDir\scripts\venta-rentabilidad-template.html'"
if ($LASTEXITCODE -ne 0) { throw "pull-venta-rentabilidad-csv.ps1 fallo con codigo $LASTEXITCODE" }

Write-Log "Recalculando proyeccion de cierre..."
& powershell -NoProfile -File "$RepoDir\scripts\pull-proyeccion-cierre.ps1"
if ($LASTEXITCODE -ne 0) { throw "pull-proyeccion-cierre.ps1 fallo con codigo $LASTEXITCODE" }

Set-Content -Path $MarcadorPath -Value $marcaActual -NoNewline

Write-Log "Subiendo a GitHub..."
Push-Location $RepoDir
try {
    git add scripts/pull-venta-rentabilidad-csv.ps1 docs/rentabilidad/data.json docs/rentabilidad/index.html docs/proyeccion/data.json docs/proyeccion/index.html
    $cambios = @(git status --porcelain -- docs/rentabilidad docs/proyeccion scripts/pull-venta-rentabilidad-csv.ps1)
    if ($cambios.Count -eq 0) {
        Write-Log "No hay cambios para commitear (el data.json ya estaba igual)."
    } else {
        git commit -m "Actualiza Pizarra de Rentabilidad desde $($elegido.Name)"
        git fetch origin
        git merge origin/main -X ours -m "Merge automatico, favorece datos del CSV"
        git push origin main
        Write-Log "Listo, subido a GitHub."
    }
} finally {
    Pop-Location
}
