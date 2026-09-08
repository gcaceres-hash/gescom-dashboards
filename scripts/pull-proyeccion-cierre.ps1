<#
Arma el dashboard de Proyeccion de Cierre: a partir de los datos ya
calculados de la Pizarra de Rentabilidad (docs/rentabilidad/data.json),
proyecta como va a terminar el mes en curso segun el ritmo de venta de los
dias habiles (lunes a viernes) transcurridos hasta ahora.

No consulta la API ni ningun archivo de Gescom -- es un calculo derivado
100% de lo que ya publico Rentabilidad, asi que conviene correrlo siempre
despues de actualizar Rentabilidad (el watcher ya lo encadena).

Formula: proyectado = actual / diasHabilesTranscurridos * diasHabilesTotal
(dias habiles = lunes a viernes del mes calendario, sin descontar
feriados). Para un mes ya cerrado, diasHabilesTranscurridos = diasHabilesTotal,
asi que proyectado = actual (no hay nada que proyectar).

Uso:
  powershell -File pull-proyeccion-cierre.ps1
#>
param(
    [string]$RentabilidadPath = (Join-Path $PSScriptRoot "../docs/rentabilidad/data.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/proyeccion"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "proyeccion-cierre-template.html")
)
$ErrorActionPreference = "Stop"
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }

function Contar-DiasHabiles($desde, $hasta) {
    $count = 0
    $d = $desde
    while ($d -le $hasta) {
        if ($d.DayOfWeek -ne [System.DayOfWeek]::Saturday -and $d.DayOfWeek -ne [System.DayOfWeek]::Sunday) { $count++ }
        $d = $d.AddDays(1)
    }
    return $count
}
function Proyectar([double]$valorActual, [int]$transcurridos, [int]$total) {
    if ($transcurridos -le 0) { return $valorActual }
    return $valorActual / $transcurridos * $total
}

if (-not (Test-Path $RentabilidadPath)) { throw "No existe $RentabilidadPath -- corre primero pull-venta-rentabilidad-csv.ps1" }
Write-Log "Leyendo $RentabilidadPath ..."
$rentaData = (Get-Content $RentabilidadPath -Raw) | ConvertFrom-Json

$mesesOut = [ordered]@{}
foreach ($mesKey in $rentaData.meses.PSObject.Properties.Name) {
    $mes = $rentaData.meses.$mesKey
    $periodoDesde = ([datetime]$mes.periodoDesde).Date
    # Get-Date -Hour 0 -Minute 0 -Second 0 NO limpia los milisegundos (queda
    # un resto de "ahora"), lo que rompia la comparacion exacta de medianoche
    # contra fechas parseadas del JSON -- se fuerza con .Date en cambio.
    $inicioMes = (Get-Date -Year $periodoDesde.Year -Month $periodoDesde.Month -Day 1).Date
    $finMes = $inicioMes.AddMonths(1).AddDays(-1)
    $periodoHasta = ([datetime]$mes.periodoHasta).Date

    $diasHabilesTotal = Contar-DiasHabiles $inicioMes $finMes
    $diasHabilesTranscurridos = Contar-DiasHabiles $inicioMes $periodoHasta
    $diasHabilesRestantes = $diasHabilesTotal - $diasHabilesTranscurridos

    $t = $mes.totales
    $rentabilidadActual = $t.ventaNeta - $t.cmv
    $actual = [pscustomobject]@{
        ventaNeta = $t.ventaNeta; ventaConImp = $t.ventaConImp; cmv = $t.cmv; descuentos = $t.descuentos
        rentabilidad = [math]::Round($rentabilidadActual,2)
        margen = if ($t.ventaNeta -ne 0) { [math]::Round($rentabilidadActual / $t.ventaNeta,4) } else { 0 }
    }
    $ventaNetaProy = Proyectar $t.ventaNeta $diasHabilesTranscurridos $diasHabilesTotal
    $ventaConImpProy = Proyectar $t.ventaConImp $diasHabilesTranscurridos $diasHabilesTotal
    $cmvProy = Proyectar $t.cmv $diasHabilesTranscurridos $diasHabilesTotal
    $descuentosProy = Proyectar $t.descuentos $diasHabilesTranscurridos $diasHabilesTotal
    $rentabilidadProy = $ventaNetaProy - $cmvProy
    $proyectado = [pscustomobject]@{
        ventaNeta = [math]::Round($ventaNetaProy,2); ventaConImp = [math]::Round($ventaConImpProy,2)
        cmv = [math]::Round($cmvProy,2); descuentos = [math]::Round($descuentosProy,2)
        rentabilidad = [math]::Round($rentabilidadProy,2)
        margen = if ($ventaNetaProy -ne 0) { [math]::Round($rentabilidadProy / $ventaNetaProy,4) } else { 0 }
    }

    $proveedoresOut = foreach ($p in $mes.proveedores) {
        $rentabProvActual = $p.ventaNeta - $p.cmv
        $ventaProy = Proyectar $p.ventaNeta $diasHabilesTranscurridos $diasHabilesTotal
        $cmvProvProy = Proyectar $p.cmv $diasHabilesTranscurridos $diasHabilesTotal
        [pscustomobject]@{
            codigo = $p.codigo
            nombre = $p.nombre
            ventaNetaActual = [math]::Round($p.ventaNeta,2)
            cmvActual = [math]::Round($p.cmv,2)
            rentabilidadActual = [math]::Round($rentabProvActual,2)
            ventaNetaProyectada = [math]::Round($ventaProy,2)
            cmvProyectado = [math]::Round($cmvProvProy,2)
            rentabilidadProyectada = [math]::Round($ventaProy - $cmvProvProy,2)
        }
    }

    $mesesOut[$mesKey] = [pscustomobject]@{
        periodoDesde = $mes.periodoDesde
        periodoHasta = $mes.periodoHasta
        cerrado = $mes.cerrado
        diasHabilesTotal = $diasHabilesTotal
        diasHabilesTranscurridos = $diasHabilesTranscurridos
        diasHabilesRestantes = $diasHabilesRestantes
        actual = $actual
        proyectado = $proyectado
        proveedores = @($proveedoresOut | Sort-Object ventaNetaActual -Descending)
    }
    Write-Log "Mes $mesKey -> dias habiles $diasHabilesTranscurridos/$diasHabilesTotal | Venta neta actual: $($actual.ventaNeta) | proyectada: $($proyectado.ventaNeta)"
}

$out = [pscustomobject]@{
    generatedAt = (Get-Date).ToString("o")
    mesActual = $rentaData.mesActual
    meses = $mesesOut
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
$jsonText = $out | ConvertTo-Json -Depth 10 -Compress
[System.IO.File]::WriteAllText($OutPath, $jsonText, (New-Object System.Text.UTF8Encoding $false))
Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force
Write-Log "Guardado: $OutPath"
