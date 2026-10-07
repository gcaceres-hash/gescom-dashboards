<#
Mismo dashboard que Cobertura General, pero el "universo" de cada vendedor se
recorta a sus clientes "potenciales": los que concentran el 80% de lo que ese
vendedor vendio en los ultimos 3 meses (curva de Pareto por vendedor).

MIGRADO el 30/9/2026 de la API de Gescom a la base compartida de Lucas
(datos-gescom.panelempresas.workers.dev) -- ver pull-cobertura-general.ps1
para el detalle de por que y el cambio de criterio de fecha (fechaPedido en
vez de fechaComprobante, ya que la base no guarda la fecha de comprobante).

Reglas:
  - Se excluyen los vendedores 1176, 43, 16, 37 (no son vendedores reales).
    1176 y 43 ya vienen excluidos de la tabla `ventas` de origen.
  - Ranking (quien es "potencial"): venta neta por cliente en los ultimos 3
    meses (ventana movil hasta hoy), solo tipo VEN/DEB (las notas de credito
    NO se suman ni se restan, se ignoran igual que en la version anterior),
    ordenado de mayor a menor, se toman los clientes hasta llegar al 80%
    acumulado de la venta de ese vendedor.
  - "Compro" (para el indicador de cobertura) = venta VEN/DEB en el MES EN
    CURSO -- igual criterio que Cobertura General.
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/potenciales"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "potenciales-template.html"),
    [double]$CorteAcumulado = 0.8
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$EXCLUDED = @("43","1176","16","37")
$DIA_LETRA = @{ "L"="lunes"; "M"="martes"; "X"="miercoles"; "J"="jueves"; "V"="viernes"; "S"="sabado"; "D"="domingo" }

function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
function Invoke-PanelSql([string]$Sql) {
    $uri = "$($config.baseUrl)/consulta?sql=" + [uri]::EscapeDataString($Sql)
    $headers = @{ Authorization = "Bearer $($config.clave)" }
    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    if ($r.truncado) { Write-Log "AVISO: la consulta vino truncada (mas de 20000 filas) -- revisar y acotar." }
    return $r.filas
}

Write-Log "Descargando clientes con cartera asignada..."
$clientesRaw = Invoke-PanelSql "SELECT codigo, nombre, localidad, vendedor, rutas FROM clientes WHERE activo = 1 AND vendedor IS NOT NULL AND vendedor <> ''"
# Guarda (7/10/2026): ver pull-cobertura-general.ps1 -- si la base viene sin vendedores, no pisar el dato anterior.
if (@($clientesRaw).Count -lt 1000) { throw "La base trae solo $(@($clientesRaw).Count) clientes con vendedor asignado (lo normal son ~3.300): posible falla de la sincronizacion de catalogos. No se actualiza para no pisar el dato anterior; avisar a Lucas." }

$clienteInfo = @{}
$vendedoresConCartera = New-Object System.Collections.Generic.HashSet[string]
foreach ($c in $clientesRaw) {
    $vcod = [string]$c.vendedor
    if ($EXCLUDED -contains $vcod) { continue }
    $dias = @()
    try {
        $rutas = $c.rutas | ConvertFrom-Json
        if ($rutas -and $rutas.Count -gt 0) {
            $letra = [string]$rutas[0].dias
            $dias = @($letra.ToCharArray() | ForEach-Object { if ($DIA_LETRA.ContainsKey([string]$_)) { $DIA_LETRA[[string]$_] } })
        }
    } catch {}
    $clienteInfo[[string]$c.codigo] = [pscustomobject]@{
        codigo = [string]$c.codigo; nombre = $c.nombre; localidad = $c.localidad
        codigoVendedor = $vcod; dias = $dias
        ventaNeta3m = 0.0; compro = $false; proveedoresComprados = (New-Object System.Collections.Generic.HashSet[string])
    }
    [void]$vendedoresConCartera.Add($vcod)
}
Write-Log "Clientes con cartera asignada (vendedor real): $($clienteInfo.Count) | vendedores distintos: $($vendedoresConCartera.Count)"

$hoy = (Get-Date).Date
$rankingDesde = $hoy.AddMonths(-3)
$inicioMesActual = Get-Date -Year $hoy.Year -Month $hoy.Month -Day 1
$fechaDesde = $rankingDesde.ToString("yyyy-MM-dd")
$fechaHasta = $inicioMesActual.AddMonths(1).ToString("yyyy-MM-dd")
$fechaDesdeMesActual = $inicioMesActual.ToString("yyyy-MM-dd")
Write-Log "Ranking 3 meses ($fechaDesde a $fechaHasta) + mes en curso desde $fechaDesdeMesActual..."

Write-Log "Descargando venta neta por cliente (ultimos 3 meses)..."
$sqlRanking = @"
SELECT cliente, SUM(neto) AS ventaNeta
FROM ventas
WHERE fecha >= '$fechaDesde' AND fecha < '$fechaHasta' AND tipo IN ('VEN','DEB')
GROUP BY cliente
"@
$ranking = Invoke-PanelSql $sqlRanking
foreach ($r in $ranking) {
    $codCli = [string]$r.cliente
    if (-not $clienteInfo.ContainsKey($codCli)) { continue }
    $clienteInfo[$codCli].ventaNeta3m = [double]$r.ventaNeta
}

Write-Log "Descargando ventas del mes en curso (cliente x proveedor)..."
$sqlPares = @"
SELECT v.cliente AS cliente, a.proveedor AS proveedor
FROM ventas v
JOIN venta_items i ON i.venta_id = v.id
LEFT JOIN articulos a ON a.codigo = i.articulo
WHERE v.fecha >= '$fechaDesdeMesActual' AND v.fecha < '$fechaHasta'
  AND v.tipo IN ('VEN','DEB')
GROUP BY v.cliente, a.proveedor
"@
$pares = Invoke-PanelSql $sqlPares
foreach ($p in $pares) {
    $codCli = [string]$p.cliente
    if (-not $clienteInfo.ContainsKey($codCli)) { continue }
    $ci = $clienteInfo[$codCli]
    $ci.compro = $true
    if ($p.proveedor) { [void]$ci.proveedoresComprados.Add([string]$p.proveedor) }
}

Write-Log "Descargando nombres de vendedores y proveedores..."
$vendedoresRaw = Invoke-PanelSql "SELECT codigo, nombre FROM vendedores"
$vendNombre = @{}
foreach ($v in $vendedoresRaw) { $vendNombre[[string]$v.codigo] = $v.nombre }
$proveedoresRaw = Invoke-PanelSql "SELECT codigo, nombre FROM proveedores"
$provNombre = @{}
foreach ($p in $proveedoresRaw) { $provNombre[[string]$p.codigo] = $p.nombre }

# --- curva de Pareto por vendedor: marcar "potencial" hasta llegar al corte acumulado ---
$porVendedor = @{}
foreach ($ci in $clienteInfo.Values) {
    if (-not $porVendedor.ContainsKey($ci.codigoVendedor)) { $porVendedor[$ci.codigoVendedor] = @() }
    $porVendedor[$ci.codigoVendedor] += $ci
}
$potencialSet = New-Object System.Collections.Generic.HashSet[string]
foreach ($vcod in $porVendedor.Keys) {
    $clis = @($porVendedor[$vcod] | Where-Object { $_.ventaNeta3m -gt 0 } | Sort-Object ventaNeta3m -Descending)
    $totalVendedor = ($clis | Measure-Object ventaNeta3m -Sum).Sum
    if ($totalVendedor -le 0) { continue }
    $acum = 0.0
    foreach ($ci in $clis) {
        $acum += $ci.ventaNeta3m
        [void]$potencialSet.Add($ci.codigo)
        if (($acum / $totalVendedor) -ge $CorteAcumulado) { break }
    }
}
Write-Log "Clientes potenciales (80% de la venta de su vendedor, ultimos 3 meses): $($potencialSet.Count) de $($clienteInfo.Count)"

$clientesOut = foreach ($cod in $clienteInfo.Keys) {
    if (-not $potencialSet.Contains($cod)) { continue }
    $ci = $clienteInfo[$cod]
    [pscustomobject]@{
        codigo = $ci.codigo; nombre = $ci.nombre; localidad = $ci.localidad
        codigoVendedor = $ci.codigoVendedor; dias = $ci.dias
        ventaNeta3m = [math]::Round($ci.ventaNeta3m,2)
        compro = $ci.compro; proveedoresComprados = @($ci.proveedoresComprados)
    }
}

$vendedoresConPotenciales = New-Object System.Collections.Generic.HashSet[string]
foreach ($c in $clientesOut) { [void]$vendedoresConPotenciales.Add($c.codigoVendedor) }
$vendedoresOut = @($vendedoresConPotenciales | ForEach-Object {
    [pscustomobject]@{ codigo = $_; nombre = if ($vendNombre.ContainsKey($_)) { $vendNombre[$_] } else { "Vendedor $_" } }
} | Sort-Object nombre)

$provConVenta = New-Object System.Collections.Generic.HashSet[string]
foreach ($c in $clientesOut) { foreach ($p in $c.proveedoresComprados) { [void]$provConVenta.Add($p) } }
$proveedoresOut = @($provConVenta | ForEach-Object {
    [pscustomobject]@{ codigo = $_; nombre = if ($provNombre.ContainsKey($_)) { $provNombre[$_] } else { "Proveedor $_" } }
} | Sort-Object nombre)

$out = [pscustomobject]@{
    generatedAt = (Get-Date).ToString("o")
    corteAcumulado = $CorteAcumulado
    rankingDesde = $rankingDesde.ToString("yyyy-MM-dd")
    rankingHasta = $hoy.ToString("yyyy-MM-dd")
    periodoDesde = $inicioMesActual.ToString("yyyy-MM-dd")
    periodoHasta = $hoy.ToString("yyyy-MM-dd")
    vendedores = $vendedoresOut
    proveedores = $proveedoresOut
    clientes = $clientesOut
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
$jsonText = $out | ConvertTo-Json -Depth 8 -Compress
[System.IO.File]::WriteAllText($OutPath, $jsonText, (New-Object System.Text.UTF8Encoding $false))
Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force
Write-Log "Guardado: $OutPath"
$sinCompra = @($clientesOut | Where-Object { -not $_.compro }).Count
Write-Log "Potenciales: $($clientesOut.Count) | sin compra este mes: $sinCompra | vendedores: $($vendedoresOut.Count) | proveedores con venta: $($proveedoresOut.Count)"
