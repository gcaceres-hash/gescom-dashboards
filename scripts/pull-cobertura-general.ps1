<#
Dashboard de cobertura general: por cada vendedor, cuantos clientes de su
cartera compraron vs el universo asignado -- en total y aperturado por
proveedor -- y tambien filtrable por dia de visita de la ruta de preventa
(lunes, martes, etc).

MIGRADO el 30/9/2026 de la API de Gescom a la base compartida de Lucas
(datos-gescom.panelempresas.workers.dev), porque Lucas decidio que esa base es
la UNICA que habla con la API de Gescom (IDEA bloqueo el usuario dos veces por
consumo) y todos los paneles tienen que leer de ahi.

Reglas (mismo criterio que los otros dashboards):
  - Se excluyen los vendedores 1176, 43, 16, 37 al armar la CARTERA (no son
    vendedores reales, un cliente no puede estar "asignado" a ellos) --
    pero una compra de ese cliente SI cuenta aunque el renglon de venta en
    si haya pasado por uno de esos codigos (ej. deposito/logistica). 1176 y 43
    ya vienen excluidos de la tabla `ventas` de origen (son el mostrador).
  - "Compro" = venta VEN o DEB (no DEV-RE/DEV-CA, no AJU/COM), con `fecha`
    dentro del mes. OJO: `fecha` en esta base es fechaPedido (dia de carga en
    Gescom), no fechaComprobante (factura) como usaba la version anterior con
    la API -- la base no guarda fechaComprobante por separado. Es el mismo
    criterio que useLucas valido contra Gescom ("criterio 1 de Lucas": el dia
    de venta es el dia de carga, verificado 100% sobre 2818 ventas).
  - Cartera = clientes.vendedor (ya es el vendedor de la primera ruta de
    preventa, segun arma la base). Los dias de esa misma ruta (primera entrada
    de clientes.rutas) definen "dia de visita".
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/cobertura"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "cobertura-general-template.html"),
    [string]$MesDesde = ""
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

# --- rango del mes a procesar ---
if ($MesDesde -ne "") { $inicioMes = Get-Date $MesDesde } else { $inicioMes = Get-Date -Day 1 }
$inicioMes = Get-Date -Year $inicioMes.Year -Month $inicioMes.Month -Day 1 -Hour 0 -Minute 0 -Second 0
$hoy = (Get-Date).Date
$esMesActual = ($inicioMes.Year -eq $hoy.Year -and $inicioMes.Month -eq $hoy.Month)
$fechaHastaReal = $inicioMes.AddMonths(1)
$fechaDesdeQuery = $inicioMes.ToString("yyyy-MM-dd")
$fechaHastaQuery = $fechaHastaReal.ToString("yyyy-MM-dd")
Write-Log "Procesando mes $($fechaDesdeQuery) a $($fechaHastaQuery) (exclusive)..."

Write-Log "Descargando clientes con cartera asignada..."
$clientesRaw = Invoke-PanelSql "SELECT codigo, nombre, localidad, vendedor, rutas FROM clientes WHERE activo = 1 AND vendedor IS NOT NULL AND vendedor <> ''"

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
        compro = $false; proveedoresComprados = (New-Object System.Collections.Generic.HashSet[string])
    }
    [void]$vendedoresConCartera.Add($vcod)
}
Write-Log "Clientes con cartera asignada (vendedor real): $($clienteInfo.Count) | vendedores distintos: $($vendedoresConCartera.Count)"

Write-Log "Descargando ventas del mes (cliente x proveedor)..."
$sqlPares = @"
SELECT v.cliente AS cliente, a.proveedor AS proveedor
FROM ventas v
JOIN venta_items i ON i.venta_id = v.id
LEFT JOIN articulos a ON a.codigo = i.articulo
WHERE v.fecha >= '$fechaDesdeQuery' AND v.fecha < '$fechaHastaQuery'
  AND v.tipo IN ('VEN','DEB')
GROUP BY v.cliente, a.proveedor
"@
$pares = Invoke-PanelSql $sqlPares
Write-Log "Pares cliente-proveedor: $($pares.Count)"
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

$clientesOut = foreach ($cod in $clienteInfo.Keys) {
    $ci = $clienteInfo[$cod]
    [pscustomobject]@{
        codigo = $ci.codigo; nombre = $ci.nombre; localidad = $ci.localidad
        codigoVendedor = $ci.codigoVendedor; dias = $ci.dias
        compro = $ci.compro; proveedoresComprados = @($ci.proveedoresComprados)
    }
}

$vendedoresOut = @($vendedoresConCartera | ForEach-Object {
    [pscustomobject]@{ codigo = $_; nombre = if ($vendNombre.ContainsKey($_)) { $vendNombre[$_] } else { "Vendedor $_" } }
} | Sort-Object nombre)

$provConVenta = New-Object System.Collections.Generic.HashSet[string]
foreach ($c in $clientesOut) { foreach ($p in $c.proveedoresComprados) { [void]$provConVenta.Add($p) } }
$proveedoresOut = @($provConVenta | ForEach-Object {
    [pscustomobject]@{ codigo = $_; nombre = if ($provNombre.ContainsKey($_)) { $provNombre[$_] } else { "Proveedor $_" } }
} | Sort-Object nombre)

$out = [pscustomobject]@{
    generatedAt = (Get-Date).ToString("o")
    periodoDesde = $inicioMes.ToString("yyyy-MM-dd")
    periodoHasta = $(if ($esMesActual) { $hoy.ToString("yyyy-MM-dd") } else { $inicioMes.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd") })
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
Write-Log "Clientes: $($clientesOut.Count) | sin compra (general): $sinCompra | vendedores: $($vendedoresOut.Count) | proveedores con venta: $($proveedoresOut.Count)"
