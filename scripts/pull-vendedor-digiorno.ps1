<#
Dashboard de seguimiento para un vendedor puntual (default: Martin Di Giorno, 076).
Junta en un solo lugar: venta neta por proveedor, rentabilidad total, cobertura de
cartera, rechazos/devoluciones y deuda de sus clientes. Guarda cada mes por
separado (igual que la Pizarra de Rentabilidad) para poder navegar meses.

MIGRADO el 30/9/2026 (venta/cobertura/rechazos) y 1/10/2026 (deuda) a la
base compartida de Lucas (datos-gescom.panelempresas.workers.dev) -- ver
pull-cobertura-general.ps1 para el detalle de por que. La seccion Deuda usa
ctacte_clientes, que Lucas agrego el 1/10/2026 (ya trae vendedor resuelto
por fila, no hace falta el batch de ventaId->vendedor de la version vieja).

Reglas (confirmadas con el usuario):
  - "Rechazos" = notas DEV-RE / DEV-CA (mercaderia rechazada o devuelta en la
    entrega) atribuidas a este vendedor: cantidad de eventos + monto.
  - "Cobertura" = % de la cartera asignada (clientes.vendedor) que compro
    (tipo VEN/DEB) al menos una vez en el periodo.
  - "fecha" en esta base es fechaPedido (dia de carga), no fechaComprobante
    -- ver la nota de migracion en pull-cobertura-general.ps1.

Uso:
  powershell -File pull-vendedor-digiorno.ps1
  powershell -File pull-vendedor-digiorno.ps1 -MesDesde 2026-07-01
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/digiorno"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "vendedor-digiorno-template.html"),
    [string]$CodigoVendedor = "076",
    [string]$MesDesde = ""
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
function Invoke-PanelSql([string]$Sql) {
    $uri = "$($config.baseUrl)/consulta?sql=" + [uri]::EscapeDataString($Sql)
    $headers = @{ Authorization = "Bearer $($config.clave)" }
    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    if ($r.truncado) { Write-Log "AVISO: la consulta vino truncada (mas de 20000 filas) -- revisar y acotar." }
    return $r.filas
}

Write-Log "Descargando nombres de proveedores, clientes y vendedores..."
$proveedores = Invoke-PanelSql "SELECT codigo, nombre FROM proveedores"
$provNombre = @{}
foreach ($p in $proveedores) { $provNombre[[string]$p.codigo] = $p.nombre }
$clientes = Invoke-PanelSql "SELECT codigo, nombre, vendedor FROM clientes WHERE activo = 1"
$clienteNombre = @{}
foreach ($c in $clientes) { $clienteNombre[[string]$c.codigo] = $c.nombre }
$vendedores = Invoke-PanelSql "SELECT codigo, nombre FROM vendedores"
$vendedorNombre = @{}
foreach ($v in $vendedores) { $vendedorNombre[[string]$v.codigo] = $v.nombre }
$nombreVendedor = $vendedorNombre[$CodigoVendedor]
if (-not $nombreVendedor) { $nombreVendedor = "Vendedor $CodigoVendedor" }

$carteraSet = New-Object System.Collections.Generic.HashSet[string]
foreach ($c in $clientes) { if ([string]$c.vendedor -eq $CodigoVendedor) { [void]$carteraSet.Add([string]$c.codigo) } }
Write-Log "Cartera asignada a $nombreVendedor : $($carteraSet.Count) clientes"

if ($MesDesde -ne "") { $inicioMes = Get-Date $MesDesde } else { $inicioMes = Get-Date -Day 1 }
$inicioMes = Get-Date -Year $inicioMes.Year -Month $inicioMes.Month -Day 1 -Hour 0 -Minute 0 -Second 0
$finMesCompleto = $inicioMes.AddMonths(1)
$hoy = (Get-Date).Date
$esMesActual = ($inicioMes.Year -eq $hoy.Year -and $inicioMes.Month -eq $hoy.Month)
$fechaDesde = $inicioMes.ToString("yyyy-MM-dd")
$fechaHasta = $finMesCompleto.ToString("yyyy-MM-dd")
$mesKey = $inicioMes.ToString("yyyy-MM")
Write-Log "Procesando mes $mesKey ($fechaDesde a $fechaHasta) para vendedor $CodigoVendedor ($nombreVendedor)..."

Write-Log "Descargando venta por proveedor (VEN/DEB/DEV-RE/DEV-CA)..."
$sqlVenta = @"
SELECT COALESCE(a.proveedor, '_SIN_PROVEEDOR_') AS proveedor,
       SUM(CASE WHEN v.tipo IN ('VEN','DEB') THEN i.neto WHEN v.tipo IN ('DEV-RE','DEV-CA') THEN -i.neto ELSE 0 END) AS ventaNeta,
       SUM(CASE WHEN v.tipo IN ('VEN','DEB') THEN i.cantidad*i.precio_costo WHEN v.tipo IN ('DEV-RE','DEV-CA') THEN -i.cantidad*i.precio_costo ELSE 0 END) AS cmv,
       SUM(CASE WHEN v.tipo IN ('VEN','DEB') THEN (i.cantidad*i.precio_unitario - i.neto) ELSE 0 END) AS descuentos
FROM ventas v
JOIN venta_items i ON i.venta_id = v.id
LEFT JOIN articulos a ON a.codigo = i.articulo
WHERE v.vendedor = '$CodigoVendedor' AND v.fecha >= '$fechaDesde' AND v.fecha < '$fechaHasta'
  AND v.tipo IN ('VEN','DEB','DEV-RE','DEV-CA') AND i.precio_costo <> 1
GROUP BY COALESCE(a.proveedor, '_SIN_PROVEEDOR_')
"@
$filasVenta = Invoke-PanelSql $sqlVenta

$proveedoresOut = foreach ($f in $filasVenta) {
    $cod = [string]$f.proveedor
    $nombre = if ($cod -eq "_SIN_PROVEEDOR_") { "Sin proveedor asignado" } else { $provNombre[$cod] }
    if (-not $nombre) { $nombre = "Proveedor $cod" }
    [pscustomobject]@{
        codigo = $cod
        nombre = $nombre
        ventaNeta = [math]::Round([double]$f.ventaNeta,2)
        cmv = [math]::Round([double]$f.cmv,2)
        descuentos = [math]::Round([double]$f.descuentos,2)
    }
}
$ventaTotales = [pscustomobject]@{
    ventaNeta = [math]::Round((($proveedoresOut | Measure-Object ventaNeta -Sum).Sum),2)
    cmv = [math]::Round((($proveedoresOut | Measure-Object cmv -Sum).Sum),2)
    descuentos = [math]::Round((($proveedoresOut | Measure-Object descuentos -Sum).Sum),2)
}

Write-Log "Descargando cobertura (clientes activos este mes)..."
$sqlActivos = @"
SELECT DISTINCT cliente FROM ventas
WHERE vendedor = '$CodigoVendedor' AND fecha >= '$fechaDesde' AND fecha < '$fechaHasta' AND tipo IN ('VEN','DEB')
"@
$activosRaw = Invoke-PanelSql $sqlActivos
$clientesActivos = New-Object System.Collections.Generic.HashSet[string]
foreach ($a in $activosRaw) { [void]$clientesActivos.Add([string]$a.cliente) }
$cobertura = [pscustomobject]@{
    carteraTotal = $carteraSet.Count
    clientesActivos = $clientesActivos.Count
    pct = if ($carteraSet.Count -gt 0) { [math]::Round($clientesActivos.Count / $carteraSet.Count, 4) } else { 0 }
    clientesSinCompra = @(
        $carteraSet | Where-Object { -not $clientesActivos.Contains($_) } | ForEach-Object {
            [pscustomobject]@{ codigo = $_; nombre = if ($clienteNombre.ContainsKey($_)) { $clienteNombre[$_] } else { "Cliente $_" } }
        } | Sort-Object nombre
    )
}

Write-Log "Descargando rechazos de este vendedor..."
$sqlRechazos = @"
SELECT cliente, entrega, fecha, nro_comprobante, tipo, motivo, neto
FROM ventas
WHERE vendedor = '$CodigoVendedor' AND fecha >= '$fechaDesde' AND fecha < '$fechaHasta' AND tipo = 'DEV-RE'
"@
# Solo DEV-RE: los DEV-CA (canjes/cambios de Ilolay por vencimiento) no son rechazos
# (ver pull-rechazos-vendedor.ps1). Siguen restando de la venta neta del vendedor.
$rechazosRaw = Invoke-PanelSql $sqlRechazos
$rechazoEventos = foreach ($r in $rechazosRaw) {
    [pscustomobject]@{
        codigoCliente = [string]$r.cliente
        nombreCliente = if ($clienteNombre.ContainsKey([string]$r.cliente)) { $clienteNombre[[string]$r.cliente] } else { "Cliente $($r.cliente)" }
        fecha = if ($r.entrega) { [string]$r.entrega } else { [string]$r.fecha }
        comprobante = [string]$r.nro_comprobante
        tipo = [string]$r.tipo
        motivo = [string]$r.motivo
        monto = [math]::Round([math]::Abs([double]$r.neto), 2)
    }
}
$rechazos = [pscustomobject]@{
    cantidad = @($rechazoEventos).Count
    monto = [math]::Round((($rechazoEventos | Measure-Object monto -Sum).Sum),2)
    eventos = @($rechazoEventos | Sort-Object fecha -Descending)
}

# --- DEUDA: cuenta corriente de clientes de este vendedor, via ctacte_clientes ---
Write-Log "Descargando cuenta corriente de este vendedor..."
function Bucket($diasVencido) {
    if ($diasVencido -le 0) { return "vigente" }
    if ($diasVencido -le 30) { return "d1_30" }
    if ($diasVencido -le 60) { return "d31_60" }
    if ($diasVencido -le 90) { return "d61_90" }
    return "d90mas"
}
$sqlDeuda = "SELECT cliente, comprobante, numero, fecha, vence, saldo FROM ctacte_clientes WHERE vendedor = '$CodigoVendedor'"
$deudaRaw = Invoke-PanelSql $sqlDeuda
$deudaClientes = @{}
foreach ($c in $deudaRaw) {
    $codCliente = [string]$c.cliente
    $saldo = [double]$c.saldo
    $fv = if ($c.vence) { [datetime]$c.vence } else { $hoy }
    $diasVencido = ($hoy - $fv.Date).Days
    if (-not $deudaClientes.ContainsKey($codCliente)) { $deudaClientes[$codCliente] = @() }
    $deudaClientes[$codCliente] += [pscustomobject]@{
        comprobante = "$($c.comprobante) $($c.numero)"; saldo = [math]::Round($saldo,2); esCredito = $saldo -lt 0
        fechaVencimiento = if ($c.vence) { [string]$c.vence } else { $null }
        diasVencido = $diasVencido; bucket = Bucket $diasVencido
    }
}
$deudaClientesOut = foreach ($cod in $deudaClientes.Keys) {
    $comps = $deudaClientes[$cod]
    [pscustomobject]@{
        codigo = $cod
        nombre = if ($clienteNombre.ContainsKey($cod)) { $clienteNombre[$cod] } else { "Cliente $cod" }
        saldoTotal = [math]::Round((($comps | Measure-Object saldo -Sum).Sum),2)
        peorDiasVencido = ($comps | Measure-Object diasVencido -Maximum).Maximum
        comprobantes = @($comps | Sort-Object fechaVencimiento)
    }
}
$deudaClientesOut = @($deudaClientesOut | Sort-Object saldoTotal -Descending)
$deudaBuckets = [ordered]@{ vigente=0.0; d1_30=0.0; d31_60=0.0; d61_90=0.0; d90mas=0.0 }
foreach ($cli in $deudaClientesOut) { foreach ($cp in $cli.comprobantes) { $deudaBuckets[$cp.bucket] += $cp.saldo } }
$deuda = [pscustomobject]@{
    saldoTotal = [math]::Round((($deudaClientesOut | Measure-Object saldoTotal -Sum).Sum),2)
    vigente = [math]::Round($deudaBuckets.vigente,2); d1_30 = [math]::Round($deudaBuckets.d1_30,2)
    d31_60 = [math]::Round($deudaBuckets.d31_60,2); d61_90 = [math]::Round($deudaBuckets.d61_90,2); d90mas = [math]::Round($deudaBuckets.d90mas,2)
    clientes = $deudaClientesOut
}

$mesData = [pscustomobject]@{
    periodoDesde = $inicioMes.ToString("yyyy-MM-dd")
    periodoHasta = $(if ($esMesActual) { $hoy.ToString("yyyy-MM-dd") } else { $finMesCompleto.AddDays(-1).ToString("yyyy-MM-dd") })
    cerrado = -not $esMesActual
    venta = [pscustomobject]@{ totales = $ventaTotales; proveedores = @($proveedoresOut | Sort-Object ventaNeta -Descending) }
    cobertura = $cobertura
    rechazos = $rechazos
}

$meses = [ordered]@{}
try {
    if (Test-Path $OutPath) {
        $rawText = [System.IO.File]::ReadAllText($OutPath, [System.Text.Encoding]::UTF8)
        $prev = $rawText | ConvertFrom-Json
        if ($prev.meses) {
            foreach ($prop in $prev.meses.PSObject.Properties) { $meses[$prop.Name] = $prop.Value }
            Write-Log "Historico cargado desde $OutPath : $($meses.Keys -join ', ')"
        }
    } else {
        Write-Log "No existe data.json previo en $OutPath, se arranca sin historico."
    }
} catch { Write-Log "Aviso: no se pudo leer el data.json existente, se arranca sin historico previo" }
foreach ($k in @($meses.Keys)) {
    if ($k -ne $mesKey -and -not $meses[$k].cerrado) { $meses[$k] | Add-Member -NotePropertyName cerrado -NotePropertyValue $true -Force }
}
$meses[$mesKey] = $mesData

$out = [pscustomobject]@{
    generatedAt = (Get-Date).ToString("o")
    mesActual = (Get-Date -Format "yyyy-MM")
    vendedor = [pscustomobject]@{ codigo = $CodigoVendedor; nombre = $nombreVendedor }
    deuda = $deuda
    meses = $meses
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
$jsonText = $out | ConvertTo-Json -Depth 12 -Compress
[System.IO.File]::WriteAllText($OutPath, $jsonText, (New-Object System.Text.UTF8Encoding $false))
Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force
Write-Log "Guardado: $OutPath"
Write-Log "Mes $mesKey -> Venta neta: $($ventaTotales.ventaNeta) | CMV: $($ventaTotales.cmv) | Cobertura: $($cobertura.clientesActivos)/$($cobertura.carteraTotal) | Rechazos: $($rechazos.cantidad) ($($rechazos.monto)) | Deuda: $($deuda.saldoTotal)"
