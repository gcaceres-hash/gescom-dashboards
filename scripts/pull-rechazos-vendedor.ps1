<#
Arma el dashboard de Rechazos / Devoluciones por vendedor: cuanto peso tienen
los rechazos sobre la venta real de cada vendedor, y cuales son los
proveedores y motivos de rechazo mas preponderantes del mes.

MIGRADO el 30/9/2026 de la API de Gescom a la base compartida de Lucas
(datos-gescom.panelempresas.workers.dev) -- ver pull-cobertura-general.ps1
para el detalle de por que.

Reglas de negocio (mismas que el resto de los dashboards):
  - excluye vendedores 1176, 43, 16, 37 (no son vendedores reales -- esto
    es un reporte agrupado POR VENDEDOR, a diferencia de Rentabilidad). 1176
    y 43 ya vienen excluidos de la tabla `ventas` de origen.
  - "venta real" del vendedor = tipo VEN/DEB, con `fecha` (fechaPedido, ver
    nota de migracion en pull-cobertura-general.ps1) dentro del mes.
  - "rechazo" = comprobantes de tipo DEV-RE (rechazo en la entrega) o DEV-CA
    (devolucion por canje), mismo criterio que Seguimiento Di Giorno.
  - pct = monto de rechazos / venta real del vendedor en el mes.

Uso:
  powershell -File pull-rechazos-vendedor.ps1                        # mes actual
  powershell -File pull-rechazos-vendedor.ps1 -MesDesde 2026-08-01   # un mes especifico
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/rechazos"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "rechazos-vendedor-template.html"),
    [string]$MesDesde = ""
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$EXCLUDED = @("43","1176","16","37")

function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
function Invoke-PanelSql([string]$Sql) {
    $uri = "$($config.baseUrl)/consulta?sql=" + [uri]::EscapeDataString($Sql)
    $headers = @{ Authorization = "Bearer $($config.clave)" }
    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    if ($r.truncado) { Write-Log "AVISO: la consulta vino truncada (mas de 20000 filas) -- revisar y acotar." }
    return $r.filas
}

if ($MesDesde -ne "") { $inicioMes = Get-Date $MesDesde } else { $inicioMes = Get-Date -Day 1 }
$inicioMes = Get-Date -Year $inicioMes.Year -Month $inicioMes.Month -Day 1 -Hour 0 -Minute 0 -Second 0
$finMesCompleto = $inicioMes.AddMonths(1)
$hoy = (Get-Date).Date
$esMesActual = ($inicioMes.Year -eq $hoy.Year -and $inicioMes.Month -eq $hoy.Month)
$fechaDesde = $inicioMes.ToString("yyyy-MM-dd")
$fechaHasta = $finMesCompleto.ToString("yyyy-MM-dd")
$mesKey = $inicioMes.ToString("yyyy-MM")
Write-Log "Procesando mes $mesKey ($fechaDesde a $fechaHasta)..."

Write-Log "Descargando nombres de vendedores y proveedores..."
$vendedoresRaw = Invoke-PanelSql "SELECT codigo, nombre FROM vendedores"
$vendNombre = @{}
foreach ($v in $vendedoresRaw) { $vendNombre[[string]$v.codigo] = $v.nombre }
$proveedoresRaw = Invoke-PanelSql "SELECT codigo, nombre FROM proveedores"
$provNombre = @{}
foreach ($p in $proveedoresRaw) { $provNombre[[string]$p.codigo] = $p.nombre }

Write-Log "Descargando venta neta por vendedor (VEN/DEB)..."
$sqlVenta = @"
SELECT vendedor, SUM(neto) AS ventaNeta
FROM ventas
WHERE fecha >= '$fechaDesde' AND fecha < '$fechaHasta' AND tipo IN ('VEN','DEB')
GROUP BY vendedor
"@
$ventaPorVendedor = Invoke-PanelSql $sqlVenta

Write-Log "Descargando rechazos del mes (DEV-RE/DEV-CA)..."
$sqlRechazos = @"
SELECT id, cliente, vendedor, fecha, nro_comprobante, tipo, motivo, motivo_codigo, neto
FROM ventas
WHERE fecha >= '$fechaDesde' AND fecha < '$fechaHasta' AND tipo IN ('DEV-RE','DEV-CA')
"@
$rechazosRaw = Invoke-PanelSql $sqlRechazos

Write-Log "Descargando proveedores de los items de esos rechazos..."
$sqlItems = @"
SELECT i.venta_id AS venta_id, a.proveedor AS proveedor
FROM ventas v
JOIN venta_items i ON i.venta_id = v.id
LEFT JOIN articulos a ON a.codigo = i.articulo
WHERE v.fecha >= '$fechaDesde' AND v.fecha < '$fechaHasta' AND v.tipo IN ('DEV-RE','DEV-CA')
GROUP BY i.venta_id, a.proveedor
"@
$itemsRaw = Invoke-PanelSql $sqlItems
$provsPorVenta = @{}
foreach ($it in $itemsRaw) {
    $vid = [string]$it.venta_id
    if (-not $provsPorVenta.ContainsKey($vid)) { $provsPorVenta[$vid] = New-Object System.Collections.Generic.HashSet[string] }
    [void]$provsPorVenta[$vid].Add($(if ($it.proveedor) { [string]$it.proveedor } else { "_SIN_PROVEEDOR_" }))
}
Write-Log "Ventas revisadas: venta neta de $($ventaPorVendedor.Count) vendedores | rechazos: $($rechazosRaw.Count)"

$porVendedor = @{}
foreach ($v in $ventaPorVendedor) {
    $cod = [string]$v.vendedor
    if ($EXCLUDED -contains $cod) { continue }
    if (-not $porVendedor.ContainsKey($cod)) { $porVendedor[$cod] = @{ ventaNeta = 0.0; rechazosMonto = 0.0; rechazosCantidad = 0 } }
    $porVendedor[$cod].ventaNeta += [double]$v.ventaNeta
}

$porProveedorRechazo = @{}
$porMotivo = @{}
$eventosPorVendedor = @{}
foreach ($r in $rechazosRaw) {
    $codVend = [string]$r.vendedor
    if ($EXCLUDED -contains $codVend) { continue }
    if (-not $porVendedor.ContainsKey($codVend)) { $porVendedor[$codVend] = @{ ventaNeta = 0.0; rechazosMonto = 0.0; rechazosCantidad = 0 } }

    $monto = [math]::Abs([double]$r.neto)
    $porVendedor[$codVend].rechazosMonto += $monto
    $porVendedor[$codVend].rechazosCantidad += 1

    $vid = [string]$r.id
    $provsVistos = if ($provsPorVenta.ContainsKey($vid)) { $provsPorVenta[$vid] } else { New-Object System.Collections.Generic.HashSet[string] }
    if ($provsVistos.Count -eq 0) { [void]$provsVistos.Add("_SIN_PROVEEDOR_") }
    $montoPorProv = $monto / $provsVistos.Count
    foreach ($p in $provsVistos) {
        if (-not $porProveedorRechazo.ContainsKey($p)) { $porProveedorRechazo[$p] = @{ monto = 0.0; cantidad = 0 } }
        $porProveedorRechazo[$p].monto += $montoPorProv
        $porProveedorRechazo[$p].cantidad += 1
    }

    $motivoDesc = if ($r.motivo) { [string]$r.motivo } else { "Sin motivo especificado" }
    if (-not $porMotivo.ContainsKey($motivoDesc)) { $porMotivo[$motivoDesc] = @{ monto = 0.0; cantidad = 0 } }
    $porMotivo[$motivoDesc].monto += $monto
    $porMotivo[$motivoDesc].cantidad += 1

    if (-not $eventosPorVendedor.ContainsKey($codVend)) { $eventosPorVendedor[$codVend] = @() }
    $eventosPorVendedor[$codVend] += [pscustomobject]@{
        codigoCliente = [string]$r.cliente
        fecha = [string]$r.fecha
        comprobante = [string]$r.nro_comprobante
        tipo = [string]$r.tipo
        motivo = $motivoDesc
        monto = [math]::Round($monto,2)
    }
}

$vendedoresOut = foreach ($cod in $porVendedor.Keys) {
    $v = $porVendedor[$cod]
    $pct = if ($v.ventaNeta -gt 0) { $v.rechazosMonto / $v.ventaNeta } else { 0.0 }
    if ($v.ventaNeta -eq 0 -and $v.rechazosMonto -eq 0) { continue }
    [pscustomobject]@{
        codigo = $cod
        nombre = if ($vendNombre.ContainsKey($cod)) { $vendNombre[$cod] } else { "Vendedor $cod" }
        ventaNeta = [math]::Round($v.ventaNeta,2)
        rechazosMonto = [math]::Round($v.rechazosMonto,2)
        rechazosCantidad = $v.rechazosCantidad
        pct = [math]::Round($pct,4)
        eventos = @(if ($eventosPorVendedor.ContainsKey($cod)) { $eventosPorVendedor[$cod] | Sort-Object fecha -Descending } else { @() })
    }
}
$vendedoresOut = @($vendedoresOut | Sort-Object pct -Descending)

$topProveedores = foreach ($cod in $porProveedorRechazo.Keys) {
    $p = $porProveedorRechazo[$cod]
    [pscustomobject]@{
        codigo = $cod
        nombre = if ($cod -eq "_SIN_PROVEEDOR_") { "Sin proveedor asignado" } else { $provNombre[$cod] }
        monto = [math]::Round($p.monto,2)
        cantidad = $p.cantidad
    }
}
$topProveedores = @($topProveedores | Sort-Object monto -Descending)

$topMotivos = foreach ($mot in $porMotivo.Keys) {
    $m = $porMotivo[$mot]
    [pscustomobject]@{ motivo = $mot; monto = [math]::Round($m.monto,2); cantidad = $m.cantidad }
}
$topMotivos = @($topMotivos | Sort-Object monto -Descending)

$totales = [pscustomobject]@{
    ventaNeta = [math]::Round((($vendedoresOut | Measure-Object ventaNeta -Sum).Sum),2)
    rechazosMonto = [math]::Round((($vendedoresOut | Measure-Object rechazosMonto -Sum).Sum),2)
    rechazosCantidad = ($vendedoresOut | Measure-Object rechazosCantidad -Sum).Sum
}
$totales | Add-Member -NotePropertyName pct -NotePropertyValue $(if ($totales.ventaNeta -gt 0) { [math]::Round($totales.rechazosMonto / $totales.ventaNeta,4) } else { 0.0 })

$mesData = [pscustomobject]@{
    periodoDesde = $inicioMes.ToString("yyyy-MM-dd")
    periodoHasta = $(if ($esMesActual) { $hoy.ToString("yyyy-MM-dd") } else { $finMesCompleto.AddDays(-1).ToString("yyyy-MM-dd") })
    cerrado = -not $esMesActual
    totales = $totales
    vendedores = $vendedoresOut
    topProveedores = $topProveedores
    topMotivos = $topMotivos
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
} catch {
    Write-Log "Aviso: no se pudo leer el data.json existente, se arranca sin historico previo ($($_.Exception.Message))"
}
foreach ($k in @($meses.Keys)) {
    if ($k -ne $mesKey -and -not $meses[$k].cerrado) { $meses[$k] | Add-Member -NotePropertyName cerrado -NotePropertyValue $true -Force }
}
$meses[$mesKey] = $mesData

$out = [pscustomobject]@{
    generatedAt = (Get-Date).ToString("o")
    mesActual = (Get-Date -Format "yyyy-MM")
    meses = $meses
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
$jsonText = $out | ConvertTo-Json -Depth 10 -Compress
[System.IO.File]::WriteAllText($OutPath, $jsonText, (New-Object System.Text.UTF8Encoding $false))
Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force
Write-Log "Guardado: $OutPath"
Write-Log "Mes $mesKey -> Venta neta: $($totales.ventaNeta) | Rechazos: $($totales.rechazosMonto) ($($totales.rechazosCantidad)) | Peso: $([math]::Round($totales.pct*100,2))% | Vendedores: $($vendedoresOut.Count)"
