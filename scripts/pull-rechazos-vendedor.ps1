<#
Arma el dashboard de Rechazos / Devoluciones por vendedor: cuanto peso tienen
los rechazos sobre la venta real de cada vendedor, y cuales son los
proveedores y motivos de rechazo mas preponderantes del mes.

Reglas de negocio (mismas que el resto de los dashboards):
  - excluye vendedores 1176, 43, 16, 37 (no son vendedores reales -- esto
    es un reporte agrupado POR VENDEDOR, a diferencia de Rentabilidad)
  - "venta real" del vendedor = venta cerrada, no rechazo, con fecha de
    COMPROBANTE (factura) dentro del mes -- ya no se exige que coincida
    con la fecha de entrega (se confirmo contra la tabla dinamica nativa
    de Gescom que esa igualdad descartaba ventas reales)
  - "rechazo" = comprobantes de tipo DEV-RE (rechazo en la entrega) o DEV-CA
    (devolucion por canje), mismo criterio que Seguimiento Di Giorno
  - pct = monto de rechazos / venta real del vendedor en el mes

Uso:
  powershell -File pull-rechazos-vendedor.ps1                        # mes actual
  powershell -File pull-rechazos-vendedor.ps1 -MesDesde 2026-08-01   # un mes especifico
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/rechazos"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "rechazos-vendedor-template.html"),
    [string]$MesDesde = ""
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$EXCLUDED = @("43","1176","16","37")

function Get-GescomToken {
    $tokenUrl = "$($config.authUrl)/realms/$($config.realm)/protocol/openid-connect/token"
    $body = @{ client_id = $config.clientId; username = $config.usuario; password = $config.clave; grant_type = "password" }
    (Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded").access_token
}
function Invoke-GescomApi {
    param([string]$Path, [hashtable]$Query = @{})
    $qs = ($Query.GetEnumerator() | ForEach-Object { "$($_.Key)=$([uri]::EscapeDataString([string]$_.Value))" }) -join "&"
    $url = "$($config.baseUrl)$Path"
    if ($qs) { $url += "?$qs" }
    for ($intento = 1; $intento -le 4; $intento++) {
        try {
            return Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $script:token" } -Method Get
        } catch {
            if ($intento -eq 4) { throw }
            Start-Sleep -Seconds ($intento * 5)
        }
    }
}
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }

Write-Log "Autenticando..."
$script:token = Get-GescomToken

Write-Log "Descargando catalogos (proveedores, articulos, vendedores)..."
$proveedores = Invoke-GescomApi -Path "/data/cmd/compras/api/v1/get-proveedores"
$articulos = Invoke-GescomApi -Path "/data/cmd/inventario/api/v2/get-articulos"
$vendedoresRaw = Invoke-GescomApi -Path "/data/cmd/ventas/api/v1/get-vendedores"
$motivosVenta = Invoke-GescomApi -Path "/data/cmd/ventas/api/v1/get-motivos-venta"

$provNombre = @{}
foreach ($p in $proveedores) { $provNombre[[string]$p.codigo] = $p.nombre }
$artProv = @{}
foreach ($a in $articulos) { $artProv[[string]$a.codigo] = [string]$a.codigoProveedor }
$vendNombre = @{}
foreach ($v in $vendedoresRaw) { $vendNombre[[string]$v.codigo] = $v.nombre }
$motivoNombre = @{}
foreach ($m in $motivosVenta) { $motivoNombre["$($m.codigoTipoVenta)|$($m.codigo)"] = $m.descripcion }

# --- rango del mes a procesar ---
if ($MesDesde -ne "") { $inicioMes = Get-Date $MesDesde } else { $inicioMes = Get-Date -Day 1 }
$inicioMes = Get-Date -Year $inicioMes.Year -Month $inicioMes.Month -Day 1 -Hour 0 -Minute 0 -Second 0
$finMesCompleto = $inicioMes.AddMonths(1)
$hoy = (Get-Date).Date
$esMesActual = ($inicioMes.Year -eq $hoy.Year -and $inicioMes.Month -eq $hoy.Month)
$fechaHastaReal = $finMesCompleto
$BUFFER_DIAS = 30
$fechaDesdeQuery = $inicioMes.AddDays(-$BUFFER_DIAS).ToString("yyyy-MM-dd")
$fechaHastaQuery = $fechaHastaReal.ToString("yyyy-MM-dd")
$mesKey = $inicioMes.ToString("yyyy-MM")
Write-Log "Procesando mes $mesKey (entregas entre $($inicioMes.ToString('yyyy-MM-dd')) y $($fechaHastaReal.ToString('yyyy-MM-dd')), consultando desde $fechaDesdeQuery)..."

$script:token = Get-GescomToken
$porVendedor = @{}
$porProveedorRechazo = @{}
$porMotivo = @{}
$eventosPorVendedor = @{}
$pagestoskip = 0
$total = 0
while ($true) {
    if ($pagestoskip -gt 0 -and $pagestoskip % 5 -eq 0) { $script:token = Get-GescomToken }
    $page = Invoke-GescomApi -Path "/data/cmd/ventas/api/v2/get" -Query @{
        fechadesde = $fechaDesdeQuery; fechahasta = $fechaHastaQuery; pagesize = 500; pagestoskip = $pagestoskip
    }
    if (-not $page -or $page.Count -eq 0) { break }
    $total += $page.Count
    foreach ($venta in $page) {
        if (-not $venta.comprobantePrincipal -or -not $venta.comprobantePrincipal.fechaComprobante) { continue }
        $fechaComprobante = ([datetime]$venta.comprobantePrincipal.fechaComprobante).Date
        if ($fechaComprobante -lt $inicioMes -or $fechaComprobante -ge $fechaHastaReal) { continue }
        $codVend = [string]$venta.codigoVendedor
        if ($EXCLUDED -contains $codVend) { continue }

        if (-not $porVendedor.ContainsKey($codVend)) {
            $porVendedor[$codVend] = @{ ventaNeta = 0.0; rechazosMonto = 0.0; rechazosCantidad = 0 }
        }

        $esRechazo = $venta.codigoTipoVenta -in @("DEV-RE", "DEV-CA")
        $esCredito = $venta.esCredito -eq $true
        $importeVenta = [double]$venta.importeNeto

        if ($esRechazo) {
            $monto = [math]::Abs($importeVenta)
            $porVendedor[$codVend].rechazosMonto += $monto
            $porVendedor[$codVend].rechazosCantidad += 1

            # proveedor(es) principal(es) de los items de esta venta rechazada
            $provsVistos = New-Object System.Collections.Generic.HashSet[string]
            foreach ($item in $venta.items) {
                $p = $artProv[[string]$item.codigoItem]
                if (-not $p) { $p = "_SIN_PROVEEDOR_" }
                [void]$provsVistos.Add($p)
            }
            if ($provsVistos.Count -eq 0) { [void]$provsVistos.Add("_SIN_PROVEEDOR_") }
            $montoPorProv = $monto / $provsVistos.Count
            foreach ($p in $provsVistos) {
                if (-not $porProveedorRechazo.ContainsKey($p)) { $porProveedorRechazo[$p] = @{ monto = 0.0; cantidad = 0 } }
                $porProveedorRechazo[$p].monto += $montoPorProv
                $porProveedorRechazo[$p].cantidad += 1
            }

            $motivoKey = "$($venta.codigoTipoVenta)|$($venta.motivo)"
            $motivoDesc = if ($motivoNombre.ContainsKey($motivoKey)) { $motivoNombre[$motivoKey] } else { [string]$venta.motivo }
            if (-not $motivoDesc) { $motivoDesc = "Sin motivo especificado" }
            if (-not $porMotivo.ContainsKey($motivoDesc)) { $porMotivo[$motivoDesc] = @{ monto = 0.0; cantidad = 0 } }
            $porMotivo[$motivoDesc].monto += $monto
            $porMotivo[$motivoDesc].cantidad += 1

            if (-not $eventosPorVendedor.ContainsKey($codVend)) { $eventosPorVendedor[$codVend] = @() }
            $eventosPorVendedor[$codVend] += [pscustomobject]@{
                codigoCliente = [string]$venta.codigoCliente
                fecha = $fechaComprobante.ToString("yyyy-MM-dd")
                comprobante = [string]$venta.numeroComprobante
                tipo = [string]$venta.codigoTipoVenta
                motivo = $motivoDesc
                monto = [math]::Round($monto,2)
            }
        } elseif (-not $esCredito) {
            $porVendedor[$codVend].ventaNeta += $importeVenta
        }
    }
    if ($page.Count -lt 500) { break }
    $pagestoskip++
}
Write-Log "Ventas revisadas: $total"

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

# --- cargar historico existente y fusionar el mes procesado ---
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
