<#
Dashboard de seguimiento para un vendedor puntual (default: Martin Di Giorno, 076).
Junta en un solo lugar: venta neta por proveedor, rentabilidad total, cobertura de
cartera, rechazos/devoluciones y deuda de sus clientes. Guarda cada mes por
separado (igual que la Pizarra de Rentabilidad) para poder navegar meses.

Reglas (confirmadas con el usuario):
  - "Rechazos" = notas DEV-RE / DEV-CA (mercaderia rechazada o devuelta en la
    entrega) atribuidas a este vendedor: cantidad de eventos + monto.
  - "Cobertura" = % de la cartera asignada (rutasPreventa.codigoVendedor) que
    compro (venta real, no credito) al menos una vez en el periodo.
  - Deuda: mismo criterio que la Torre de Cobranzas (notas de credito restan).

Uso:
  powershell -File pull-vendedor-digiorno.ps1
  powershell -File pull-vendedor-digiorno.ps1 -MesDesde 2026-07-01
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/digiorno"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "vendedor-digiorno-template.html"),
    [string]$CodigoVendedor = "076",
    [string]$MesDesde = ""
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

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

Write-Log "Descargando catalogos (articulos, proveedores, clientes, vendedores)..."
$articulos = Invoke-GescomApi -Path "/data/cmd/inventario/api/v2/get-articulos"
$proveedores = Invoke-GescomApi -Path "/data/cmd/compras/api/v1/get-proveedores"
$clientes = Invoke-GescomApi -Path "/data/cmd/ventas/api/v1/get-clientes"
$vendedores = Invoke-GescomApi -Path "/data/cmd/ventas/api/v1/get-vendedores"

$artInfo = @{}
foreach ($a in $articulos) { $artInfo[[string]$a.codigo] = @{ prov = [string]$a.codigoProveedor } }
$provNombre = @{}
foreach ($p in $proveedores) { $provNombre[[string]$p.codigo] = $p.nombre }
$clienteNombre = @{}
foreach ($c in $clientes) { $clienteNombre[[string]$c.codigo] = $c.nombre }
$vendedorNombre = @{}
foreach ($v in $vendedores) { $vendedorNombre[[string]$v.codigo] = $v.nombre }
$nombreVendedor = $vendedorNombre[$CodigoVendedor]
if (-not $nombreVendedor) { $nombreVendedor = "Vendedor $CodigoVendedor" }

# --- cartera asignada: clientes cuya ruta de preventa lo tiene como vendedor ---
$carteraSet = New-Object System.Collections.Generic.HashSet[string]
foreach ($c in $clientes) {
    foreach ($r in $c.rutasPreventa) {
        if ([string]$r.codigoVendedor -eq $CodigoVendedor) { [void]$carteraSet.Add([string]$c.codigo); break }
    }
}
Write-Log "Cartera asignada a $nombreVendedor : $($carteraSet.Count) clientes"

# --- determinar rango del mes a procesar ---
if ($MesDesde -ne "") { $inicioMes = Get-Date $MesDesde } else { $inicioMes = Get-Date -Day 1 }
$inicioMes = Get-Date -Year $inicioMes.Year -Month $inicioMes.Month -Day 1 -Hour 0 -Minute 0 -Second 0
$finMesCompleto = $inicioMes.AddMonths(1)
$hoy = (Get-Date).Date
$esMesActual = ($inicioMes.Year -eq $hoy.Year -and $inicioMes.Month -eq $hoy.Month)
# el filtro de fechaEntrega va hasta fin de mes SIEMPRE (incluye entregas ya
# facturadas con fecha futura dentro del mes en curso); el limite de la
# CONSULTA a la API (fecha de creacion) si se recorta a hoy en el mes actual.
$fechaHastaReal = $finMesCompleto
# buffer hacia atras: la API filtra por fecha de creacion, pero el criterio real
# es fecha de ENTREGA (una venta creada semanas antes puede entregarse este mes).
# Se acorto de 30 a 12 dias: el circuito entrega/facturacion tarda mucho menos
# que eso en la practica, y una consulta mas angosta es menos propensa a que el
# servidor de Gescom devuelva paginas inconsistentes en consultas muy pesadas.
$BUFFER_DIAS = 12
$fechaDesdeQuery = $inicioMes.AddDays(-$BUFFER_DIAS).ToString("yyyy-MM-dd")
$fechaHastaQuery = $fechaHastaReal.ToString("yyyy-MM-dd")
$mesKey = $inicioMes.ToString("yyyy-MM")
Write-Log "Procesando mes $mesKey (entregas entre $($inicioMes.ToString('yyyy-MM-dd')) y $($fechaHastaReal.ToString('yyyy-MM-dd')), consultando desde $fechaDesdeQuery) para vendedor $CodigoVendedor ($nombreVendedor)..."

$script:token = Get-GescomToken
$porProveedor = @{}
$clientesActivos = New-Object System.Collections.Generic.HashSet[string]
$rechazoEventos = @()
$pagestoskip = 0
$total = 0
$totalVendedor = 0
while ($true) {
    if ($pagestoskip -gt 0 -and $pagestoskip % 5 -eq 0) { $script:token = Get-GescomToken }
    $page = Invoke-GescomApi -Path "/data/cmd/ventas/api/v2/get" -Query @{
        fechadesde = $fechaDesdeQuery; fechahasta = $fechaHastaQuery; pagesize = 500; pagestoskip = $pagestoskip
    }
    if (-not $page -or $page.Count -eq 0) { break }
    $total += $page.Count
    foreach ($venta in $page) {
        if (-not $venta.fechaEntrega) { continue }
        if (-not $venta.comprobantePrincipal -or -not $venta.comprobantePrincipal.fechaComprobante) { continue }
        $fechaEntrega = ([datetime]$venta.fechaEntrega).Date
        $fechaComprobante = ([datetime]$venta.comprobantePrincipal.fechaComprobante).Date
        if ($fechaEntrega -ne $fechaComprobante) { continue }
        if ($fechaEntrega -lt $inicioMes -or $fechaEntrega -ge $fechaHastaReal) { continue }
        if ([string]$venta.codigoVendedor -ne $CodigoVendedor) { continue }
        $totalVendedor++
        $esCredito = $venta.esCredito -eq $true
        $esRechazo = $venta.codigoTipoVenta -in @("DEV-RE", "DEV-CA")

        if ($esRechazo) {
            $rechazoEventos += [pscustomobject]@{
                codigoCliente = [string]$venta.codigoCliente
                nombreCliente = if ($clienteNombre.ContainsKey([string]$venta.codigoCliente)) { $clienteNombre[[string]$venta.codigoCliente] } else { "Cliente $($venta.codigoCliente)" }
                fecha = if ($venta.fechaEntrega) { ([datetime]$venta.fechaEntrega).ToString("yyyy-MM-dd") } else { ([datetime]$venta.fechaPedido).ToString("yyyy-MM-dd") }
                comprobante = [string]$venta.numeroComprobante
                tipo = [string]$venta.codigoTipoVenta
                motivo = [string]$venta.motivo
                monto = [math]::Round([math]::Abs([double]$venta.importeNeto), 2)
            }
        } elseif (-not $esCredito) {
            [void]$clientesActivos.Add([string]$venta.codigoCliente)
        }

        $signo = if ($esCredito) { -1.0 } else { 1.0 }
        foreach ($item in $venta.items) {
            if ([double]$item.precioCosto -eq 1.0) { continue }
            $info = $artInfo[[string]$item.codigoItem]
            $prov = if ($info -and $info.prov) { $info.prov } else { $null }
            if (-not $prov) {
                $prov = "_SIN_PROVEEDOR_"
            }
            $neto = $signo * [double]$item.importeNeto
            $cmvItem = $signo * ([double]$item.cantidad * [double]$item.precioCosto)
            $desc = if (-not $esCredito) { [double]$item.descuentoNeto } else { 0.0 }

            if (-not $porProveedor.ContainsKey($prov)) { $porProveedor[$prov] = @{ ventaNeta=0.0; cmv=0.0; descuentos=0.0 } }
            $porProveedor[$prov].ventaNeta += $neto
            $porProveedor[$prov].cmv += $cmvItem
            $porProveedor[$prov].descuentos += $desc
        }
    }
    if ($page.Count -lt 500) { break }
    $pagestoskip++
}
Write-Log "Ventas totales revisadas: $total | de $nombreVendedor : $totalVendedor"

$proveedoresOut = foreach ($cod in $porProveedor.Keys) {
    $p = $porProveedor[$cod]
    $nombre = if ($cod -eq "_SIN_PROVEEDOR_") { "Sin proveedor asignado" } else { $provNombre[$cod] }
    if (-not $nombre) { $nombre = "Proveedor $cod" }
    [pscustomobject]@{
        codigo = $cod
        nombre = $nombre
        ventaNeta = [math]::Round($p.ventaNeta,2)
        cmv = [math]::Round($p.cmv,2)
        descuentos = [math]::Round($p.descuentos,2)
    }
}
$ventaTotales = [pscustomobject]@{
    ventaNeta = [math]::Round((($proveedoresOut | Measure-Object ventaNeta -Sum).Sum),2)
    cmv = [math]::Round((($proveedoresOut | Measure-Object cmv -Sum).Sum),2)
    descuentos = [math]::Round((($proveedoresOut | Measure-Object descuentos -Sum).Sum),2)
}

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

$rechazos = [pscustomobject]@{
    cantidad = $rechazoEventos.Count
    monto = [math]::Round((($rechazoEventos | Measure-Object monto -Sum).Sum),2)
    eventos = @($rechazoEventos | Sort-Object fecha -Descending)
}

# --- deuda: mismo criterio que la Torre de Cobranzas, filtrado a este vendedor ---
Write-Log "Descargando cuenta corriente y resolviendo vendedor por venta..."
$detalle = Invoke-GescomApi -Path "/data/cmd/ctacte/api/v3/get-ctacte-clientes-detalle"
$ventaIds = New-Object System.Collections.Generic.HashSet[string]
foreach ($cli in $detalle) { foreach ($c in $cli.comprobantes) { if ($c.ventaId) { [void]$ventaIds.Add([string]$c.ventaId) } } }
$ventaIdsList = @($ventaIds)
$ventaVendedor = @{}
$batchSize = 50
for ($i = 0; $i -lt $ventaIdsList.Count; $i += $batchSize) {
    $batch = $ventaIdsList[$i..([math]::Min($i+$batchSize-1, $ventaIdsList.Count-1))]
    $idsParam = $batch -join ","
    if (($i / $batchSize) % 10 -eq 0 -and $i -gt 0) { $script:token = Get-GescomToken }
    try {
        $r = Invoke-GescomApi -Path "/data/cmd/ventas/api/v2/get" -Query @{ ids = $idsParam; pagesize = 50 }
        foreach ($v in $r) { $ventaVendedor[[string]$v.id] = [string]$v.codigoVendedor }
    } catch { Write-Log "Aviso: fallo batch de ventaId en indice $i" }
}

$hoyDate = (Get-Date).Date
function Bucket($diasVencido) {
    if ($diasVencido -le 0) { return "vigente" }
    if ($diasVencido -le 30) { return "d1_30" }
    if ($diasVencido -le 60) { return "d31_60" }
    if ($diasVencido -le 90) { return "d61_90" }
    return "d90mas"
}
$deudaClientes = @{}
foreach ($cli in $detalle) {
    $codCliente = [string]$cli.codigoCliente
    foreach ($c in $cli.comprobantes) {
        $vendCod = if ($c.ventaId -and $ventaVendedor.ContainsKey([string]$c.ventaId)) { $ventaVendedor[[string]$c.ventaId] } else { $null }
        if ($vendCod -ne $CodigoVendedor) { continue }
        $esCreditoC = $null -ne $c.creditoId
        $saldo = [double]$c.saldo * $(if ($esCreditoC) { -1 } else { 1 })
        $fv = if ($c.fechaVencimiento) { [datetime]$c.fechaVencimiento } else { $hoyDate }
        $diasVencido = ($hoyDate - $fv.Date).Days
        if (-not $deudaClientes.ContainsKey($codCliente)) { $deudaClientes[$codCliente] = @() }
        $deudaClientes[$codCliente] += [pscustomobject]@{
            comprobante = $c.comprobante; saldo = [math]::Round($saldo,2); esCredito = $esCreditoC
            fechaVencimiento = if ($c.fechaVencimiento) { $fv.ToString("yyyy-MM-dd") } else { $null }
            diasVencido = $diasVencido; bucket = Bucket $diasVencido
        }
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

# --- cargar historico existente (del data.json ya commiteado en el repo) y fusionar el mes procesado ---
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
