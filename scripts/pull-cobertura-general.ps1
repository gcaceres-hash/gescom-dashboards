<#
Dashboard de cobertura general: por cada vendedor, cuantos clientes de su
cartera compraron vs el universo asignado -- en total y aperturado por
proveedor -- y tambien filtrable por dia de visita de la ruta de preventa
(lunes, martes, etc).

Reglas (mismo criterio que los otros dashboards):
  - Se excluyen los vendedores 1176, 43, 16, 37 (no son vendedores reales).
  - "Compro" = venta cerrada (facturada fiscalmente), no nota de credito, en el mes en curso.
  - Cartera = clientes cuya ruta de preventa (primera entrada) tiene a ese
    vendedor asignado. Los dias de esa misma ruta definen "dia de visita".
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/cobertura"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "cobertura-general-template.html"),
    [string]$MesDesde = ""
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$EXCLUDED = @("43","1176","16","37")
$DIAS_SEMANA = @("lunes","martes","miercoles","jueves","viernes","sabado","domingo")

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

Write-Log "Descargando catalogos (clientes, vendedores, proveedores, articulos)..."
$clientesRaw = Invoke-GescomApi -Path "/data/cmd/ventas/api/v1/get-clientes"
$vendedoresRaw = Invoke-GescomApi -Path "/data/cmd/ventas/api/v1/get-vendedores"
$proveedoresRaw = Invoke-GescomApi -Path "/data/cmd/compras/api/v1/get-proveedores"
$articulos = Invoke-GescomApi -Path "/data/cmd/inventario/api/v2/get-articulos"

$provNombre = @{}
foreach ($p in $proveedoresRaw) { $provNombre[[string]$p.codigo] = $p.nombre }
$vendNombre = @{}
foreach ($v in $vendedoresRaw) { $vendNombre[[string]$v.codigo] = $v.nombre }
$artProv = @{}
foreach ($a in $articulos) { $artProv[[string]$a.codigo] = [string]$a.codigoProveedor }

# --- armar cliente -> {vendedor, dias} a partir de la primera ruta de preventa ---
$clienteInfo = @{}
$vendedoresConCartera = New-Object System.Collections.Generic.HashSet[string]
foreach ($c in $clientesRaw) {
    $ruta = $c.rutasPreventa | Select-Object -First 1
    if (-not $ruta) { continue }
    $vcod = [string]$ruta.codigoVendedor
    if (-not $vcod -or $EXCLUDED -contains $vcod) { continue }
    $dias = @($DIAS_SEMANA | Where-Object { $ruta.$_ -eq $true })
    $clienteInfo[[string]$c.codigo] = [pscustomobject]@{
        codigo = [string]$c.codigo; nombre = $c.nombre; localidad = $c.localidad
        codigoVendedor = $vcod; dias = $dias
        compro = $false; proveedoresComprados = (New-Object System.Collections.Generic.HashSet[string])
    }
    [void]$vendedoresConCartera.Add($vcod)
}
Write-Log "Clientes con cartera asignada (vendedor real): $($clienteInfo.Count) | vendedores distintos: $($vendedoresConCartera.Count)"

# --- rango del mes a procesar ---
if ($MesDesde -ne "") { $inicioMes = Get-Date $MesDesde } else { $inicioMes = Get-Date -Day 1 }
$inicioMes = Get-Date -Year $inicioMes.Year -Month $inicioMes.Month -Day 1 -Hour 0 -Minute 0 -Second 0
$hoy = (Get-Date).Date
$esMesActual = ($inicioMes.Year -eq $hoy.Year -and $inicioMes.Month -eq $hoy.Month)
# el filtro de fechaEntrega va hasta fin de mes SIEMPRE (incluye entregas ya
# facturadas con fecha futura dentro del mes en curso); el limite de la
# CONSULTA a la API (fecha de creacion) si se recorta a hoy en el mes actual.
$fechaHastaReal = $inicioMes.AddMonths(1)
# buffer hacia atras: la API filtra por fecha de creacion, pero el criterio real
# es fecha de ENTREGA (una venta creada semanas antes puede entregarse este mes).
# OJO: se probo acortar esto a 12 dias pero se comprobo que hay ventas cuyo
# circuito entrega/facturacion tarda mas que eso -- se dejaban clientes/ventas
# reales afuera. Se mantiene en 30 dias.
$BUFFER_DIAS = 30
$fechaDesdeQuery = $inicioMes.AddDays(-$BUFFER_DIAS).ToString("yyyy-MM-dd")
$fechaHastaQuery = $fechaHastaReal.ToString("yyyy-MM-dd")
Write-Log "Revisando entregas del mes ($($inicioMes.ToString('yyyy-MM-dd')) a $($fechaHastaReal.ToString('yyyy-MM-dd')), consultando desde $fechaDesdeQuery)..."

$script:token = Get-GescomToken
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
        if (-not $venta.fechaEntrega) { continue }
        if (-not $venta.comprobantePrincipal -or -not $venta.comprobantePrincipal.fechaComprobante) { continue }
        $fechaEntrega = ([datetime]$venta.fechaEntrega).Date
        $fechaComprobante = ([datetime]$venta.comprobantePrincipal.fechaComprobante).Date
        if ($fechaEntrega -ne $fechaComprobante) { continue }
        if ($fechaEntrega -lt $inicioMes -or $fechaEntrega -ge $fechaHastaReal) { continue }
        if ($venta.esCredito -eq $true) { continue }
        $codCli = [string]$venta.codigoCliente
        if (-not $clienteInfo.ContainsKey($codCli)) { continue }
        $ci = $clienteInfo[$codCli]
        $ci.compro = $true
        foreach ($item in $venta.items) {
            $prov = $artProv[[string]$item.codigoItem]
            if ($prov) { [void]$ci.proveedoresComprados.Add($prov) }
        }
    }
    if ($page.Count -lt 500) { break }
    $pagestoskip++
}
Write-Log "Ventas del mes revisadas: $total"

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

# solo proveedores que efectivamente tuvieron alguna venta este mes (para no ensuciar el selector)
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
