<#
Dashboard de cobertura general: por cada vendedor, cuantos clientes de su
cartera compraron vs el universo asignado -- en total y aperturado por
proveedor -- y tambien filtrable por dia de visita de la ruta de preventa
(lunes, martes, etc).

Reglas (mismo criterio que los otros dashboards):
  - Se excluyen los vendedores 1176, 43, 16, 37 (no son vendedores reales).
  - "Compro" = venta Finalizada, no nota de credito, en el mes en curso.
  - Cartera = clientes cuya ruta de preventa (primera entrada) tiene a ese
    vendedor asignado. Los dias de esa misma ruta definen "dia de visita".
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string]$OutPath = (Join-Path $PSScriptRoot "cobertura-general-data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "cobertura-general-template.html"),
    [string]$MesDesde = "",
    [string]$NetlifyToken = $env:NETLIFY_TOKEN,
    [string]$NetlifySiteId = "3e738d78-35b0-488b-b223-3ffc432f733c",
    [string]$NetlifySiteUrl = "https://cobertura-general-fusionlpelebes.netlify.app"
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
    Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $script:token" } -Method Get
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
$fechaHastaReal = if ($esMesActual) { $hoy.AddDays(1) } else { $inicioMes.AddMonths(1) }
$fechaDesde = $inicioMes.ToString("yyyy-MM-dd")
$fechaHasta = $fechaHastaReal.ToString("yyyy-MM-dd")
Write-Log "Revisando ventas del mes ($fechaDesde a $fechaHasta)..."

$pagestoskip = 0
$total = 0
while ($true) {
    if ($pagestoskip -gt 0 -and $pagestoskip % 10 -eq 0) { $script:token = Get-GescomToken }
    $page = Invoke-GescomApi -Path "/data/cmd/ventas/api/v2/get" -Query @{
        fechadesde = $fechaDesde; fechahasta = $fechaHasta; pagesize = 500; pagestoskip = $pagestoskip
    }
    if (-not $page -or $page.Count -eq 0) { break }
    $total += $page.Count
    foreach ($venta in $page) {
        if ($venta.estado -ne "Finalizada") { continue }
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
    periodoDesde = $fechaDesde
    periodoHasta = $(if ($esMesActual) { $hoy.ToString("yyyy-MM-dd") } else { $inicioMes.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd") })
    vendedores = $vendedoresOut
    proveedores = $proveedoresOut
    clientes = $clientesOut
}
$jsonText = $out | ConvertTo-Json -Depth 8 -Compress
[System.IO.File]::WriteAllText($OutPath, $jsonText, (New-Object System.Text.UTF8Encoding $false))
Write-Log "Guardado local: $OutPath"
$sinCompra = @($clientesOut | Where-Object { -not $_.compro }).Count
Write-Log "Clientes: $($clientesOut.Count) | sin compra (general): $sinCompra | vendedores: $($vendedoresOut.Count) | proveedores con venta: $($proveedoresOut.Count)"

if ($NetlifySiteId -and $NetlifySiteUrl) {
    Write-Log "Desplegando a Netlify..."
    $deployDir = Join-Path $env:TEMP "cobertura-general-netlify-deploy"
    if (Test-Path $deployDir) { Remove-Item $deployDir -Recurse -Force }
    New-Item -ItemType Directory -Path $deployDir | Out-Null
    Copy-Item $TemplatePath (Join-Path $deployDir "index.html")
    Copy-Item $OutPath (Join-Path $deployDir "data.json")
    $zipPath = Join-Path $deployDir "site.zip"
    Compress-Archive -Path (Join-Path $deployDir "index.html"), (Join-Path $deployDir "data.json") -DestinationPath $zipPath -Force
    $netlifyHeaders = @{ Authorization = "Bearer $NetlifyToken" }
    $deployResp = Invoke-RestMethod -Uri "https://api.netlify.com/api/v1/sites/$NetlifySiteId/deploys" -Headers $netlifyHeaders -Method Post -InFile $zipPath -ContentType "application/zip"
    Write-Log "Deploy Netlify: id=$($deployResp.id) state=$($deployResp.state)"
    Write-Log "Publicado en: $NetlifySiteUrl"
}
