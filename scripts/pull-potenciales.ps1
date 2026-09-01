<#
Mismo dashboard que Cobertura General, pero el "universo" de cada vendedor se
recorta a sus clientes "potenciales": los que concentran el 80% de lo que ese
vendedor vendio en los ultimos 3 meses (curva de Pareto por vendedor).

Reglas:
  - Se excluyen los vendedores 1176, 43, 16, 37 (no son vendedores reales).
  - Ranking (quien es "potencial"): venta neta por cliente en los ultimos 3
    meses (ventana movil hasta hoy), ordenado de mayor a menor, se toman los
    clientes hasta llegar al 80% acumulado de la venta de ese vendedor.
  - "Compro" (para el indicador de cobertura) = venta finalizada, no nota de
    credito, en el MES EN CURSO -- igual criterio que Cobertura General.
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string]$OutPath = (Join-Path $PSScriptRoot "potenciales-data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "potenciales-template.html"),
    [double]$CorteAcumulado = 0.8,
    [string]$NetlifyToken = $env:NETLIFY_TOKEN,
    [string]$NetlifySiteId = "b3801462-4190-4f1f-88c0-aecaecb57b21",
    [string]$NetlifySiteUrl = "https://potenciales-fusionlpelebes.netlify.app"
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

# --- cliente -> {vendedor, dias} a partir de la primera ruta de preventa ---
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
        ventaNeta3m = 0.0; compro = $false; proveedoresComprados = (New-Object System.Collections.Generic.HashSet[string])
    }
    [void]$vendedoresConCartera.Add($vcod)
}
Write-Log "Clientes con cartera asignada (vendedor real): $($clienteInfo.Count) | vendedores distintos: $($vendedoresConCartera.Count)"

# --- ventana de 3 meses moviles (ranking) + mes en curso (indicador de compra), en una sola pasada ---
$hoy = (Get-Date).Date
$rankingDesde = $hoy.AddMonths(-3)
$inicioMesActual = Get-Date -Year $hoy.Year -Month $hoy.Month -Day 1
$fechaDesde = $rankingDesde.ToString("yyyy-MM-dd")
$fechaHasta = $hoy.AddDays(1).ToString("yyyy-MM-dd")
Write-Log "Revisando ventas ($fechaDesde a $fechaHasta) -- ranking 3 meses + mes en curso desde $($inicioMesActual.ToString('yyyy-MM-dd'))..."

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
        $ci.ventaNeta3m += [double]$venta.importeNeto
        $fechaPedido = [datetime]$venta.fechaPedido
        if ($fechaPedido -ge $inicioMesActual) {
            $ci.compro = $true
            foreach ($item in $venta.items) {
                $prov = $artProv[[string]$item.codigoItem]
                if ($prov) { [void]$ci.proveedoresComprados.Add($prov) }
            }
        }
    }
    if ($page.Count -lt 500) { break }
    $pagestoskip++
}
Write-Log "Ventas revisadas: $total"

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

# solo vendedores que tengan al menos 1 cliente potencial (evita entradas vacias en el selector)
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
$jsonText = $out | ConvertTo-Json -Depth 8 -Compress
[System.IO.File]::WriteAllText($OutPath, $jsonText, (New-Object System.Text.UTF8Encoding $false))
Write-Log "Guardado local: $OutPath"
$sinCompra = @($clientesOut | Where-Object { -not $_.compro }).Count
Write-Log "Potenciales: $($clientesOut.Count) | sin compra este mes: $sinCompra | vendedores: $($vendedoresOut.Count) | proveedores con venta: $($proveedoresOut.Count)"

if ($NetlifySiteId -and $NetlifySiteUrl) {
    Write-Log "Desplegando a Netlify..."
    $deployDir = Join-Path $env:TEMP "potenciales-netlify-deploy"
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
