<#
Clientes por ramo (categoria: almacen, autoservicio, kiosco, mayorista...) y por vendedor.
Mismo criterio de cartera que Cobertura General y Potenciales.

Por cada mes (desde que empieza la base compartida) y por cada vendedor muestra:
  - cuantos clientes tiene en cada ramo (su cartera),
  - cuantos de ellos compraron en el mes y que cobertura da,
  - cuanta venta neta generaron y que parte de la venta total del vendedor es cada ramo.

Reglas:
  - El ramo es clientes.subramo; el nombre sale de la tabla subramos. Los subramos "gemelos" que
    cargo Georgalos (georgalos-7 = Almacen, georgalos-1 = Autoservicio, etc.) se unifican con los
    de Gescom en $GRUPOS; para unificar otros, agregarlos ahi.
  - Cartera = clientes activos con vendedor asignado (clientes.vendedor). Se excluyen los vendedores
    1176, 43, 16 y 37 (no son vendedores reales).
  - Venta = venta neta VEN/DEB del cliente en el mes (las notas de credito no se restan, igual que
    en Potenciales). "Compro" = tuvo al menos un comprobante VEN/DEB en el mes.
  - La venta se atribuye al vendedor DUENO de la cartera del cliente (no a quien cargo el pedido).
  - RESPALDO (7/10/2026): si la base no trae vendedor en los clientes (una sincronizacion de
    catalogos los dejo en NULL), se usa el reparto de clientes publicado en docs/cobertura/data.json
    (ultimo bueno) y se avisa en el tablero.
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/ramos"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "ramos-clientes-template.html"),
    [string]$CarteraRespaldoPath = (Join-Path $PSScriptRoot "../docs/cobertura/data.json"),
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
function R2($x) { return [math]::Round([double]$x, 2) }

# ramos que son el mismo en Gescom y en Georgalos (o duplicados): id -> id del grupo
$GRUPOS = @{ "georgalos-7"="3"; "georgalos-1"="4"; "9"="4"; "10"="4"; "georgalos-4"="6"; "georgalos-32"="7"; "14"="8" }   # 9 y 10 = Autoservicio A y B, pedido de Gisela (7/10/2026)
# nombre para mostrar (los de Gescom vienen en MAYUSCULAS y sin acentos)
$aE = [string][char]0xE9; $aI = [string][char]0xED; $aO = [string][char]0xF3   # e, i, o con acento (el script se mantiene ASCII)
$NOMBRES_FIX = @{
    "ALMACEN"=("Almac" + $aE + "n"); "AUTOSERVICIO"="Autoservicio"; "KIOSCO"="Kiosco"; "MAYORISTA"="Mayorista"; "FIAMBRERIA"=("Fiambrer" + $aI + "a")
    "DIETETICA"=("Diet" + $aE + "tica"); "ESTACION DE SERVICIO"=("Estaci" + $aO + "n de servicio"); "VARIOS"="Varios"; "SIN SUB"="Sin ramo asignado"
    "Perfumeria"=("Perfumer" + $aI + "a"); "Articulo de limpieza"=("Art" + $aI + "culo de limpieza")
}

# --- 1) catalogo de ramos ---
$subr = Invoke-PanelSql "SELECT codigo, descripcion FROM subramos"
$ramoNombre = @{}
foreach ($s in $subr) {
    $id = [string]$s.codigo
    if ($GRUPOS.ContainsKey($id)) { continue }
    $d = ([string]$s.descripcion).Trim()
    if ($NOMBRES_FIX.ContainsKey($d)) { $d = $NOMBRES_FIX[$d] }
    $ramoNombre[$id] = $d
}
function Ramo-De([string]$subramo) {
    if (-not $subramo) { return "SIN" }
    $id = if ($GRUPOS.ContainsKey($subramo)) { $GRUPOS[$subramo] } else { $subramo }
    if ($ramoNombre.ContainsKey($id)) { return $id }
    return "SIN"
}
$ramoNombre["SIN"] = "Sin ramo"

# --- 2) clientes y cartera ---
Write-Log "Descargando clientes activos..."
$clientesRaw = Invoke-PanelSql "SELECT codigo, vendedor, subramo FROM clientes WHERE activo = 1"
$conVend = @($clientesRaw | Where-Object { $_.vendedor }).Count
$subramoDe = @{}
foreach ($c in $clientesRaw) { $subramoDe[[string]$c.codigo] = [string]$c.subramo }
$vendedorDe = @{}
$carteraFuente = "base"
if ($conVend -ge 1000) {
    foreach ($c in $clientesRaw) { if ($c.vendedor) { $vendedorDe[[string]$c.codigo] = [string]$c.vendedor } }
} else {
    Write-Log "AVISO: la base trae solo $conVend clientes con vendedor (falla de sincronizacion de catalogos). Se usa el reparto del respaldo: $CarteraRespaldoPath"
    if (-not (Test-Path $CarteraRespaldoPath)) { throw "No hay vendedores de cliente en la base y tampoco existe el respaldo $CarteraRespaldoPath" }
    $resp = Get-Content $CarteraRespaldoPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($c in $resp.clientes) { $vendedorDe[[string]$c.codigo] = [string]$c.codigoVendedor }
    $carteraFuente = "respaldo del " + ([datetime]$resp.generatedAt).ToString("dd/MM/yyyy")
}
$cartera = @{}   # codigo -> @{ vendedor; ramo }
foreach ($cod in $vendedorDe.Keys) {
    $v = $vendedorDe[$cod]
    if ($EXCLUDED -contains $v) { continue }
    if (-not $subramoDe.ContainsKey($cod)) { continue }   # cliente que ya no esta activo
    $cartera[$cod] = @{ vendedor = $v; ramo = (Ramo-De $subramoDe[$cod]) }
}
Write-Log "Cartera: $($cartera.Count) clientes con vendedor real | fuente: $carteraFuente"

$vend = @{}
foreach ($v in (Invoke-PanelSql "SELECT codigo, nombre FROM vendedores")) { $vend[[string]$v.codigo] = $v.nombre }

# --- 3) meses a procesar ---
$base = (Invoke-PanelSql "SELECT MIN(fecha) AS desde FROM ventas")[0].desde
$baseDesde = [datetime]::ParseExact(([string]$base).Substring(0,10), "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
$hoy = (Get-Date).Date
$primerMes = if ($MesDesde) { [datetime]::ParseExact($MesDesde + "-01", "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture) } else { [datetime]::new($baseDesde.Year, $baseDesde.Month, 1) }
$mesActualKey = $hoy.ToString("yyyy-MM")
$meses = [ordered]@{}
for ($m = $primerMes; $m -le $hoy; $m = $m.AddMonths(1)) {
    $mk = $m.ToString("yyyy-MM")
    $ini = $m; $sig = $m.AddMonths(1)
    $desdeReal = if ($ini -lt $baseDesde) { $baseDesde } else { $ini }
    $hastaReal = if ($sig.AddDays(-1) -gt $hoy) { $hoy } else { $sig.AddDays(-1) }
    Write-Log "Mes ${mk}: descargando venta por cliente..."
    $filas = Invoke-PanelSql "SELECT cliente, COUNT(*) AS comprobantes, SUM(neto) AS venta FROM ventas WHERE fecha >= '$($ini.ToString('yyyy-MM-dd'))' AND fecha < '$($sig.ToString('yyyy-MM-dd'))' AND tipo IN ('VEN','DEB') GROUP BY cliente"
    $ventaCli = @{}
    $fuera = 0.0
    foreach ($f in $filas) {
        $cod = [string]$f.cliente
        if ($cartera.ContainsKey($cod)) { $ventaCli[$cod] = @{ compro = ([int]$f.comprobantes -gt 0); venta = [double]$f.venta } } else { $fuera += [double]$f.venta }
    }
    $porRamo = @{}; $porVend = @{}
    foreach ($cod in $cartera.Keys) {
        $c = $cartera[$cod]; $v = $c.vendedor; $r = $c.ramo
        $compro = $false; $venta = 0.0
        if ($ventaCli.ContainsKey($cod)) { $compro = $ventaCli[$cod].compro; $venta = $ventaCli[$cod].venta }
        if (-not $porRamo.ContainsKey($r)) { $porRamo[$r] = @{ clientes = 0; compraron = 0; venta = 0.0 } }
        $porRamo[$r].clientes++; if ($compro) { $porRamo[$r].compraron++ }; $porRamo[$r].venta += $venta
        if (-not $porVend.ContainsKey($v)) { $porVend[$v] = @{ clientes = 0; compraron = 0; venta = 0.0; ramos = @{} } }
        $pv = $porVend[$v]; $pv.clientes++; if ($compro) { $pv.compraron++ }; $pv.venta += $venta
        if (-not $pv.ramos.ContainsKey($r)) { $pv.ramos[$r] = @{ clientes = 0; compraron = 0; venta = 0.0 } }
        $pv.ramos[$r].clientes++; if ($compro) { $pv.ramos[$r].compraron++ }; $pv.ramos[$r].venta += $venta
    }
    $ramosOut = @($porRamo.Keys | ForEach-Object { [ordered]@{ id = $_; nombre = $ramoNombre[$_]; clientes = $porRamo[$_].clientes; compraron = $porRamo[$_].compraron; venta = (R2 $porRamo[$_].venta) } } | Sort-Object { -$_.venta })
    $vendOut = @($porVend.Keys | ForEach-Object {
        $pv = $porVend[$_]
        [ordered]@{
            codigo = $_; nombre = $(if ($vend.ContainsKey($_)) { $vend[$_] } else { "Vendedor $_" })
            clientes = $pv.clientes; compraron = $pv.compraron; venta = (R2 $pv.venta)
            ramos = @($pv.ramos.Keys | ForEach-Object { [ordered]@{ id = $_; nombre = $ramoNombre[$_]; clientes = $pv.ramos[$_].clientes; compraron = $pv.ramos[$_].compraron; venta = (R2 $pv.ramos[$_].venta) } } | Sort-Object { -$_.venta })
        } } | Sort-Object { -$_.venta })
    $totClientes = 0; $totComp = 0; $totVenta = 0.0
    foreach ($ro in $ramosOut) { $totClientes += $ro.clientes; $totComp += $ro.compraron; $totVenta += $ro.venta }
    $meses[$mk] = [ordered]@{
        cerrado = ($mk -ne $mesActualKey)
        periodoDesde = $desdeReal.ToString("yyyy-MM-dd"); periodoHasta = $hastaReal.ToString("yyyy-MM-dd")
        totales = [ordered]@{ clientes = $totClientes; compraron = $totComp; venta = (R2 $totVenta); ventaFueraDeCartera = (R2 $fuera) }
        ramos = $ramosOut; vendedores = $vendOut
    }
    Write-Log "  $mk -> clientes $totClientes | compraron $totComp | venta $(R2 $totVenta) | fuera de cartera $(R2 $fuera)"
}

# --- 4) guardar ---
$ids = @($cartera.Values | ForEach-Object { $_.ramo } | Group-Object | Sort-Object Count -Descending | ForEach-Object { $_.Name })
$out = [ordered]@{
    generatedAt = (Get-Date).ToString("o")
    carteraFuente = $carteraFuente
    mesActual = $mesActualKey
    ramos = @($ids | ForEach-Object { [ordered]@{ id = $_; nombre = $ramoNombre[$_] } })
    meses = $meses
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
[System.IO.File]::WriteAllText($OutPath, (ConvertTo-Json -InputObject $out -Depth 10 -Compress), (New-Object System.Text.UTF8Encoding $false))
if (Test-Path $TemplatePath) { Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force }
Write-Log "Guardado: $OutPath (meses: $($meses.Keys -join ', '))"
