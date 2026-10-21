<#
Seguimiento del Combo Cobertura Softys (combo 1137 en Gescom), lanzado el 19/10/2026.

Objetivo: subir 9 puntos la cobertura de Softys (~300 clientes nuevos sobre la cartera).

Clientes objetivo (foto fija al dia del lanzamiento, de la cartera de vendedores reales):
  - "Nunca Softys": compraron algo en los 90 dias anteriores al lanzamiento pero ninguna vez Softys.
  - "Recuperar": compraron Softys en los 90 dias anteriores, pero no en el mes del lanzamiento.
  - "Sin compras": no compraron nada en los 90 dias anteriores (clientes a reactivar).
Recuperado = cliente objetivo con al menos una factura de Softys (proveedor 150) desde el lanzamiento,
con combo o sin combo.

Combos vendidos: venta_combos con combo 1137 (VEN/DEB desde el lanzamiento), por el vendedor que
cargo el pedido. Si la base todavia no trae ninguna fila del combo, se buscan las ventas que tienen
los dos articulos del combo en la proporcion del combo (1 Balanced + 10 TF Normal) y se avisa en el
tablero que el conteo esta "detectado por articulos".

Regla del combo (Gisela, 21/10/2026): se puede vender a los tres grupos (no compradores de Softys,
clientes sin compras y los que compraron Softys antes pero no en el mes del lanzamiento). "Fuera del
objetivo" = combo vendido a un cliente que ya habia comprado Softys en el mes del lanzamiento antes de lanzar.

Reglas: cartera = clientes activos con vendedor asignado; se excluyen los vendedores 1176, 43, 16 y 37.
El objetivo de cada vendedor es los 300 clientes repartidos segun cuantos clientes "Nunca Softys" y
"Recuperar" tiene (los "Sin compras" suman al avance pero no a la meta).
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/combo-softys"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "combo-softys-template.html"),
    [string]$Combo = "1137",
    [string]$Inicio = "2026-10-19",
    [string]$Proveedor = "150",
    [int]$ObjetivoClientes = 300,
    [string[]]$ArticulosCombo = @("1501627", "1509764"),
    [int]$UnidadesSegundoArticulo = 10
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$EXCLUDED = @("43", "1176", "16", "37")
$EXCL_SQL = "('43','1176','16','37')"
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
function Invoke-PanelSql([string]$Sql) {
    $uri = "$($config.baseUrl)/consulta?sql=" + [uri]::EscapeDataString($Sql)
    $headers = @{ Authorization = "Bearer $($config.clave)" }
    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    if ($r.truncado) { Write-Log "AVISO: la consulta vino truncada (mas de 20000 filas) -- revisar y acotar." }
    return $r.filas
}
function R2($x) { return [math]::Round([double]$x, 2) }
$inv = [Globalization.CultureInfo]::InvariantCulture
$ini = [datetime]::ParseExact($Inicio, "yyyy-MM-dd", $inv)
$desde90 = $ini.AddDays(-90).ToString("yyyy-MM-dd")
$mesIni = $ini.ToString("yyyy-MM-01")
$mesAnt = $ini.AddMonths(-1).ToString("yyyy-MM-01")
$SOFTYS = "SELECT codigo FROM articulos WHERE proveedor='$Proveedor'"

# --- 1) cartera ---
$vendNom = @{}; foreach ($v in (Invoke-PanelSql "SELECT codigo, nombre FROM vendedores")) { $vendNom[[string]$v.codigo] = ([string]$v.nombre).Trim() }
$cartera = @{}
foreach ($c in (Invoke-PanelSql "SELECT codigo, razon_social, nombre, localidad, vendedor FROM clientes WHERE activo=1 AND vendedor IS NOT NULL AND vendedor NOT IN $EXCL_SQL")) {
    $cartera[[string]$c.codigo] = $c
}
if ($cartera.Count -lt 1000) { throw "La base trae solo $($cartera.Count) clientes con vendedor: no se pisa el tablero (ver incidente del 7/10)." }
Write-Log "Cartera: $($cartera.Count) clientes con vendedor real"

# --- 2) clientes objetivo (foto al lanzamiento) ---
$compro90 = @{}; foreach ($x in (Invoke-PanelSql "SELECT DISTINCT cliente FROM ventas WHERE tipo IN ('VEN','DEB') AND vendedor NOT IN $EXCL_SQL AND fecha >= '$desde90' AND fecha < '$Inicio'")) { $compro90[[string]$x.cliente] = $true }
$ultSoftys = @{}
foreach ($x in (Invoke-PanelSql "SELECT v.cliente, MAX(v.fecha) f FROM ventas v JOIN venta_items i ON i.venta_id = v.id WHERE v.tipo IN ('VEN','DEB') AND v.fecha >= '$desde90' AND v.fecha < '$Inicio' AND i.articulo IN ($SOFTYS) GROUP BY v.cliente")) { $ultSoftys[[string]$x.cliente] = ([string]$x.f).Substring(0, 10) }
$objetivo = @{}
foreach ($k in $cartera.Keys) {
    if (-not $compro90.ContainsKey($k)) { $objetivo[$k] = "sinCompras"; continue }
    if (-not $ultSoftys.ContainsKey($k)) { $objetivo[$k] = "nunca" }
    elseif ($ultSoftys[$k] -lt $mesIni) { $objetivo[$k] = "recuperar" }
}
Write-Log ("Clientes objetivo: {0} (nunca Softys {1}, recuperar {2}, sin compras {3})" -f $objetivo.Count, @($objetivo.Values | Where-Object { $_ -eq 'nunca' }).Count, @($objetivo.Values | Where-Object { $_ -eq 'recuperar' }).Count, @($objetivo.Values | Where-Object { $_ -eq 'sinCompras' }).Count)

# --- 3) compras de Softys desde el lanzamiento ---
$softysDesde = @{}
foreach ($x in (Invoke-PanelSql "SELECT v.cliente, MIN(v.fecha) f, SUM(i.neto) neto FROM ventas v JOIN venta_items i ON i.venta_id = v.id WHERE v.tipo IN ('VEN','DEB') AND v.fecha >= '$Inicio' AND i.articulo IN ($SOFTYS) GROUP BY v.cliente")) {
    $softysDesde[[string]$x.cliente] = @{ fecha = ([string]$x.f).Substring(0, 10); neto = [double]$x.neto }
}

# --- 4) combos vendidos ---
$fuenteCombos = "venta_combos"
$filasCombo = @(Invoke-PanelSql "SELECT v.id, substr(v.fecha,1,10) fecha, v.cliente, v.vendedor, SUM(c.cantidad) q FROM venta_combos c JOIN ventas v ON v.id = c.venta_id WHERE c.combo = '$Combo' AND v.tipo IN ('VEN','DEB') AND v.fecha >= '$Inicio' GROUP BY v.id")
if ($filasCombo.Count -eq 0) {
    $a1 = $ArticulosCombo[0]; $a2 = $ArticulosCombo[1]
    $filasCombo = @(Invoke-PanelSql @"
SELECT v.id, substr(v.fecha,1,10) fecha, v.cliente, v.vendedor,
       SUM(CASE WHEN i.articulo='$a1' THEN i.cantidad ELSE 0 END) q1, SUM(CASE WHEN i.articulo='$a2' THEN i.cantidad ELSE 0 END) q2
FROM ventas v JOIN venta_items i ON i.venta_id = v.id
WHERE v.tipo IN ('VEN','DEB') AND v.fecha >= '$Inicio' AND i.articulo IN ('$a1','$a2')
GROUP BY v.id HAVING q1 > 0 AND q2 = q1 * $UnidadesSegundoArticulo
"@) | ForEach-Object { [pscustomobject]@{ id = $_.id; fecha = $_.fecha; cliente = $_.cliente; vendedor = $_.vendedor; q = $_.q1 } }
    $fuenteCombos = if ($filasCombo.Count) { "articulos" } else { "sin ventas" }
}
Write-Log "Combos: $(@($filasCombo).Count) ventas (fuente: $fuenteCombos)"
$comboCli = @{}
foreach ($f in $filasCombo) { $k = [string]$f.cliente; if (-not $comboCli[$k]) { $comboCli[$k] = 0 }; $comboCli[$k] += [double]$f.q }

# --- 5) por vendedor ---
$totalObj = [math]::Max(1, @($objetivo.Values | Where-Object { $_ -ne 'sinCompras' }).Count)
# fuera del objetivo: el cliente ya habia comprado Softys en el mes del lanzamiento antes de lanzar
function Fuera-Objetivo([string]$cli) { return ($ultSoftys.ContainsKey($cli) -and $ultSoftys[$cli] -ge $mesIni) }
$vend = @{}
function VendRow($cod) {
    if (-not $vend.ContainsKey($cod)) {
        $vend[$cod] = [ordered]@{ codigo = $cod; nombre = $(if ($vendNom[$cod]) { $vendNom[$cod] } else { "Vendedor $cod" }); cartera = 0; objetivoNunca = 0; objetivoRecuperar = 0; objetivoSinCompras = 0
            recuperadosNunca = 0; recuperadosRecuperar = 0; recuperadosSinCompras = 0; recuperadosConCombo = 0; combosVendidos = 0; ventasConCombo = 0; clientesConCombo = 0; combosFueraObjetivo = 0; meta = 0 }
    }
    return $vend[$cod]
}
foreach ($k in $cartera.Keys) { $r = VendRow ([string]$cartera[$k].vendedor); $r.cartera++ }
foreach ($k in $objetivo.Keys) {
    $r = VendRow ([string]$cartera[$k].vendedor)
    switch ($objetivo[$k]) { 'nunca' { $r.objetivoNunca++ } 'recuperar' { $r.objetivoRecuperar++ } default { $r.objetivoSinCompras++ } }
    if ($softysDesde.ContainsKey($k)) {
        switch ($objetivo[$k]) { 'nunca' { $r.recuperadosNunca++ } 'recuperar' { $r.recuperadosRecuperar++ } default { $r.recuperadosSinCompras++ } }
        if ($comboCli.ContainsKey($k)) { $r.recuperadosConCombo++ }
    }
}
$cliPorVend = @{}
foreach ($f in $filasCombo) {
    $cod = [string]$f.vendedor; if ($EXCLUDED -contains $cod) { continue }
    $r = VendRow $cod; $r.combosVendidos += [double]$f.q; $r.ventasConCombo++
    $k = [string]$f.cliente
    if (-not $cliPorVend[$cod]) { $cliPorVend[$cod] = @{} }; $cliPorVend[$cod][$k] = $true
    if (Fuera-Objetivo $k) { $r.combosFueraObjetivo += [double]$f.q }
}
foreach ($cod in $cliPorVend.Keys) { $vend[$cod].clientesConCombo = $cliPorVend[$cod].Count }
foreach ($r in $vend.Values) { $r.meta = [math]::Round($ObjetivoClientes * ($r.objetivoNunca + $r.objetivoRecuperar) / $totalObj) }
$vendOut = @($vend.Values | Where-Object { $_.objetivoNunca + $_.objetivoRecuperar + $_.objetivoSinCompras + $_.combosVendidos -gt 0 } | Sort-Object { -($_.objetivoNunca + $_.objetivoRecuperar) })

# --- 6) clientes objetivo (lista para trabajar) ---
$clientes = @($objetivo.Keys | ForEach-Object {
    $c = $cartera[$_]; $s = $softysDesde[$_]
    [ordered]@{ codigo = $_; nombre = $(if ($c.razon_social) { ([string]$c.razon_social).Trim() } else { ([string]$c.nombre).Trim() }); localidad = ([string]$c.localidad).Trim()
        vendedor = [string]$c.vendedor; grupo = $objetivo[$_]; ultimaSoftys = $ultSoftys[$_]
        recuperado = [bool]$s; fechaRecupero = $(if ($s) { $s.fecha } else { $null }); ventaSoftys = $(if ($s) { R2 $s.neto } else { 0 })
        combos = $(if ($comboCli.ContainsKey($_)) { $comboCli[$_] } else { 0 }) }
} | Sort-Object { $_.vendedor }, { -[int]$_.recuperado }, { $_.nombre })

# --- 7) por dia ---
$hoy = (Get-Date).Date
$dias = @()
for ($d = $ini; $d -le $hoy; $d = $d.AddDays(1)) {
    if ($d.DayOfWeek -eq 'Sunday') { continue }
    $ds = $d.ToString("yyyy-MM-dd")
    $dias += [ordered]@{ fecha = $ds
        recuperados = @($objetivo.Keys | Where-Object { $softysDesde.ContainsKey($_) -and $softysDesde[$_].fecha -eq $ds }).Count
        combos = (@($filasCombo | Where-Object { $_.fecha -eq $ds }) | Measure-Object q -Sum).Sum }
}
foreach ($d in $dias) { if ($null -eq $d.combos) { $d.combos = 0 } }

# --- 8) cobertura Softys del mes en curso ---
$mesHoy = $hoy.ToString("yyyy-MM-01")
$compSoftysMes = @{}; foreach ($x in (Invoke-PanelSql "SELECT DISTINCT v.cliente FROM ventas v JOIN venta_items i ON i.venta_id = v.id WHERE v.tipo IN ('VEN','DEB') AND v.fecha >= '$mesHoy' AND i.articulo IN ($SOFTYS)")) { $compSoftysMes[[string]$x.cliente] = $true }
$cobMes = @($cartera.Keys | Where-Object { $compSoftysMes.ContainsKey($_) }).Count
$compLanz = @{}; foreach ($x in (Invoke-PanelSql "SELECT DISTINCT v.cliente FROM ventas v JOIN venta_items i ON i.venta_id = v.id WHERE v.tipo IN ('VEN','DEB') AND v.fecha >= '$mesIni' AND v.fecha < '$Inicio' AND i.articulo IN ($SOFTYS)")) { $compLanz[[string]$x.cliente] = $true }
$cobLanz = @($cartera.Keys | Where-Object { $compLanz.ContainsKey($_) }).Count

$recup = @($objetivo.Keys | Where-Object { $softysDesde.ContainsKey($_) })
$out = [ordered]@{
    generatedAt = (Get-Date).ToString("o")
    combo = $Combo; inicio = $Inicio; fuenteCombos = $fuenteCombos
    objetivoClientes = $ObjetivoClientes; objetivoPuntos = 9
    cartera = $cartera.Count
    coberturaLanzamiento = [ordered]@{ mes = $mesIni.Substring(0, 7); compradores = $cobLanz; pct = (R2 ($cobLanz / $cartera.Count)) }
    coberturaMes = [ordered]@{ mes = $mesHoy.Substring(0, 7); compradores = $cobMes; pct = (R2 ($cobMes / $cartera.Count)) }
    totales = [ordered]@{
        objetivo = $objetivo.Count
        objetivoNunca = @($objetivo.Values | Where-Object { $_ -eq 'nunca' }).Count
        objetivoRecuperar = @($objetivo.Values | Where-Object { $_ -eq 'recuperar' }).Count
        objetivoSinCompras = @($objetivo.Values | Where-Object { $_ -eq 'sinCompras' }).Count
        recuperados = $recup.Count
        recuperadosNunca = @($recup | Where-Object { $objetivo[$_] -eq 'nunca' }).Count
        recuperadosRecuperar = @($recup | Where-Object { $objetivo[$_] -eq 'recuperar' }).Count
        recuperadosSinCompras = @($recup | Where-Object { $objetivo[$_] -eq 'sinCompras' }).Count
        recuperadosConCombo = @($recup | Where-Object { $comboCli.ContainsKey($_) }).Count
        combosVendidos = [double](($filasCombo | Measure-Object q -Sum).Sum)
        ventasConCombo = @($filasCombo).Count
        clientesConCombo = @($comboCli.Keys).Count
        combosFueraObjetivo = [double](($filasCombo | Where-Object { Fuera-Objetivo ([string]$_.cliente) } | Measure-Object q -Sum).Sum)
    }
    combosFuera = @($filasCombo | Where-Object { Fuera-Objetivo ([string]$_.cliente) } | ForEach-Object {
        $c = $cartera[[string]$_.cliente]
        [ordered]@{ fecha = $_.fecha; vendedor = [string]$_.vendedor; cliente = [string]$_.cliente
            nombre = $(if ($c) { ([string]$c.razon_social).Trim() } else { "" }); localidad = $(if ($c) { ([string]$c.localidad).Trim() } else { "" })
            ultimaSoftys = $ultSoftys[[string]$_.cliente]; combos = [double]$_.q }
    })
    vendedores = $vendOut
    dias = $dias
    clientes = $clientes
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
[System.IO.File]::WriteAllText($OutPath, (ConvertTo-Json -InputObject $out -Depth 8 -Compress), (New-Object System.Text.UTF8Encoding $false))
if (Test-Path $TemplatePath) { Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force }
Write-Log ("Guardado: {0} | recuperados {1} de {2} (meta {3}) | combos {4} | cobertura Softys {5}: {6}/{7}" -f $OutPath, $recup.Count, $objetivo.Count, $ObjetivoClientes, $out.totales.combosVendidos, $out.coberturaMes.mes, $cobMes, $cartera.Count)
