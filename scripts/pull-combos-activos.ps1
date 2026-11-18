<#
Seguimiento de TODOS los combos activos de Gescom (reemplaza al tablero del combo 1137, que terminó).

Combo activo = el que figura hoy en la tabla `combos` de la base (Gescom la sincroniza: cuando un combo
se da de baja, desaparece). Para cada uno:
  - lanzamiento = primera venta con ese combo (VEN/DEB) en los últimos 150 días
  - combos vendidos, facturas, clientes distintos, venta neta y margen de los artículos del combo
    (dentro de cada factura se toma solo la cantidad que corresponde al combo, el resto no se cuenta)
  - clientes nuevos en el producto = no habían comprado ninguno de los artículos del combo en los
    90 días previos al lanzamiento
  - por vendedor, por día y la lista de clientes que lo compraron
Reglas: se excluyen los vendedores 1176, 43, 16 y 37.
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/combos"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "combos-template.html"),
    [int]$DiasAtras = 150
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$EXCL_SQL = "('43','1176','16','37')"
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
function Invoke-PanelSql([string]$Sql) {
    $uri = "$($config.baseUrl)/consulta?sql=" + [uri]::EscapeDataString($Sql)
    $r = Invoke-RestMethod -Uri $uri -Headers @{ Authorization = "Bearer $($config.clave)" } -Method Get
    if ($r.truncado) { throw "La consulta vino truncada (mas de 20000 filas): acotar." }
    return @($r.filas)
}
function R2($x) { return [math]::Round([double]$x, 2) }
function Q([string[]]$xs) { return "(" + (($xs | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }) -join ",") + ")" }
$hoy = (Get-Date).Date
$desde = $hoy.AddDays(-$DiasAtras).ToString("yyyy-MM-dd")

# --- 1) definicion de los combos activos ---
$defs = Invoke-PanelSql "SELECT codigo, descripcion, actualizado FROM combos ORDER BY codigo"
if ($defs.Count -eq 0) { throw "La tabla combos vino vacia: no se pisa el tablero." }
$comps = Invoke-PanelSql "SELECT combo, componente, nombre, cantidad, descuento, articulos FROM combo_componentes ORDER BY combo, componente"
$vendNom = @{}; foreach ($v in (Invoke-PanelSql "SELECT codigo, nombre FROM vendedores")) { $vendNom[[string]$v.codigo] = ([string]$v.nombre).Trim() }
$provNom = @{}; foreach ($p in (Invoke-PanelSql "SELECT codigo, nombre FROM proveedores")) { $provNom[[string]$p.codigo] = ([string]$p.nombre).Trim() }

# articulos de cada combo: los de sus componentes + los de componentes con el mismo nombre en otros combos
# (Gescom a veces cambia el codigo del componente y las facturas viejas traen el anterior)
$artPorNombre = @{}
foreach ($c in $comps) { foreach ($a in ([string]$c.articulos -split ',')) { $a = $a.Trim(); if (-not $a) { continue }; $n = ([string]$c.nombre).Trim().ToUpper(); if (-not $artPorNombre[$n]) { $artPorNombre[$n] = @{} }; $artPorNombre[$n][$a] = $true } }
$combos = [ordered]@{}
foreach ($d in $defs) {
    $cod = [string]$d.codigo
    $mis = @($comps | Where-Object { [string]$_.combo -eq $cod })
    $porArt = @{}   # articulo -> unidades por combo (suma de componentes que lo usan)
    $alias = @{}    # articulo alternativo -> articulo principal
    foreach ($c in $mis) {
        $arts = @(([string]$c.articulos -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if (-not $arts.Count) { continue }
        $p = $arts[0]
        if (-not $porArt.ContainsKey($p)) { $porArt[$p] = 0 }
        $porArt[$p] += [double]$c.cantidad
        foreach ($a in $arts) { $alias[$a] = $p }
        foreach ($a in $artPorNombre[([string]$c.nombre).Trim().ToUpper()].Keys) { if (-not $alias.ContainsKey($a)) { $alias[$a] = $p } }
    }
    $combos[$cod] = [ordered]@{ codigo = $cod; descripcion = ([string]$d.descripcion).Trim(); comps = $mis; porArt = $porArt; alias = $alias }
}
$todosArts = @($combos.Values | ForEach-Object { $_.alias.Keys } | Sort-Object -Unique)
$artInfo = @{}
foreach ($a in (Invoke-PanelSql "SELECT codigo, descripcion, proveedor FROM articulos WHERE codigo IN $(Q $todosArts)")) { $artInfo[[string]$a.codigo] = $a }

# --- 2) ventas con combo ---
$codsSql = Q @($combos.Keys)
$filas = Invoke-PanelSql @"
SELECT v.id, substr(v.fecha,1,10) fecha, v.cliente, v.vendedor, c.combo, SUM(c.cantidad) q
FROM venta_combos c JOIN ventas v ON v.id = c.venta_id
WHERE c.combo IN $codsSql AND v.tipo IN ('VEN','DEB') AND v.fecha >= '$desde' AND v.vendedor NOT IN $EXCL_SQL
GROUP BY v.id, c.combo
"@
Write-Log "Ventas con combo: $($filas.Count) (desde $desde)"
$items = @{}
$ids = @($filas | ForEach-Object { [string]$_.id } | Sort-Object -Unique)
for ($i = 0; $i -lt $ids.Count; $i += 400) {
    $lote = $ids[$i..([math]::Min($i + 399, $ids.Count - 1))]
    foreach ($it in (Invoke-PanelSql "SELECT venta_id, articulo, SUM(cantidad) q, SUM(neto) neto, SUM(precio_costo*cantidad) costo FROM venta_items WHERE venta_id IN ($($lote -join ',')) AND articulo IN $(Q $todosArts) GROUP BY venta_id, articulo")) {
        $k = [string]$it.venta_id; if (-not $items[$k]) { $items[$k] = @() }; $items[$k] += $it
    }
}
$cliNom = @{}
$clis = @($filas | ForEach-Object { [string]$_.cliente } | Sort-Object -Unique)
for ($i = 0; $i -lt $clis.Count; $i += 400) {
    $lote = $clis[$i..([math]::Min($i + 399, $clis.Count - 1))]
    foreach ($c in (Invoke-PanelSql "SELECT codigo, razon_social, nombre, localidad, vendedor FROM clientes WHERE codigo IN $(Q $lote)")) { $cliNom[[string]$c.codigo] = $c }
}

# --- 3) valor de cada venta: solo la cantidad del combo de cada articulo ---
$ventas = @()
foreach ($f in $filas) {
    $cb = $combos[[string]$f.combo]; $q = [double]$f.q
    $neto = 0.0; $costo = 0.0
    $grp = @{}
    foreach ($it in @($items[[string]$f.id])) {
        if (-not $it) { continue }
        $p = $cb.alias[[string]$it.articulo]; if (-not $p) { continue }
        if (-not $grp[$p]) { $grp[$p] = @{ q = 0.0; neto = 0.0; costo = 0.0 } }
        $grp[$p].q += [double]$it.q; $grp[$p].neto += [double]$it.neto; $grp[$p].costo += [double]$it.costo
    }
    foreach ($p in $grp.Keys) {
        $g = $grp[$p]; if ($g.q -le 0) { continue }
        $tope = $q * [double]$cb.porArt[$p]
        $fr = [math]::Min(1.0, $tope / $g.q)
        $neto += $g.neto * $fr; $costo += $g.costo * $fr
    }
    $ventas += [pscustomobject]@{ id = [string]$f.id; fecha = [string]$f.fecha; cliente = [string]$f.cliente; vendedor = [string]$f.vendedor; combo = [string]$f.combo; q = $q; neto = $neto; costo = $costo }
}

# --- 4) por combo ---
$out = @()
$diasSet = @{}
foreach ($cb in $combos.Values) {
    $vs = @($ventas | Where-Object { $_.combo -eq $cb.codigo })
    $lanz = if ($vs.Count) { ($vs | Sort-Object fecha | Select-Object -First 1).fecha } else { $null }
    $ult = if ($vs.Count) { ($vs | Sort-Object fecha -Descending | Select-Object -First 1).fecha } else { $null }
    # clientes que ya compraban los articulos del combo en los 90 dias previos al lanzamiento
    $prev = @{}
    if ($lanz) {
        $l = [datetime]::ParseExact($lanz, "yyyy-MM-dd", $null)
        $sql = "SELECT DISTINCT v.cliente FROM ventas v JOIN venta_items i ON i.venta_id = v.id WHERE v.tipo IN ('VEN','DEB') AND v.fecha >= '$($l.AddDays(-90).ToString('yyyy-MM-dd'))' AND v.fecha < '$lanz' AND i.articulo IN $(Q @($cb.alias.Keys))"
        foreach ($x in (Invoke-PanelSql $sql)) { $prev[[string]$x.cliente] = $true }
    }
    $porCli = @{}
    foreach ($v in $vs) {
        if (-not $porCli[$v.cliente]) { $porCli[$v.cliente] = [ordered]@{ q = 0.0; neto = 0.0; facturas = 0; primera = $v.fecha; ultima = $v.fecha; vendedor = $v.vendedor } }
        $pc = $porCli[$v.cliente]; $pc.q += $v.q; $pc.neto += $v.neto; $pc.facturas++
        if ($v.fecha -lt $pc.primera) { $pc.primera = $v.fecha }; if ($v.fecha -gt $pc.ultima) { $pc.ultima = $v.fecha; $pc.vendedor = $v.vendedor }
    }
    $porVend = @{}
    foreach ($v in $vs) {
        if (-not $porVend[$v.vendedor]) { $porVend[$v.vendedor] = [ordered]@{ codigo = $v.vendedor; nombre = $(if ($vendNom[$v.vendedor]) { $vendNom[$v.vendedor] } else { "Vendedor $($v.vendedor)" }); q = 0.0; facturas = 0; neto = 0.0; costo = 0.0; clientes = @{}; nuevos = @{} } }
        $pv = $porVend[$v.vendedor]; $pv.q += $v.q; $pv.facturas++; $pv.neto += $v.neto; $pv.costo += $v.costo; $pv.clientes[$v.cliente] = $true
        if (-not $prev.ContainsKey($v.cliente)) { $pv.nuevos[$v.cliente] = $true }
    }
    $porDia = @{}
    foreach ($v in $vs) { if (-not $porDia[$v.fecha]) { $porDia[$v.fecha] = 0.0 }; $porDia[$v.fecha] += $v.q; $diasSet[$v.fecha] = $true }
    $provs = @($cb.alias.Keys | ForEach-Object { $a = $artInfo[$_]; if ($a) { [string]$a.proveedor } } | Sort-Object -Unique)
    $neto = ($vs | Measure-Object neto -Sum).Sum; $costo = ($vs | Measure-Object costo -Sum).Sum
    $nuevos = @($porCli.Keys | Where-Object { -not $prev.ContainsKey($_) }).Count
    $out += [ordered]@{
        codigo = $cb.codigo; descripcion = $cb.descripcion
        proveedor = (($provs | ForEach-Object { if ($provNom[$_]) { $provNom[$_] } else { $_ } }) -join ", ")
        componentes = @($cb.comps | ForEach-Object { [ordered]@{ nombre = ([string]$_.nombre).Trim(); articulo = (([string]$_.articulos -split ',')[0]).Trim(); cantidad = [double]$_.cantidad; descuento = [double]$_.descuento } })
        lanzamiento = $lanz; ultimaVenta = $ult
        diasSinVender = $(if ($ult) { [int]($hoy - [datetime]::ParseExact($ult, "yyyy-MM-dd", $null)).TotalDays } else { $null })
        combos = [double]($vs | Measure-Object q -Sum).Sum; facturas = $vs.Count; clientes = $porCli.Count
        clientesNuevos = $nuevos; clientesQueYaCompraban = $prev.Count
        neto = (R2 $neto); margen = $(if ($neto -gt 0) { R2 (($neto - $costo) / $neto * 100) } else { $null })
        ultimos7 = [double](($vs | Where-Object { $_.fecha -gt $hoy.AddDays(-7).ToString("yyyy-MM-dd") } | Measure-Object q -Sum).Sum)
        dias = $porDia
        vendedores = @($porVend.Values | Sort-Object { -$_.q } | ForEach-Object { [ordered]@{ codigo = $_.codigo; nombre = $_.nombre; combos = $_.q; facturas = $_.facturas; clientes = $_.clientes.Count; nuevos = $_.nuevos.Count; neto = (R2 $_.neto); margen = $(if ($_.neto -gt 0) { R2 (($_.neto - $_.costo) / $_.neto * 100) } else { $null }) } })
        clientesDetalle = @($porCli.Keys | ForEach-Object {
            $c = $cliNom[$_]; $pc = $porCli[$_]
            [ordered]@{ codigo = $_; nombre = $(if ($c) { if ($c.razon_social) { ([string]$c.razon_social).Trim() } else { ([string]$c.nombre).Trim() } } else { "" }); localidad = $(if ($c) { ([string]$c.localidad).Trim() } else { "" })
                vendedor = $pc.vendedor; combos = $pc.q; facturas = $pc.facturas; neto = (R2 $pc.neto); primera = $pc.primera; ultima = $pc.ultima; nuevo = (-not $prev.ContainsKey($_)) }
        } | Sort-Object { -$_.combos })
    }
}
$out = @($out | Sort-Object { if ($_.ultimaVenta) { $_.ultimaVenta } else { '0' } } -Descending)
$data = [ordered]@{
    generatedAt = (Get-Date).ToString("o")
    desde = $desde
    combos = $out
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
[System.IO.File]::WriteAllText($OutPath, (ConvertTo-Json -InputObject $data -Depth 8 -Compress), (New-Object System.Text.UTF8Encoding $false))
if (Test-Path $TemplatePath) { Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force }
foreach ($c in $out) { Write-Log ("{0} {1}: {2} combos, {3} clientes ({4} nuevos), neto {5:N0}, margen {6}%, ultima {7}" -f $c.codigo, $c.descripcion, $c.combos, $c.clientes, $c.clientesNuevos, $c.neto, $c.margen, $c.ultimaVenta) }
