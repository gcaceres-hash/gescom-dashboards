<#
Compras: stock, dias de stock, sugerido de pedido y rotacion, por producto.
Alimenta el tablero docs/compras-stock (protegido con contrasena: los datos se publican cifrados).

Todo sale de la base compartida de Lucas (datos-gescom), asi que se actualiza solo todos los dias:
  - Stock: deposito PRI (el principal; el 2 es CAMBIOS y no cuenta).
  - Venta: facturas y notas de debito menos rechazos (VEN + DEB - DEV-RE) por producto, en una ventana movil
    de hasta 90 dias que termina AYER (hoy esta incompleto). La base empieza el 3/8/2026: hasta que junte 90
    dias de historia se completa con el bloque de julio de scripts/compras-stock/julio.json (sacado del export
    de Gescom), que va saliendo solo de la ventana y desaparece cuando la base ya tiene los 90 dias.
  - Unidades: las mismas que usa el stock. Los productos que se pesan vienen en gramos (unidades_bulto 1000).
  - Costo: costo_bulto / unidades_bulto, sin impuestos.
  - Categoria: la "familia" de Gescom; no esta en la base, sale de scripts/categorias-articulos.json, que se
    completa cada vez que se carga el CSV de ventas (pull-venta-rentabilidad-csv.ps1).
  - Clase ABC por proveedor (Pareto de la venta neta de la ventana): A hasta el 80%, B hasta el 95%, C el resto.
  - Compras: ingresos cargados (remitos y facturas de compra, sin notas de credito/debito ni liquidaciones).
Los parametros (plazos, dias de seguridad, umbrales) se cambian en el propio tablero.
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/compras-stock"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "compras-stock-template.html"),
    [string]$SuplementoPath = (Join-Path $PSScriptRoot "compras-stock/julio.json"),
    [string]$CategoriasPath = (Join-Path $PSScriptRoot "categorias-articulos.json"),
    [int]$VentanaDias = 90
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
$inv = [Globalization.CultureInfo]::InvariantCulture
$utf8 = New-Object System.Text.UTF8Encoding $false
$hoy = (Get-Date).Date
$ayer = $hoy.AddDays(-1)
$sinCat = "Sin categor" + [string][char]0xED + "a"

# --- ventana ---
$baseDesdeTxt = ([string](Invoke-PanelSql "SELECT MIN(fecha) AS d FROM ventas")[0].d).Substring(0, 10)
$baseDesde = [datetime]::ParseExact($baseDesdeTxt, "yyyy-MM-dd", $inv)
$w = $ayer.AddDays(-($VentanaDias - 1))
$desde = if ($w -gt $baseDesde) { $w } else { $baseDesde }
Write-Log "Ventana de venta: $($desde.ToString('yyyy-MM-dd')) a $($ayer.ToString('yyyy-MM-dd')) (la base empieza el $baseDesdeTxt)"
$desdeTxt = $desde.ToString("yyyy-MM-dd"); $hastaTxt = $ayer.ToString("yyyy-MM-dd")
$inicioMes = (Get-Date -Year $hoy.Year -Month $hoy.Month -Day 1).Date.ToString("yyyy-MM-dd")

Write-Log "Descargando catalogo, stock, ventas y compras..."
$arts = Invoke-PanelSql "SELECT codigo, descripcion, proveedor, unidades_bulto, costo_bulto FROM articulos"
$provNombre = @{}; foreach ($p in (Invoke-PanelSql "SELECT codigo, nombre FROM proveedores")) { $provNombre[[string]$p.codigo] = [string]$p.nombre }
$stock = @{}; foreach ($s in (Invoke-PanelSql "SELECT articulo, SUM(cantidad) AS cantidad FROM stock WHERE deposito = 'PRI' GROUP BY articulo")) { $stock[[string]$s.articulo] = [double]$s.cantidad }
$ventas = Invoke-PanelSql "SELECT i.articulo AS articulo, v.tipo AS tipo, SUM(i.cantidad * i.factor) AS unidades, SUM(i.neto) AS neto, MAX(v.fecha) AS ultima FROM venta_items i JOIN ventas v ON v.id = i.venta_id WHERE v.fecha >= '$desdeTxt' AND v.fecha <= '$hastaTxt' AND v.tipo IN ('VEN','DEB','DEV-RE') GROUP BY i.articulo, v.tipo"
$ultVenta = @{}; foreach ($u in (Invoke-PanelSql "SELECT i.articulo AS articulo, MAX(v.fecha) AS ultima FROM venta_items i JOIN ventas v ON v.id = i.venta_id WHERE v.tipo IN ('VEN','DEB') GROUP BY i.articulo")) { $ultVenta[[string]$u.articulo] = ([string]$u.ultima).Substring(0, 10) }
$compras = @{}; foreach ($c in (Invoke-PanelSql "SELECT i.articulo AS articulo, MAX(c.carga) AS ultima, SUM(CASE WHEN c.carga >= '$inicioMes' THEN i.cantidad * i.factor ELSE 0 END) AS u_mes FROM compra_items i JOIN compras c ON c.clave = i.compra WHERE c.tipo NOT LIKE 'NCR%' AND c.tipo NOT LIKE 'NDB%' AND c.tipo NOT LIKE 'LIQ%' GROUP BY i.articulo")) { $compras[[string]$c.articulo] = $c }
$diasOp = Invoke-PanelSql "SELECT fecha, COUNT(*) AS n FROM ventas WHERE tipo = 'VEN' AND fecha >= '$desdeTxt' AND fecha <= '$hastaTxt' GROUP BY fecha"
$diasBase = @($diasOp | Where-Object { $_.n -ge 50 }).Count
Write-Log "Articulos: $($arts.Count) | con stock: $($stock.Count) | filas de venta: $($ventas.Count) | dias con operacion en la base: $diasBase"

$vb = @{}
foreach ($v in $ventas) {
    $k = [string]$v.articulo
    if (-not $vb.ContainsKey($k)) { $vb[$k] = @{ u = 0.0; n = 0.0 } }
    $sg = if ($v.tipo -eq 'DEV-RE') { -1.0 } else { 1.0 }
    $vb[$k].u += $sg * [double]$v.unidades; $vb[$k].n += $sg * [double]$v.neto
}

# --- bloque de julio (solo mientras la base no tenga los 90 dias) ---
$sup = $null; $fSup = 0.0; $supDias = 0.0
if (Test-Path $SuplementoPath) {
    $sup = [System.IO.File]::ReadAllText($SuplementoPath, $utf8) | ConvertFrom-Json
    $sd = [datetime]::ParseExact($sup.desde, "yyyy-MM-dd", $inv); $sh = [datetime]::ParseExact($sup.hasta, "yyyy-MM-dd", $inv)
    $ini = if ($w -gt $sd) { $w } else { $sd }
    if ($ini -le $sh) {
        $fSup = (($sh - $ini).Days + 1) / (($sh - $sd).Days + 1)
        $supDias = [double]$sup.dias * $fSup
    }
    Write-Log ("Bloque de julio: " + $(if ($fSup -gt 0) { "se usa el {0:P0} ({1:N1} dias de operacion)" -f $fSup, $supDias } else { "ya salio de la ventana" }))
}
$supArt = @{}; if ($sup) { foreach ($p in $sup.articulos.PSObject.Properties) { $supArt[$p.Name] = $p.Value } }
$pesConv = @{}; if ($sup) { foreach ($c in $sup.pesablesConvertidos) { $pesConv[[string]$c] = $true } }

$cat = @{}
if (Test-Path $CategoriasPath) { foreach ($p in ([System.IO.File]::ReadAllText($CategoriasPath, $utf8) | ConvertFrom-Json).PSObject.Properties) { $cat[$p.Name] = [string]$p.Value } }

$filas = New-Object System.Collections.Generic.List[object]
foreach ($a in $arts) {
    $cod = [string]$a.codigo
    $st = if ($stock.ContainsKey($cod)) { $stock[$cod] } else { 0.0 }
    $u = 0.0; $n = 0.0
    if ($vb.ContainsKey($cod)) { $u = $vb[$cod].u; $n = $vb[$cod].n }
    $ubn = if ($null -ne $a.unidades_bulto -and [double]$a.unidades_bulto -gt 0) { [double]$a.unidades_bulto } else { 1.0 }
    $dias = [double]$diasBase
    if ($fSup -gt 0) {
        $esPes = ($ubn -ge 1000)
        if (-not $esPes -or $pesConv.ContainsKey($cod)) {
            $dias += $supDias
            if ($supArt.ContainsKey($cod)) { $u += [double]$supArt[$cod][0] * $fSup; $n += [double]$supArt[$cod][1] * $fSup }
        }
    }
    if ($st -le 0 -and $u -eq 0 -and $n -eq 0) { continue }
    $cb = if ($null -ne $a.costo_bulto) { [double]$a.costo_bulto } else { 0.0 }
    $cx = if ($compras.ContainsKey($cod)) { $compras[$cod] } else { $null }
    $filas.Add([pscustomobject]@{
        c = $cod; n = [string]$a.descripcion
        p = $(if ($provNombre.ContainsKey([string]$a.proveedor)) { $provNombre[[string]$a.proveedor] } else { "Proveedor " + [string]$a.proveedor })
        g = $(if ($cat.ContainsKey($cod)) { $cat[$cod] } else { $sinCat })
        u = $ubn; k = [math]::Round($cb / $ubn, 6); s = [math]::Round($st, 3)
        v = [math]::Round($u, 3); m = [math]::Round($n, 2); d = [math]::Round($dias, 1); a = "-"
        uv = $(if ($ultVenta.ContainsKey($cod)) { $ultVenta[$cod] } else { "" })
        uc = $(if ($cx) { ([string]$cx.ultima).Substring(0, 10) } else { "" })
        cm = $(if ($cx) { [math]::Round([double]$cx.u_mes, 3) } else { 0 })
    })
}
# clase ABC por proveedor
foreach ($g in ($filas | Group-Object p)) {
    $con = @($g.Group | Where-Object { $_.m -gt 0 } | Sort-Object m -Descending)
    $tot = 0.0; foreach ($x in $con) { $tot += $x.m }
    $cum = 0.0
    foreach ($x in $con) { $x.a = if (($cum / $tot) -lt 0.80) { "A" } elseif (($cum / $tot) -lt 0.95) { "B" } else { "C" }; $cum += $x.m }
}
$prov = @($filas | Select-Object -ExpandProperty p -Unique | Sort-Object)
$out = [ordered]@{
    generatedAt = (Get-Date).ToString("o")
    fechaCorte = $hoy.ToString("yyyy-MM-dd")
    inicioMes = $inicioMes
    ventana = [ordered]@{ desde = $(if ($fSup -gt 0) { $ini.ToString("yyyy-MM-dd") } else { $desdeTxt }); hasta = $hastaTxt; diasBase = $diasBase; diasJulio = [math]::Round($supDias, 1); usaJulio = ($fSup -gt 0); diasTotal = [math]::Round($diasBase + $supDias, 1) }
    params = [ordered]@{ plazo = 7; segA = 7; segB = 5; segC = 2; alto = 60; muyLento = 120; reciente = 30 }
    proveedores = $prov
    articulos = $filas
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
[System.IO.File]::WriteAllText($OutPath, (ConvertTo-Json -InputObject $out -Depth 5 -Compress), $utf8)
if (Test-Path $TemplatePath) { Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force }
$sv = 0.0; foreach ($x in $filas) { $sv += $x.s * $x.k }
Write-Log ("Guardado: $OutPath | productos: {0} | proveedores: {1} | stock a costo: {2:N0} | ventana: {3} dias de operacion" -f $filas.Count, $prov.Count, $sv, ($diasBase + $supDias))
