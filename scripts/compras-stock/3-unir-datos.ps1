$ErrorActionPreference = "Stop"
$sp = if ($env:COMPRAS_TMP) { $env:COMPRAS_TMP } else { Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "tmp" }; if (-not (Test-Path $sp)) { New-Item -ItemType Directory $sp -Force | Out-Null }
$utf8 = New-Object System.Text.UTF8Encoding $false
$b = [System.IO.File]::ReadAllText("$sp\compras-base.json", $utf8) | ConvertFrom-Json
$c = [System.IO.File]::ReadAllText("$sp\compras-csv.json", $utf8) | ConvertFrom-Json

$provNombre = @{}; foreach ($p in $b.proveedores) { $provNombre[[string]$p.codigo] = [string]$p.nombre }
$stock = @{}; foreach ($s in $b.stock) { $stock[[string]$s.articulo] = [double]$s.cantidad }
$comp = @{}; foreach ($x in $b.compras) { $comp[[string]$x.articulo] = $x }
# ventas de la base: neto de rechazos (VEN + DEB - DEV-RE)
$bv = @{}
foreach ($v in $b.ventas) {
    $k = [string]$v.articulo
    if (-not $bv.ContainsKey($k)) { $bv[$k] = @{ u = 0.0; n = 0.0; ago = 0.0; ultima = "" } }
    $sg = if ($v.tipo -eq 'DEV-RE') { -1.0 } else { 1.0 }
    $bv[$k].u += $sg * [double]$v.unidades; $bv[$k].n += $sg * [double]$v.neto
    if ($v.mes -eq '2026-08') { $bv[$k].ago += $sg * [double]$v.unidades }
    if ($v.tipo -ne 'DEV-RE' -and [string]$v.ultima -gt $bv[$k].ultima) { $bv[$k].ultima = [string]$v.ultima }
}
$cs = @{}; foreach ($a in $c.articulos) { $cs[[string]$a.codigo] = $a }

# dias de operacion: base (desde 3/8) + export (7/7 al 2/8)
$diasBase = @($b.dias | Where-Object { $_.n -ge 50 }).Count
$diasJul = @($c.dias | Where-Object { $_.fecha -lt '2026-08-03' -and $_.filas -ge 50 }).Count
$diasTotal = $diasBase + $diasJul
Write-Host "Dias de operacion: base $diasBase + julio $diasJul = $diasTotal"

$filas = New-Object System.Collections.Generic.List[object]
$kStats = @{ ok = 0; ajustados = 0; sinAgosto = 0; raros = @() }
foreach ($a in $b.articulos) {
    $cod = [string]$a.codigo
    $st = if ($stock.ContainsKey($cod)) { $stock[$cod] } else { 0.0 }
    $ventaB = if ($bv.ContainsKey($cod)) { $bv[$cod] } else { $null }
    $csvA = if ($cs.ContainsKey($cod)) { $cs[$cod] } else { $null }
    $julU = 0.0; $julN = 0.0; $diasArt = $diasTotal
    if ($csvA) {
        $julN = [double]$csvA.julN
        $k = 1.0
        $esPesable = ([double]$a.unidades_bulto -ge 1000)   # la base los cuenta en gramos; el export en piezas
        $baseAgo = if ($ventaB) { $ventaB.ago } else { 0.0 }
        if ($esPesable) {
            if ([double]$csvA.ago -ne 0 -and $baseAgo -ne 0) { $k = $baseAgo / [double]$csvA.ago; $kStats.ajustados++ }
            else { $k = 0.0; $julN = 0.0; $diasArt = $diasBase; $kStats.sinAgosto++ }   # sin forma de convertir: solo los dias de la base
        } else { $kStats.ok++ }
        $julU = [double]$csvA.jul * $k
    } elseif ([double]$a.unidades_bulto -ge 1000) { $diasArt = $diasBase }
    elseif (-not $csvA) { $diasArt = $diasTotal }
    $baseU = if ($ventaB) { $ventaB.u } else { 0.0 }; $baseN = if ($ventaB) { $ventaB.n } else { 0.0 }
    $u3 = $baseU + $julU; $n3 = $baseN + $julN
    if ($st -le 0 -and $u3 -eq 0 -and $n3 -eq 0) { continue }
    $ub = [double]$a.unidades_bulto; if ($ub -le 0) { $ub = 1.0 }
    $cb = if ($null -ne $a.costo_bulto) { [double]$a.costo_bulto } else { 0.0 }
    $fam = if ($csvA -and $csvA.familia) { [string]$csvA.familia } else { "" }
    $cx = if ($comp.ContainsKey($cod)) { $comp[$cod] } else { $null }
    $filas.Add([pscustomobject]@{
        codigo = $cod; producto = [string]$a.descripcion
        proveedor = $(if ($provNombre.ContainsKey([string]$a.proveedor)) { $provNombre[[string]$a.proveedor] } else { "Proveedor " + [string]$a.proveedor })
        categoria = $(if ($fam) { $fam } else { "Sin categor" + [char]0xED + "a" })
        marca = $(if ($csvA) { [string]$csvA.marca } else { "" })
        ub = $ub; costoUnit = [math]::Round($cb / $ub, 6); stock = $st
        venta3mU = [math]::Round($u3, 3); venta3mN = [math]::Round($n3, 2); dias = $diasArt
        ultimaVenta = $(if ($ventaB -and $ventaB.ultima) { $ventaB.ultima } else { "" })
        ultimaCompra = $(if ($cx) { [string]$cx.ultima } else { "" })
        comprasMesU = $(if ($cx) { [double]$cx.u_mes } else { 0.0 })
        clase = ""
    })
}
Write-Host "Articulos en Datos: $($filas.Count) | no pesables (unidades iguales): $($kStats.ok) | pesables convertidos con agosto: $($kStats.ajustados) | pesables sin agosto (solo dias de la base): $($kStats.sinAgosto)"
if ($kStats.raros.Count) { Write-Host "k fuera de 0,5-2 (revisar):"; $kStats.raros | Select-Object -First 12 | ForEach-Object { Write-Host "   $_" } }

# ABC por proveedor (Pareto de la venta neta de 3 meses): A hasta 80%, B hasta 95%, C el resto
foreach ($g in ($filas | Group-Object proveedor)) {
    $con = @($g.Group | Where-Object { $_.venta3mN -gt 0 } | Sort-Object venta3mN -Descending)
    $tot = 0.0; foreach ($x in $con) { $tot += $x.venta3mN }
    $cum = 0.0
    foreach ($x in $con) {
        $x.clase = if ($tot -le 0) { "C" } elseif (($cum / $tot) -lt 0.80) { "A" } elseif (($cum / $tot) -lt 0.95) { "B" } else { "C" }
        $cum += $x.venta3mN
    }
    foreach ($x in $g.Group) { if (-not $x.clase) { $x.clase = "Sin venta" } }
}
$out = [ordered]@{ hoy = [string]$b.hoy; diasBase = $diasBase; diasJulio = $diasJul; diasTotal = $diasTotal; ventanaDesde = "2026-07-07"; articulos = $filas }
[System.IO.File]::WriteAllText("$sp\compras-datos.json", (ConvertTo-Json -InputObject $out -Depth 5 -Compress), $utf8)

Write-Host "--- resumen ---"
$filas | Group-Object clase | ForEach-Object { "clase {0}: {1} articulos" -f $_.Name, $_.Count }
$sinCat = @($filas | Where-Object { $_.categoria -like 'Sin categor*' })
"sin categoria: {0} articulos | venta neta 3m sin categoria: {1:N0} de {2:N0}" -f $sinCat.Count, ($sinCat | Measure-Object venta3mN -Sum).Sum, ($filas | Measure-Object venta3mN -Sum).Sum
"proveedores: " + @($filas | Select-Object -ExpandProperty proveedor -Unique).Count
"stock valorizado total: {0:N0}" -f (($filas | ForEach-Object { $_.stock * $_.costoUnit } | Measure-Object -Sum).Sum)
"venta 3m neta total: {0:N0}" -f (($filas | Measure-Object venta3mN -Sum).Sum)
