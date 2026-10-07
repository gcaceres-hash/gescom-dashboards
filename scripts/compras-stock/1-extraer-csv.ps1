param(
    [string]$CsvAnual = "C:\Users\gcaceres\Downloads\ventas-Detallado de ventas extendido-20261001-161119.csv",
    [string]$CsvReciente = "C:\Users\gcaceres\Downloads\ventas-Detallado de ventas extendido-20261006-145231.csv"
)
$ErrorActionPreference = "Stop"
$sp = if ($env:COMPRAS_TMP) { $env:COMPRAS_TMP } else { Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "tmp" }; if (-not (Test-Path $sp)) { New-Item -ItemType Directory $sp -Force | Out-Null }
function N($x) { $s = ([string]$x).Trim(); if ($s -eq '') { return 0.0 }; $d = 0.0; if ([double]::TryParse($s.Replace(',','.'), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }; return 0.0 }
$arch = @($CsvAnual, $CsvReciente)
$art = @{}          # codigo -> datos
$dias = @{}         # fecha -> filas (para detectar dias de operacion)
$tipos = @{}
$vistos = 0
foreach ($f in $arch) {
  Write-Host ("Leyendo " + $f)
  $rows = Import-Csv $f -Delimiter ';' -Encoding Default
  foreach ($r in $rows) {
    $fecha = ([string]$r.FechaComprobante).Substring(0,10)
    if ($fecha -lt "2026-07-07") { continue }
    if ($f -eq $CsvReciente -and $fecha -lt "2026-10-01") { continue }   # el export chico trae 30/9 repetido
    $vistos++
    $tipos[[string]$r.TipoDeVenta] = 1 + [int]$tipos[[string]$r.TipoDeVenta]
    if (@("1176","43") -contains [string]$r.CodVendedor) { continue }
    if ((N $r.PrecioCosto) -eq 1.0) { continue }
    $nom = ([string]$r.Articulo).Trim()
    if ($nom -eq 'Productos varios' -or $nom -eq 'DESCUENTO') { continue }
    $dias[$fecha] = 1 + [int]$dias[$fecha]
    $cod = [string]$r.Codigo
    if (-not $art.ContainsKey($cod)) { $art[$cod] = [ordered]@{ codigo = $cod; articulo = $nom; proveedor = ([string]$r.Proveedor).Trim(); familia = ([string]$r.Familia).Trim(); marca = ([string]$r.Marca).Trim(); jul = 0.0; julN = 0.0; ago = 0.0; agoN = 0.0; sep = 0.0; sepN = 0.0; oct = 0.0; octN = 0.0; agoPeso = 0.0 } }
    $a = $art[$cod]
    if ($a.familia -eq '' -and ([string]$r.Familia).Trim() -ne '') { $a.familia = ([string]$r.Familia).Trim() }
    $tipo = [string]$r.TipoDeVenta
    $cant = N $r.CantBase; $neto = N $r.ImporteNetoItem
    # criterio de la base: venta (+) y rechazo (-); los canjes por vencimiento no cuentan
    $cuenta = ($tipo -eq 'Venta') -or ($tipo -like 'Devoluci*Rechazo*') -or ($tipo -like 'Nota*D*bito*')
    if (-not $cuenta) { continue }
    if ($fecha -lt "2026-08-03") { $a.jul += $cant; $a.julN += $neto }
    elseif ($fecha -lt "2026-09-01") { $a.ago += $cant; $a.agoN += $neto }
    elseif ($fecha -lt "2026-10-01") { $a.sep += $cant; $a.sepN += $neto }
    else { $a.oct += $cant; $a.octN += $neto }
  }
  $rows = $null; [GC]::Collect()
}
Write-Host "Filas desde 7/7 revisadas: $vistos | articulos: $($art.Count)"
Write-Host "Tipos de venta: " + (($tipos.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join " | ")
$out = [ordered]@{ articulos = @($art.Values); dias = @($dias.GetEnumerator() | Sort-Object Name | ForEach-Object { [ordered]@{ fecha = $_.Name; filas = $_.Value } }) }
[System.IO.File]::WriteAllText("$sp\compras-csv.json", (ConvertTo-Json -InputObject $out -Depth 5 -Compress), (New-Object System.Text.UTF8Encoding $false))
Write-Host "Guardado compras-csv.json"
