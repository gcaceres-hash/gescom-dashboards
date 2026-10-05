<#
Arma el dataset del tablero de Mercado Libre a partir del "Reporte de ventas"
(Excel) que se descarga a mano desde Mercado Libre (Ventas > Descargar Excel
de ventas, periodo "Ultimo ano"). Las ventas de Mercado Libre NO pasan por
Gescom, por eso este tablero tiene su propia fuente y se actualiza a mano.

Que calcula (todo agregado, NUNCA se guardan datos de compradores):
  - Facturado = "Ingresos por productos" (precio final, con IVA).
  - Venta neta = facturado / (1 + IVA del articulo). El IVA sale de la lista 1
    de Gescom (si el SKU no esta, se asume 21%).
  - Cobrado = columna "Total (ARS)" del reporte: lo que Mercado Libre
    liquida (ya descontados cargos, envio, impuestos y bonificaciones).
  - Cargos de ML = cargo por venta + costo fijo + costo por cuotas.
  - Envio neto = costos de envio menos ingresos por envio.
  - Costo de mercaderia = costo del bulto / unidades del bulto del articulo en
    Gescom (el SKU de Mercado Libre es el codigo de articulo de Gescom) x
    unidades vendidas. Si el SKU no existe en Gescom no se inventa costo: la
    linea queda "sin costo" y no entra en el resultado.
  - Resultado estimado = cobrado - IVA de la venta - costo de mercaderia.
Reglas del reporte: una venta cancelada por el comprador no suma (se informa
aparte); un "paquete de varios productos" trae el dinero en una fila madre y
los productos en filas hijas sin importes: se reparte por precio x unidades.

Uso:  pull-mercadolibre.ps1 -XlsxPath "C:\ruta\Ventas.xlsx"
Cada corrida reemplaza los meses que trae el archivo y conserva los meses
anteriores que ya estaban en data.json (el reporte cubre un maximo de 12 meses).
#>
param(
    [Parameter(Mandatory=$true)][string]$XlsxPath,
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/mercadolibre"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "mercadolibre-template.html"),
    [string]$EquivPath = (Join-Path $PSScriptRoot "ml-equivalencias.json"),
    [double]$IvaDefault = 0.21
)
$ErrorActionPreference = "Stop"
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
function Invoke-PanelSql([string]$Sql) {
    $uri = "$($config.baseUrl)/consulta?sql=" + [uri]::EscapeDataString($Sql)
    $headers = @{ Authorization = "Bearer $($config.clave)" }
    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    if ($r.truncado) { Write-Log "AVISO: la consulta vino truncada (mas de 20000 filas) -- revisar y acotar." }
    return $r.filas
}
function R2($x) { return [math]::Round([double]$x, 2) }

# --- 1) leer el Excel (Excel COM; el reporte tiene titulos arriba y el encabezado en la fila 6) ---
if (-not (Test-Path $XlsxPath)) { throw "No existe el archivo $XlsxPath" }
$ext = [System.IO.Path]::GetExtension($XlsxPath).ToLower()
$origen = $XlsxPath
if ($ext -ne ".xlsx" -and $ext -ne ".xls") {
    $origen = Join-Path ([System.IO.Path]::GetTempPath()) ("ml-ventas-" + [guid]::NewGuid().ToString("N") + ".xlsx")
    Copy-Item $XlsxPath $origen -Force
}
Write-Log "Leyendo $XlsxPath ..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($origen, 0, $true)
    $ws = $wb.Worksheets.Item(1)
    $nFilas = $ws.UsedRange.Rows.Count
    $nCols = $ws.UsedRange.Columns.Count
    $V = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($nFilas,$nCols)).Value2
    $wb.Close($false)
} finally {
    $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    if ($origen -ne $XlsxPath) { Remove-Item $origen -Force -ErrorAction SilentlyContinue }
}

# encabezado: la fila cuya primera celda es "# de venta"
$filaHdr = 0
for ($i = 1; $i -le [math]::Min(20, $nFilas); $i++) { if (([string]$V[$i,1]).Trim() -eq '# de venta') { $filaHdr = $i; break } }
if ($filaHdr -eq 0) { throw "No se encontro la fila de encabezado ('# de venta'). Es el reporte de ventas de Mercado Libre?" }
function Col([string]$patron) {
    for ($j = 1; $j -le $nCols; $j++) { if (([string]$V[$filaHdr,$j]).Trim() -match $patron) { return $j } }
    throw "El reporte no tiene la columna esperada /$patron/."
}
$cVenta = Col '^# de venta$'; $cFecha = Col '^Fecha de venta'; $cEstado = Col '^Estado$'
$cUnid = Col '^Unidades$'; $cIngr = Col '^Ingresos por productos'; $cCargo = Col '^Cargo por venta'
$cFijo = Col '^Costo fijo'; $cCuotas = Col '^Costo por ofrecer cuotas'
$cIngEnv = Col '^Ingresos por env'; $cCostEnv = Col '^Costos de env'
$cEnvMed = Col '^Costo de env.o basado'; $cEnvDif = Col '^Cargo por diferencias'
$cImp = Col '^Impuestos$'; $cDesc = Col '^Descuentos y bonif'; $cAnul = Col '^Anulaciones y reembolsos'
$cTotal = Col '^Total'; $cSku = Col '^SKU$'; $cTitulo = Col '^T.tulo de la publicaci'
$cPrecio = Col '^Precio unitario de venta'; $cEntrega = Col '^Forma de entrega'
function Cel($r, $c) { return $V[$r,$c] }
function Txt($r, $c) { return ([string]$V[$r,$c]).Trim() }
function Num($r, $c) {
    $x = $V[$r,$c]
    if ($x -is [double]) { return $x }
    $s = ([string]$x).Trim(); if ($s -eq '') { return 0.0 }
    $d = 0.0
    if ([double]::TryParse($s.Replace(',','.'), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    return 0.0
}
$MESES_ES = @{ enero=1; febrero=2; marzo=3; abril=4; mayo=5; junio=6; julio=7; agosto=8; septiembre=9; setiembre=9; octubre=10; noviembre=11; diciembre=12 }
function ToFecha([string]$s) {
    if ($s -match '^\s*(\d{1,2}) de ([a-z]+) de (\d{4})') {
        $mes = $MESES_ES[$Matches[2].ToLower()]
        if ($mes) { return [datetime]::new([int]$Matches[3], [int]$mes, [int]$Matches[1]) }
    }
    return $null
}
function GrupoEstado([string]$e) {
    if ($e -match '^Entregado|^Venta entregada') { return 'Entregada' }
    if ($e -match 'camino|retiro') { return 'En camino / punto de retiro' }
    if ($e -match 'Etiqueta|despachar') { return 'Por despachar' }
    if ($e -match 'Cancel') { return 'Cancelada' }
    if ($e -match 'Devol|Reclamo') { return 'Con devolucion / reclamo' }
    return 'Otro'
}

# --- 2) armar lineas (una por producto vendido), repartiendo los paquetes de varios productos ---
$CAMPOS = @('ingresos','cargoVenta','costoFijo','costoCuotas','ingEnvio','costoEnvio','costoEnvioMed','cargoDif','impuestos','descuentos','anulaciones','total')
function NuevaFila($r) {
    return [pscustomobject]@{
        fila = $r; venta = (Txt $r $cVenta); fecha = (ToFecha (Txt $r $cFecha)); estado = (Txt $r $cEstado)
        unidades = (Num $r $cUnid); sku = (Txt $r $cSku); titulo = (Txt $r $cTitulo); precio = (Num $r $cPrecio); entrega = (Txt $r $cEntrega)
        ingresos = (Num $r $cIngr); cargoVenta = (Num $r $cCargo); costoFijo = (Num $r $cFijo); costoCuotas = (Num $r $cCuotas)
        ingEnvio = (Num $r $cIngEnv); costoEnvio = (Num $r $cCostEnv); costoEnvioMed = (Num $r $cEnvMed); cargoDif = (Num $r $cEnvDif)
        impuestos = (Num $r $cImp); descuentos = (Num $r $cDesc); anulaciones = (Num $r $cAnul); total = (Num $r $cTotal)
        peso = 0.0; mes = $null; ventaNeta = 0.0; cargosML = 0.0; envioNeto = 0.0; retenciones = 0.0; costo = $null; resultado = $null; grupo = $null
    }
}
$filas = New-Object System.Collections.Generic.List[object]
for ($r = $filaHdr + 1; $r -le $nFilas; $r++) {
    if ((Txt $r $cVenta) -eq '') { continue }
    $filas.Add((NuevaFila $r))
}
Write-Log "Filas de venta en el reporte: $($filas.Count)"

$lineas = New-Object System.Collections.Generic.List[object]
$cancel = @{}   # venta -> ingresos
$paquetesRepartidos = 0
for ($i = 0; $i -lt $filas.Count; $i++) {
    $f = $filas[$i]
    if ($f.estado -match 'Cancel') { $cancel[$f.venta] = $f.ingresos; continue }
    $esMadre = ($f.ingresos -ne 0 -and $f.sku -eq '' -and $f.unidades -eq 0)
    if ($esMadre) {
        $hijas = @()
        $k = $i + 1
        while ($k -lt $filas.Count -and $filas[$k].ingresos -eq 0 -and $filas[$k].sku -ne '') { $hijas += $filas[$k]; $k++ }
        if ($hijas.Count -eq 0) { Write-Log "Aviso: fila $($f.fila) trae dinero pero no tiene producto ni filas hijas, se omite."; continue }
        $pesoTotal = 0.0
        foreach ($h in $hijas) { $h.peso = [math]::Max(1.0, $h.precio * $h.unidades); $pesoTotal += $h.peso }
        foreach ($h in $hijas) {
            $parte = $h.peso / $pesoTotal
            foreach ($campo in $CAMPOS) { $h.$campo = $f.$campo * $parte }
            $h.venta = $f.venta
            if (-not $h.fecha) { $h.fecha = $f.fecha }
            if (-not $h.entrega) { $h.entrega = $f.entrega }
            $lineas.Add($h)
        }
        $paquetesRepartidos++
        $i = $k - 1
        continue
    }
    if ($f.ingresos -eq 0) { continue }     # hijas ya consumidas / filas vacias
    if (-not $f.fecha) { Write-Log "Aviso: fila $($f.fila) con fecha ilegible '$(Txt $f.fila $cFecha)', se omite."; continue }
    $lineas.Add($f)
}
Write-Log "Lineas de producto: $($lineas.Count) | paquetes repartidos: $paquetesRepartidos | ventas canceladas: $($cancel.Count)"
if ($lineas.Count -eq 0) { throw "No quedaron lineas de venta para procesar." }

# --- 3) costo e IVA desde Gescom ---
$skus = @($lineas | ForEach-Object { $_.sku } | Where-Object { $_ -match '^[A-Za-z0-9]+$' } | Select-Object -Unique)
$info = @{}
if ($skus.Count -gt 0) {
    $lista = ($skus | ForEach-Object { "'" + $_ + "'" }) -join ","
    Write-Log "Buscando costo de $($skus.Count) articulos en Gescom..."
    foreach ($a in (Invoke-PanelSql "SELECT codigo, descripcion, costo_bulto, unidades_bulto FROM articulos WHERE codigo IN ($lista)")) {
        $ub = [double]$a.unidades_bulto
        if ($ub -gt 0 -and $null -ne $a.costo_bulto) { $info[[string]$a.codigo] = @{ costoUnit = ([double]$a.costo_bulto / $ub); desc = [string]$a.descripcion; iva = $null } }
    }
    foreach ($p in (Invoke-PanelSql "SELECT articulo, alicuota_iva FROM precios WHERE lista = 1 AND articulo IN ($lista)")) {
        $k2 = [string]$p.articulo
        if ($info.ContainsKey($k2) -and $null -eq $info[$k2].iva -and $null -ne $p.alicuota_iva) { $info[$k2].iva = [double]$p.alicuota_iva }
    }
}
# Equivalencias: a veces una publicacion de ML vende una cantidad distinta a la "unidad"
# de Gescom (ej. pack de 4 rollos vs bulto de 8). ml-equivalencias.json = { "SKU": factor }
# con el factor = unidades de Gescom que contiene UNA unidad vendida en ML (default 1).
$equiv = @{}
if ($EquivPath -and (Test-Path $EquivPath)) {
    $eqRaw = Get-Content $EquivPath -Raw -Encoding UTF8
    if ($eqRaw.Trim()) {
        foreach ($p in ((ConvertFrom-Json $eqRaw).PSObject.Properties)) { $equiv[[string]$p.Name] = [double]$p.Value }
    }
    if ($equiv.Count -gt 0) { Write-Log "Equivalencias aplicadas: $(($equiv.Keys | ForEach-Object { "$_ x$($equiv[$_])" }) -join ', ')" }
}
$sinCosto = @($skus | Where-Object { -not $info.ContainsKey($_) })
if ($sinCosto.Count -gt 0) { Write-Log "SKUs sin costo en Gescom (quedan 'sin costo'): $($sinCosto -join ', ')" }

foreach ($l in $lineas) {
    $iva = $IvaDefault
    $costoUnit = $null
    if ($info.ContainsKey($l.sku)) {
        $costoUnit = $info[$l.sku].costoUnit
        if ($equiv.ContainsKey($l.sku)) { $costoUnit = $costoUnit * $equiv[$l.sku] }
        if ($null -ne $info[$l.sku].iva) { $iva = $info[$l.sku].iva }
    }
    $l.mes = $l.fecha.ToString("yyyy-MM")
    $l.ventaNeta = $l.ingresos / (1 + $iva)
    $l.cargosML = -($l.cargoVenta + $l.costoFijo + $l.costoCuotas)
    $l.envioNeto = -($l.ingEnvio + $l.costoEnvio + $l.costoEnvioMed + $l.cargoDif)
    $l.retenciones = -$l.impuestos
    if ($null -ne $costoUnit) {
        $l.costo = $costoUnit * $l.unidades
        $l.resultado = $l.total - ($l.ingresos - $l.ventaNeta) - $l.costo
    }
    $l.grupo = GrupoEstado $l.estado
}

# --- 4) agregacion por mes ---
function Sum($lista, [string]$campo) { $s = 0.0; foreach ($x in $lista) { if ($null -ne $x.$campo) { $s += [double]$x.$campo } }; return $s }
function Agrupar($lista, [scriptblock]$clave) {
    $g = [ordered]@{}
    foreach ($x in $lista) { $kk = & $clave $x; if (-not $g.Contains($kk)) { $g[$kk] = New-Object System.Collections.Generic.List[object] }; $g[$kk].Add($x) }
    return $g
}
$fechaMax = ($lineas | Measure-Object -Property fecha -Maximum).Maximum
$mesActualKey = $fechaMax.ToString("yyyy-MM")
$cancelPorMes = @{}
foreach ($f in $filas) {
    if ($f.estado -match 'Cancel' -and $f.fecha) {
        $mk = $f.fecha.ToString("yyyy-MM")
        if (-not $cancelPorMes.ContainsKey($mk)) { $cancelPorMes[$mk] = @{ cant = 0; monto = 0.0 } }
        $cancelPorMes[$mk].cant++; $cancelPorMes[$mk].monto += $f.ingresos
    }
}
$nuevos = [ordered]@{}
foreach ($mk in (@($lineas | ForEach-Object { $_.mes } | Select-Object -Unique | Sort-Object))) {
    $ls = @($lineas | Where-Object { $_.mes -eq $mk })
    $conCosto = @($ls | Where-Object { $null -ne $_.costo })
    $ventasIds = @($ls | ForEach-Object { $_.venta } | Select-Object -Unique)
    $ing = Sum $ls 'ingresos'; $tot = Sum $ls 'total'
    $vnCosto = Sum $conCosto 'ventaNeta'; $res = Sum $conCosto 'resultado'
    $cm = $cancelPorMes[$mk]; if (-not $cm) { $cm = @{ cant = 0; monto = 0.0 } }
    $totales = [ordered]@{
        ventas = $ventasIds.Count; unidades = (Sum $ls 'unidades'); ingresos = (R2 $ing); ventaNeta = (R2 (Sum $ls 'ventaNeta'))
        cobrado = (R2 $tot); cargosML = (R2 (Sum $ls 'cargosML')); envioNeto = (R2 (Sum $ls 'envioNeto')); retenciones = (R2 (Sum $ls 'retenciones'))
        bonificaciones = (R2 (Sum $ls 'descuentos')); costo = (R2 (Sum $conCosto 'costo')); ventaNetaConCosto = (R2 $vnCosto)
        resultado = (R2 $res); margenPct = $(if ($vnCosto -ne 0) { [math]::Round($res / $vnCosto, 4) } else { $null })
        coberturaCostoPct = $(if ((Sum $ls 'ventaNeta') -ne 0) { [math]::Round($vnCosto / (Sum $ls 'ventaNeta'), 4) } else { $null })
        ticketPromedio = $(if ($ventasIds.Count -gt 0) { (R2 ($ing / $ventasIds.Count)) } else { 0 })
        cancelaciones = [ordered]@{ cantidad = $cm.cant; monto = (R2 $cm.monto) }
    }
    $dias = @((Agrupar $ls { param($x) $x.fecha.ToString("yyyy-MM-dd") }).GetEnumerator() | Sort-Object Name | ForEach-Object {
        [ordered]@{ fecha = $_.Name; ventas = @($_.Value | ForEach-Object { $_.venta } | Select-Object -Unique).Count; unidades = (Sum $_.Value 'unidades'); ingresos = (R2 (Sum $_.Value 'ingresos')) } })
    $productos = @((Agrupar $ls { param($x) $x.sku }).GetEnumerator() | ForEach-Object {
        $g = $_.Value; $gc = @($g | Where-Object { $null -ne $_.costo }); $vn = Sum $gc 'ventaNeta'; $rs = Sum $gc 'resultado'
        $tit = ($g | Where-Object { $_.titulo } | Select-Object -First 1).titulo
        [ordered]@{
            sku = $_.Name; titulo = $tit; unidades = (Sum $g 'unidades'); ventas = @($g | ForEach-Object { $_.venta } | Select-Object -Unique).Count
            ingresos = (R2 (Sum $g 'ingresos')); cobrado = (R2 (Sum $g 'total'))
            costo = $(if ($gc.Count -gt 0 -and $gc.Count -eq $g.Count) { (R2 (Sum $g 'costo')) } else { $null })
            resultado = $(if ($gc.Count -gt 0) { (R2 $rs) } else { $null })
            ventaNetaConCosto = (R2 $vn)
            margenPct = $(if ($gc.Count -gt 0 -and $vn -ne 0) { [math]::Round($rs / $vn, 4) } else { $null })
            sinCosto = ($gc.Count -ne $g.Count)
        } } | Sort-Object { -$_.ingresos })
    $entrega = @((Agrupar $ls { param($x) $(if ($x.entrega) { $x.entrega } else { 'Sin dato' }) }).GetEnumerator() | ForEach-Object {
        [ordered]@{ nombre = $_.Name; ventas = @($_.Value | ForEach-Object { $_.venta } | Select-Object -Unique).Count; ingresos = (R2 (Sum $_.Value 'ingresos')) } } | Sort-Object { -$_.ingresos })
    $estados = @((Agrupar $ls { param($x) $x.grupo }).GetEnumerator() | ForEach-Object {
        [ordered]@{ nombre = $_.Name; ventas = @($_.Value | ForEach-Object { $_.venta } | Select-Object -Unique).Count; ingresos = (R2 (Sum $_.Value 'ingresos')) } } | Sort-Object { -$_.ingresos })
    $fechas = @($ls | ForEach-Object { $_.fecha } | Sort-Object)
    $nuevos[$mk] = [ordered]@{
        cerrado = ($mk -ne $mesActualKey)
        periodoDesde = $fechas[0].ToString("yyyy-MM-dd"); periodoHasta = $fechas[$fechas.Count-1].ToString("yyyy-MM-dd")
        totales = $totales; dias = $dias; productos = $productos; entrega = $entrega; estados = $estados
    }
    Write-Log "Mes $mk -> ventas $($totales.ventas) | facturado $($totales.ingresos) | cobrado $($totales.cobrado) | resultado $($totales.resultado) (margen $($totales.margenPct))"
}

# --- 5) fusionar con el historico y guardar ---
$meses = [ordered]@{}
if (Test-Path $OutPath) {
    try {
        $prev = Get-Content $OutPath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($p in $prev.meses.PSObject.Properties) { $meses[$p.Name] = $p.Value }
        Write-Log "Historico cargado desde $OutPath : $($meses.Keys -join ', ')"
    } catch { Write-Log "Aviso: no se pudo leer el data.json existente, se arranca sin historico ($($_.Exception.Message))" }
}
foreach ($mk in $nuevos.Keys) {
    if ($meses.Contains($mk) -and $mk -ne $mesActualKey -and [double]$meses[$mk].totales.ventas -gt [double]$nuevos[$mk].totales.ventas) {
        Write-Log "Aviso: el archivo trae solo parte del mes cerrado $mk ($($nuevos[$mk].totales.ventas) ventas vs $($meses[$mk].totales.ventas) guardadas) -- se conserva el historico."
        continue
    }
    $meses[$mk] = $nuevos[$mk]
}
$ordenadas = [ordered]@{}
foreach ($mk in ($meses.Keys | Sort-Object)) { $ordenadas[$mk] = $meses[$mk] }
$out = [ordered]@{
    generadoEn = (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
    fuente = "Reporte de ventas de Mercado Libre (Excel, descarga manual) + costo e IVA de Gescom"
    mesActual = $mesActualKey
    supuestos = [ordered]@{ ivaPorDefecto = $IvaDefault; skusSinCosto = @($sinCosto) }
    meses = $ordenadas
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
$json = ConvertTo-Json -InputObject $out -Depth 10
[System.IO.File]::WriteAllText($OutPath, $json, (New-Object System.Text.UTF8Encoding $false))
Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force
Write-Log "Guardado: $OutPath (meses: $($ordenadas.Keys -join ', '))"
