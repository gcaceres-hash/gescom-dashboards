$ErrorActionPreference = "Stop"
$sp = if ($env:COMPRAS_TMP) { $env:COMPRAS_TMP } else { Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "tmp" }; if (-not (Test-Path $sp)) { New-Item -ItemType Directory $sp -Force | Out-Null }
$salida = "C:\Users\gcaceres\Desktop\Compras - Stock, pedidos y rotación.xlsx"
$utf8 = New-Object System.Text.UTF8Encoding $false
$d = [System.IO.File]::ReadAllText("$sp\compras-datos.json", $utf8) | ConvertFrom-Json
$hoy = [datetime]::ParseExact($d.hoy, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
function Fecha($s) { if (-not $s) { return $null }; return [datetime]::ParseExact(([string]$s).Substring(0,10), "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture).ToOADate() }
function Seguro([string]$t) { if ($t -match '^[=+\-@]') { return "'" + $t }; return $t }

# ---------- datos ordenados ----------
$orden = @{ "A" = 1; "B" = 2; "C" = 3; "Sin venta" = 4 }
$art = @($d.articulos | Sort-Object proveedor, @{ e = { -[double]$_.venta3mN } })
$N = $art.Count
$ultFila = $N + 1
$prov = @($art | Select-Object -ExpandProperty proveedor -Unique | Sort-Object)

# metricas por articulo calculadas aca (solo para ordenar/seleccionar las listas; los numeros del Excel son formulas)
$m = @{}
for ($i = 0; $i -lt $N; $i++) {
    $a = $art[$i]
    $vd = if ($a.dias -gt 0) { $a.venta3mU / $a.dias } else { 0.0 }
    $m[$i] = [pscustomobject]@{ vd = $vd; dst = $(if ($vd -gt 0) { $a.stock / $vd } else { 1e9 }); val = $a.stock * $a.costoUnit }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.ScreenUpdating = $false
$wb = $excel.Workbooks.Add()
$nombres = @("Días de stock", "Por categoría", "Sin compra del mes", "Sugerido de pedido", "Stock valorizado", "Stock alto - venta lenta", "Parámetros", "Datos", "Notas")
while ($wb.Worksheets.Count -lt $nombres.Count) { [void]$wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count)) }
for ($i = 0; $i -lt $nombres.Count; $i++) { $wb.Worksheets.Item($i + 1).Name = $nombres[$i] }
$wsD = $wb.Worksheets.Item("Datos"); $wsP = $wb.Worksheets.Item("Parámetros"); $wsN = $wb.Worksheets.Item("Notas")
foreach ($ws in $wb.Worksheets) { $ws.Cells.Font.Name = "Arial"; $ws.Cells.Font.Size = 10 }
# precalentamiento: una primera escritura chica de una matriz evita un error de conversion de Excel al cargar la tabla grande
$w = New-Object 'object[,]' 2, 2; $w[0,0] = "a"; $w[0,1] = 1.5; $w[1,0] = "b"; $w[1,1] = 2.5
$wsD.Range("AB1:AC2").Value2 = $w; $wsD.Range("AB1:AC2").ClearContents() | Out-Null
$HDR = 3615007; $BLANCO = 16777215; $AZUL = 16711680; $AMARILLO = 13434879; $GRIS = 15921906; $ROJO = 13551615; $NARANJA = 10284031; $MUYLENTO = 8696052; $VERDE = 13561798
$MONEY = '$#.##0;($#.##0);-'; $UNID = '#.##0;-#.##0;-'; $DEC1 = '#.##0,0'; $PCT = '0,0%'
function Encabezado($r) { $r.Font.Bold = $true; $r.Font.Color = $BLANCO; $r.Interior.Color = $HDR; $r.WrapText = $true; $r.VerticalAlignment = -4108; $r.HorizontalAlignment = -4108 }
function Entrada($c) { $c.Font.Color = $AZUL; $c.Interior.Color = $AMARILLO }
function Nombre($n, $ref) { [void]$wb.Names.Add($n, $ref) }
function Ancho($ws, $anchos) { for ($i = 0; $i -lt $anchos.Count; $i++) { $ws.Columns.Item($i + 1).ColumnWidth = $anchos[$i] } }
function Titulo($ws, $t, $sub) { $ws.Cells.Item(1,1) = $t; $ws.Cells.Item(1,1).Font.Bold = $true; $ws.Cells.Item(1,1).Font.Size = 14; if ($sub) { $ws.Cells.Item(2,1) = $sub; $ws.Cells.Item(2,1).Font.Italic = $true; $ws.Cells.Item(2,1).Font.Color = 5658198 } }
function Congelar($ws, $celda) { $ws.Activate(); $excel.ActiveWindow.FreezePanes = $false; $ws.Range($celda).Select(); $excel.ActiveWindow.FreezePanes = $true }

# ---------- Parámetros ----------
Titulo $wsP "Parámetros" "Las celdas azules sobre amarillo se pueden cambiar: todo el libro se recalcula."
$wsP.Cells.Item(3,1) = "Fecha de corte (stock y datos al)"; $wsP.Cells.Item(3,2).Value2 = $hoy.ToOADate(); $wsP.Cells.Item(3,2).NumberFormat = "dd/mm/yyyy"
$wsP.Cells.Item(4,1) = "Período de venta analizado"; $wsP.Cells.Item(4,2) = "7/7/2026 al " + $hoy.ToString("d/M/yyyy") + " (3 meses)"
$wsP.Cells.Item(5,1) = "Días hábiles con operación en el período"; $wsP.Cells.Item(5,2).Value2 = [double]$d.diasTotal
$wsP.Cells.Item(7,1) = "Plazo de reposición por defecto (días entre pedir y recibir)"; $wsP.Cells.Item(7,2).Value2 = 7
$wsP.Cells.Item(8,1) = "Stock de seguridad productos A (días de venta)"; $wsP.Cells.Item(8,2).Value2 = 7
$wsP.Cells.Item(9,1) = "Stock de seguridad productos B (días de venta)"; $wsP.Cells.Item(9,2).Value2 = 5
$wsP.Cells.Item(10,1) = "Stock de seguridad productos C (días de venta)"; $wsP.Cells.Item(10,2).Value2 = 2
$wsP.Cells.Item(12,1) = "Stock alto: más de (días de stock)"; $wsP.Cells.Item(12,2).Value2 = 60
$wsP.Cells.Item(13,1) = "Venta muy lenta: más de (días de stock)"; $wsP.Cells.Item(13,2).Value2 = 120
$wsP.Cells.Item(14,1) = "Venta reciente: se considera vigente si vendió en los últimos (días)"; $wsP.Cells.Item(14,2).Value2 = 30
foreach ($c in @("B7","B8","B9","B10","B12","B13","B14")) { Entrada $wsP.Range($c) }
$wsP.Cells.Item(15,1) = "Clasificación ABC (por proveedor, sobre la venta neta de 3 meses)"; $wsP.Cells.Item(15,1).Font.Bold = $true
$wsP.Cells.Item(16,1) = "A = los productos que juntan hasta el 80% de la venta del proveedor"
$wsP.Cells.Item(17,1) = "B = del 80% al 95%  |  C = el 5% restante  |  Sin venta = sin movimiento en el período"
$wsP.Cells.Item(17,1).Font.Italic = $true
$wsP.Cells.Item(19,1) = "Plazo de reposición por proveedor (días)"; $wsP.Cells.Item(19,1).Font.Bold = $true
$wsP.Cells.Item(20,1) = "Proveedor"; $wsP.Cells.Item(20,2) = "Plazo (días)"; Encabezado $wsP.Range("A20:B20")
for ($i = 0; $i -lt $prov.Count; $i++) { $wsP.Cells.Item(21 + $i, 1) = Seguro $prov[$i]; $wsP.Cells.Item(21 + $i, 2).Formula = "=Plazo_def"; }
$ultP = 20 + $prov.Count
Entrada $wsP.Range("B21:B$ultP")
$wsP.Cells.Item(21 + $prov.Count + 1, 1) = "Cada proveedor arranca con el plazo por defecto (B7). Escribí el plazo real encima para los que tarden más o menos."
$wsP.Cells.Item(21 + $prov.Count + 1, 1).Font.Italic = $true
Ancho $wsP @(62, 30)
Nombre "Fecha_corte" "=Parámetros!`$B`$3"; Nombre "Plazo_def" "=Parámetros!`$B`$7"
Nombre "Seg_A" "=Parámetros!`$B`$8"; Nombre "Seg_B" "=Parámetros!`$B`$9"; Nombre "Seg_C" "=Parámetros!`$B`$10"
Nombre "Dias_alto" "=Parámetros!`$B`$12"; Nombre "Dias_muy_lento" "=Parámetros!`$B`$13"; Nombre "Dias_reciente" "=Parámetros!`$B`$14"
Nombre "t_Plazos" "=Parámetros!`$A`$21:`$B`$$ultP"

# ---------- Datos ----------
$cab = @("Código","Producto","Proveedor","Categoría","Clase ABC","Unid. por bulto","Costo unit. s/IVA","Stock (u)","Venta 3m (u)","Venta 3m (`$ neto)","Días considerados","Venta diaria (u)","Días de stock","Stock (bultos)","Stock a costo (`$)","Venta diaria a costo (`$)","Última venta","Última compra","Compras desde 1/10 (u)","Días sin comprar","Seguridad (días)","Plazo (días)","Stock objetivo (u)","A pedir (u)","Bultos a pedir","Importe a pedir (`$)","Días sin vender")
for ($c = 0; $c -lt $cab.Count; $c++) { $wsD.Cells.Item(1, $c + 1) = $cab[$c] }
Encabezado $wsD.Range("A1:AA1"); $wsD.Rows.Item(1).RowHeight = 42
$arr = New-Object 'object[,]' $N, 11
for ($i = 0; $i -lt $N; $i++) {
    $a = $art[$i]
    $arr[$i,0] = Seguro $a.codigo; $arr[$i,1] = Seguro $a.producto; $arr[$i,2] = Seguro $a.proveedor; $arr[$i,3] = Seguro $a.categoria; $arr[$i,4] = $a.clase
    $arr[$i,5] = [double]$a.ub; $arr[$i,6] = [double]$a.costoUnit; $arr[$i,7] = [double]$a.stock; $arr[$i,8] = [double]$a.venta3mU; $arr[$i,9] = [double]$a.venta3mN; $arr[$i,10] = [double]$a.dias
}
$wsD.Range("A2:A$ultFila").NumberFormatLocal = "@"
$wsD.Range("A2:K$ultFila").Value2 = $arr
# fechas y compras (columnas Q, R, S)
$arrF = New-Object 'object[,]' $N, 3
for ($i = 0; $i -lt $N; $i++) { $a = $art[$i]; $arrF[$i,0] = Fecha $a.ultimaVenta; $arrF[$i,1] = Fecha $a.ultimaCompra; $arrF[$i,2] = [double]$a.comprasMesU }
$wsD.Range("Q2:S$ultFila").Value2 = $arrF
# formulas
$wsD.Range("L2:L$ultFila").Formula = '=IF(K2>0,I2/K2,0)'
$wsD.Range("M2:M$ultFila").Formula = '=IF(L2>0,H2/L2,"")'
$wsD.Range("N2:N$ultFila").Formula = '=H2/F2'
$wsD.Range("O2:O$ultFila").Formula = '=H2*G2'
$wsD.Range("P2:P$ultFila").Formula = '=L2*G2'
$wsD.Range("T2:T$ultFila").Formula = '=IF(R2="","",Fecha_corte-R2)'
$wsD.Range("U2:U$ultFila").Formula = '=IF(E2="A",Seg_A,IF(E2="B",Seg_B,IF(E2="C",Seg_C,0)))'
$wsD.Range("V2:V$ultFila").Formula = '=IFERROR(VLOOKUP(C2,t_Plazos,2,FALSE),Plazo_def)'
$wsD.Range("W2:W$ultFila").Formula = '=L2*(U2+V2)'
$wsD.Range("X2:X$ultFila").Formula = '=MAX(0,W2-H2)'
$wsD.Range("Y2:Y$ultFila").Formula = '=IF(X2>0,ROUNDUP(X2/F2,0),0)'
$wsD.Range("Z2:Z$ultFila").Formula = '=Y2*F2*G2'
$wsD.Range("AA2:AA$ultFila").Formula = '=IF(Q2="","",Fecha_corte-Q2)'
$wsD.Range("AA2:AA$ultFila").NumberFormatLocal = '#.##0'
$wsD.Range("F2:F$ultFila").NumberFormatLocal = $UNID; $wsD.Range("G2:G$ultFila").NumberFormatLocal = '$#.##0,00'
$wsD.Range("H2:I$ultFila").NumberFormatLocal = $UNID; $wsD.Range("J2:J$ultFila").NumberFormatLocal = $MONEY
$wsD.Range("K2:K$ultFila").NumberFormatLocal = $UNID; $wsD.Range("L2:L$ultFila").NumberFormatLocal = $DEC1; $wsD.Range("M2:M$ultFila").NumberFormatLocal = $DEC1
$wsD.Range("N2:N$ultFila").NumberFormatLocal = $DEC1; $wsD.Range("O2:P$ultFila").NumberFormatLocal = $MONEY
$wsD.Range("Q2:R$ultFila").NumberFormat = "dd/mm/yyyy"; $wsD.Range("S2:S$ultFila").NumberFormatLocal = $UNID; $wsD.Range("T2:V$ultFila").NumberFormatLocal = $UNID
$wsD.Range("W2:X$ultFila").NumberFormatLocal = $UNID; $wsD.Range("Y2:Y$ultFila").NumberFormatLocal = $UNID; $wsD.Range("Z2:Z$ultFila").NumberFormatLocal = $MONEY
Ancho $wsD @(12,46,28,22,10,9,12,11,11,15,10,11,10,10,15,15,12,12,12,10,10,9,11,11,10,15,11)
[void]$wsD.Range("A1:AA$ultFila").AutoFilter(); Congelar $wsD "C2"
$nomCol = @{ d_Cod="A"; d_Prod="B"; d_Prov="C"; d_Cat="D"; d_Clase="E"; d_UB="F"; d_Costo="G"; d_StockU="H"; d_VentaU="I"; d_VentaN="J"; d_VDu="L"; d_DiasStock="M"; d_StockB="N"; d_StockV="O"; d_VDcosto="P"; d_UltVenta="Q"; d_UltCompra="R"; d_ComprasMes="S"; d_DiasSinComprar="T"; d_Seg="U"; d_Plazo="V"; d_Objetivo="W"; d_APedirU="X"; d_BultosPedir="Y"; d_Importe="Z"; d_DiasSinVender="AA" }
foreach ($k in $nomCol.Keys) { $col = $nomCol[$k]; Nombre $k "=Datos!`$${col}`$2:`$${col}`$$ultFila" }

# colores condicionales de dias de stock (se usa en varias hojas)
function CF-Dias($rango) {
    # reglas sobre el valor de la celda (no dependen del idioma de Excel). La primera ignora las celdas vacias
    # (un texto vacio se compara como "mayor" que cualquier numero) y corta la evaluacion.
    $rango.FormatConditions.Delete() | Out-Null
    $c0 = $rango.FormatConditions.Add(1, 3, '=""'); $c0.StopIfTrue = $true
    $c1 = $rango.FormatConditions.Add(1, 5, "=Dias_muy_lento"); $c1.Interior.Color = $MUYLENTO; $c1.Font.Bold = $true; $c1.StopIfTrue = $true
    $c2 = $rango.FormatConditions.Add(1, 5, "=Dias_alto"); $c2.Interior.Color = $NARANJA; $c2.StopIfTrue = $true
    $c3 = $rango.FormatConditions.Add(1, 6, "=Plazo_def"); $c3.Interior.Color = $ROJO
}

# ---------- Días de stock por proveedor ----------
$wsA = $wb.Worksheets.Item("Días de stock")
Titulo $wsA "Días de stock por proveedor" ("Venta diaria = promedio de los últimos 3 meses por día hábil con operación (" + $d.diasTotal + " días). Stock: depósito principal al " + $hoy.ToString("d/M/yyyy") + ". Colores: rojo = quedan menos días que el plazo de reposición (puede faltar); amarillo = stock alto (más de 60 días); naranja = venta muy lenta (más de 120).")
$h = @("Proveedor","Artículos","Venta 3 meses (`$ neto)","Venta diaria a costo (`$)","Stock (u)","Stock (bultos)","Stock a costo (`$)","DÍAS DE STOCK (valorizado a costo)","Venta diaria (u)","Días de stock (por unidades)")
for ($c = 0; $c -lt $h.Count; $c++) { $wsA.Cells.Item(4, $c + 1) = $h[$c] }
Encabezado $wsA.Range("A4:J4"); $wsA.Rows.Item(4).RowHeight = 42
$wsA.Cells.Item(3,1) = "QUIEBRES (venden y hoy sin stock)"; $wsA.Range("B3").Formula = '=COUNTIFS(d_VentaU,">0",d_StockU,"<=0")'
$wsA.Cells.Item(3,3) = "Su venta de 3 meses (`$)"; $wsA.Range("D3").Formula = '=SUMIFS(d_VentaN,d_VentaU,">0",d_StockU,"<=0")'
$wsA.Cells.Item(3,5) = "% de la venta"; $wsA.Range("F3").Formula = '=IF(SUM(d_VentaN)>0,D3/SUM(d_VentaN),0)'
$wsA.Range("B3").NumberFormatLocal = '#.##0'; $wsA.Range("D3").NumberFormatLocal = $MONEY; $wsA.Range("F3").NumberFormatLocal = $PCT
$wsA.Range("A3:F3").Font.Bold = $true; $wsA.Range("A3:F3").Interior.Color = $ROJO
$provOrd = @($art | Group-Object proveedor | ForEach-Object { $s = 0.0; foreach ($x in $_.Group) { $s += $x.stock * $x.costoUnit }; [pscustomobject]@{ p = $_.Name; v = $s } } | Where-Object { $_.v -gt 0 -or $true } | Sort-Object { -$_.v })
$nP = $provOrd.Count
for ($i = 0; $i -lt $nP; $i++) { $wsA.Cells.Item(5 + $i, 1) = Seguro $provOrd[$i].p }
$u = 4 + $nP
$wsA.Range("B5:B$u").Formula = '=COUNTIFS(d_Prov,$A5)'
$wsA.Range("C5:C$u").Formula = '=SUMIFS(d_VentaN,d_Prov,$A5)'
$wsA.Range("D5:D$u").Formula = '=SUMIFS(d_VDcosto,d_Prov,$A5)'
$wsA.Range("E5:E$u").Formula = '=SUMIFS(d_StockU,d_Prov,$A5)'
$wsA.Range("F5:F$u").Formula = '=SUMIFS(d_StockB,d_Prov,$A5)'
$wsA.Range("G5:G$u").Formula = '=SUMIFS(d_StockV,d_Prov,$A5)'
$wsA.Range("H5:H$u").Formula = '=IF(D5>0,G5/D5,"")'
$wsA.Range("I5:I$u").Formula = '=SUMIFS(d_VDu,d_Prov,$A5)'
$wsA.Range("J5:J$u").Formula = '=IF(I5>0,E5/I5,"")'
$t = $u + 1
$wsA.Cells.Item($t,1) = "TOTAL"
foreach ($c in @("B","C","D","E","F","G","I")) { $wsA.Range("$c$t").Formula = "=SUM(${c}5:${c}$u)" }
$wsA.Range("H$t").Formula = "=IF(D$t>0,G$t/D$t,"""")"
$wsA.Range("A${t}:J$t").Font.Bold = $true; $wsA.Range("A${t}:J$t").Interior.Color = $GRIS
$wsA.Range("B5:B$t").NumberFormatLocal = $UNID; $wsA.Range("C5:D$t").NumberFormatLocal = $MONEY; $wsA.Range("E5:F$t").NumberFormatLocal = $UNID
$wsA.Range("G5:G$t").NumberFormatLocal = $MONEY; $wsA.Range("H5:H$t").NumberFormatLocal = $DEC1; $wsA.Range("I5:I$t").NumberFormatLocal = $UNID; $wsA.Range("J5:J$t").NumberFormatLocal = $DEC1
$wsA.Range("H5:H$t").Font.Bold = $true
CF-Dias $wsA.Range("H5:H$u"); CF-Dias $wsA.Range("J5:J$u")
$wsA.Cells.Item($t + 2, 1) = "Los días en unidades suman productos de distinto tamaño (pañales con jabones); el valorizado a costo es el más representativo. En los productos que se pesan las unidades son gramos."
$wsA.Cells.Item($t + 2, 1).Font.Italic = $true
Ancho $wsA @(42,10,22,18,13,13,18,20,13,16); Congelar $wsA "B5"

# ---------- Por categoría ----------
$wsC = $wb.Worksheets.Item("Por categoría")
Titulo $wsC "Días de stock por proveedor y categoría" "Cada proveedor abierto por las categorías de Gescom (papel higiénico, toallas femeninas, pañales...). Se puede filtrar por proveedor."
$h = @("Proveedor","Categoría","Artículos","Venta 3 meses (`$ neto)","Venta diaria a costo (`$)","Stock (u)","Stock (bultos)","Stock a costo (`$)","DÍAS DE STOCK (valorizado a costo)","Venta diaria (u)","Días de stock (por unidades)")
for ($c = 0; $c -lt $h.Count; $c++) { $wsC.Cells.Item(4, $c + 1) = $h[$c] }
Encabezado $wsC.Range("A4:K4"); $wsC.Rows.Item(4).RowHeight = 42
$pares = @($art | Group-Object proveedor, categoria | ForEach-Object { $s = 0.0; foreach ($x in $_.Group) { $s += $x.stock * $x.costoUnit }; [pscustomobject]@{ p = $_.Group[0].proveedor; c = $_.Group[0].categoria; v = $s } })
$posProv = @{}; for ($i = 0; $i -lt $nP; $i++) { $posProv[$provOrd[$i].p] = $i }
$pares = @($pares | Sort-Object @{ e = { $posProv[$_.p] } }, @{ e = { -$_.v } })
$nC = $pares.Count
$arrC = New-Object 'object[,]' $nC, 2
for ($i = 0; $i -lt $nC; $i++) { $arrC[$i,0] = Seguro $pares[$i].p; $arrC[$i,1] = Seguro $pares[$i].c }
$wsC.Range($wsC.Cells.Item(5, 1), $wsC.Cells.Item(4 + $nC, 2)).Value2 = $arrC
$uc = 4 + $nC
$wsC.Range("C5:C$uc").Formula = '=COUNTIFS(d_Prov,$A5,d_Cat,$B5)'
$wsC.Range("D5:D$uc").Formula = '=SUMIFS(d_VentaN,d_Prov,$A5,d_Cat,$B5)'
$wsC.Range("E5:E$uc").Formula = '=SUMIFS(d_VDcosto,d_Prov,$A5,d_Cat,$B5)'
$wsC.Range("F5:F$uc").Formula = '=SUMIFS(d_StockU,d_Prov,$A5,d_Cat,$B5)'
$wsC.Range("G5:G$uc").Formula = '=SUMIFS(d_StockB,d_Prov,$A5,d_Cat,$B5)'
$wsC.Range("H5:H$uc").Formula = '=SUMIFS(d_StockV,d_Prov,$A5,d_Cat,$B5)'
$wsC.Range("I5:I$uc").Formula = '=IF(E5>0,H5/E5,"")'
$wsC.Range("J5:J$uc").Formula = '=SUMIFS(d_VDu,d_Prov,$A5,d_Cat,$B5)'
$wsC.Range("K5:K$uc").Formula = '=IF(J5>0,F5/J5,"")'
$wsC.Range("C5:C$uc").NumberFormatLocal = $UNID; $wsC.Range("D5:E$uc").NumberFormatLocal = $MONEY; $wsC.Range("F5:G$uc").NumberFormatLocal = $UNID
$wsC.Range("H5:H$uc").NumberFormatLocal = $MONEY; $wsC.Range("I5:I$uc").NumberFormatLocal = $DEC1; $wsC.Range("J5:J$uc").NumberFormatLocal = $UNID; $wsC.Range("K5:K$uc").NumberFormatLocal = $DEC1
$wsC.Range("I5:I$uc").Font.Bold = $true
CF-Dias $wsC.Range("I5:I$uc"); CF-Dias $wsC.Range("K5:K$uc")
[void]$wsC.Range("A4:K$uc").AutoFilter()
Ancho $wsC @(32,26,10,18,18,13,13,18,20,13,16); Congelar $wsC "C5"

# ---------- listas por articulo (INDEX sobre Datos) ----------
function Lista($ws, $filaCab, $idx, $cols, $col0 = 1) {
    # $cols: array de @(titulo, formula con $A{r}, formato, ancho). La 1ra columna guarda la posicion del producto en Datos.
    $n = $idx.Count
    $letra = ($ws.Cells.Item(1, $col0).Address($false, $false) -replace '\d+', '')
    for ($c = 0; $c -lt $cols.Count; $c++) { $ws.Cells.Item($filaCab, $col0 + $c) = $cols[$c][0] }
    Encabezado $ws.Range($ws.Cells.Item($filaCab, $col0), $ws.Cells.Item($filaCab, $col0 + $cols.Count - 1)); $ws.Rows.Item($filaCab).RowHeight = 42
    if ($n -eq 0) { return 0 }
    $arr = New-Object 'object[,]' $n, 1
    for ($i = 0; $i -lt $n; $i++) { $arr[$i,0] = $idx[$i] + 1 }   # posicion dentro de Datos
    $f1 = $filaCab + 1; $fN = $filaCab + $n
    $ws.Range($ws.Cells.Item($f1, $col0), $ws.Cells.Item($fN, $col0)).Value2 = $arr
    for ($c = 1; $c -lt $cols.Count; $c++) {
        $rg = $ws.Range($ws.Cells.Item($f1, $col0 + $c), $ws.Cells.Item($fN, $col0 + $c))
        $rg.Formula = $cols[$c][1].Replace('$A{r}', '$' + $letra + [string]$f1).Replace("{r}", [string]$f1)
        if ($cols[$c][2]) { if ($cols[$c][2] -eq 'FECHA') { $rg.NumberFormat = "dd/mm/yyyy" } else { $rg.NumberFormatLocal = $cols[$c][2] } }
    }
    $ws.Range($ws.Cells.Item($f1, $col0), $ws.Cells.Item($fN, $col0)).Font.Color = 8421504
    for ($c = 0; $c -lt $cols.Count; $c++) { $ws.Columns.Item($col0 + $c).ColumnWidth = $cols[$c][3] }
    return $n
}
$ix = { param($rango) '=INDEX(' + $rango + ',$A{r})' }

# ---------- Sin compra del mes ----------
$wsS = $wb.Worksheets.Item("Sin compra del mes")
Titulo $wsS "Productos con venta que no se compraron en el mes" "Productos que vendieron en los últimos 3 meses y no tuvieron ingreso de compra desde el 1/10/2026 (remitos y facturas de compra, sin notas de crédito). Ordenados por proveedor, clase y urgencia."
function Vencida($a) { return ((-not $a.ultimaVenta) -or ([datetime]::ParseExact($a.ultimaVenta.Substring(0,10), "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture) -lt $hoy.AddDays(-30))) }
$idxS = @(0..($N - 1) | Where-Object { $art[$_].venta3mU -gt 0 -and $art[$_].comprasMesU -le 0 } | Sort-Object @{ e = { $art[$_].proveedor } }, @{ e = { [int](Vencida $art[$_]) } }, @{ e = { $orden[$art[$_].clase] } }, @{ e = { $m[$_].dst } })
$colsS = @(
    @("Fila en Datos", "", $null, 8),
    @("Proveedor", (& $ix "d_Prov"), $null, 28), @("Categoría", (& $ix "d_Cat"), $null, 22), @("Código", (& $ix "d_Cod"), $null, 11), @("Producto", (& $ix "d_Prod"), $null, 44),
    @("Clase", (& $ix "d_Clase"), $null, 7), @("Venta diaria (u)", (& $ix "d_VDu"), $DEC1, 11), @("Stock (u)", (& $ix "d_StockU"), $UNID, 11), @("Stock (bultos)", (& $ix "d_StockB"), $DEC1, 10),
    @("Días de stock", (& $ix "d_DiasStock"), $DEC1, 10), @("Última compra", ('=IF(INDEX(d_UltCompra,$A{r})="","sin registro",INDEX(d_UltCompra,$A{r}))'), 'FECHA', 13),
    @("Días sin comprar", (& $ix "d_DiasSinComprar"), $UNID, 10),
    @("Prioridad", ('=IF(INDEX(d_DiasSinVender,$A{r})>Dias_reciente,"Sin venta reciente",IF(J{r}="","Sin venta",IF(J{r}<INDEX(d_Plazo,$A{r}),"Urgente",IF(J{r}<INDEX(d_Plazo,$A{r})+INDEX(d_Seg,$A{r}),"Pedir pronto","Hay stock"))))'), $null, 17),
    @("Días sin vender", (& $ix "d_DiasSinVender"), $UNID, 10)
)
$nS = Lista $wsS 4 $idxS $colsS
if ($nS -gt 0) {
    $uS = 4 + $nS
    CF-Dias $wsS.Range("J5:J$uS")
    $fc = $wsS.Range("M5:M$uS").FormatConditions
    $x1 = $fc.Add(1, 3, '="Urgente"'); $x1.Interior.Color = $ROJO; $x1.Font.Bold = $true
    $x2 = $fc.Add(1, 3, '="Pedir pronto"'); $x2.Interior.Color = $NARANJA
    $x3 = $fc.Add(1, 3, '="Sin venta reciente"'); $x3.Font.Color = 8421504
    [void]$wsS.Range("A4:N$uS").AutoFilter()
}
Congelar $wsS "F5"

# ---------- Sugerido de pedido ----------
$wsG = $wb.Worksheets.Item("Sugerido de pedido")
Titulo $wsG "Sugerido de pedido" "Objetivo = venta diaria x (stock de seguridad + plazo de reposición). A pedir = objetivo - stock, redondeado hacia arriba a bultos. Seguridad: A 7 días, B 5, C 2 (cambiables en Parámetros). Filtrado por defecto: solo lo que hay que pedir y que vendió en los últimos 30 días (sacá el filtro de 'Días sin vender' para ver también los demás)."
$idxG = @(0..($N - 1) | Where-Object { @("A","B","C") -contains $art[$_].clase } | Sort-Object @{ e = { $art[$_].proveedor } }, @{ e = { $orden[$art[$_].clase] } }, @{ e = { $art[$_].producto } })
$colsG = @(
    @("Fila en Datos", "", $null, 8),
    @("Proveedor", (& $ix "d_Prov"), $null, 28), @("Categoría", (& $ix "d_Cat"), $null, 22), @("Código", (& $ix "d_Cod"), $null, 11), @("Producto", (& $ix "d_Prod"), $null, 44),
    @("Clase", (& $ix "d_Clase"), $null, 7), @("Venta diaria (u)", (& $ix "d_VDu"), $DEC1, 11), @("Stock (u)", (& $ix "d_StockU"), $UNID, 11), @("Días de stock", (& $ix "d_DiasStock"), $DEC1, 10),
    @("Seguridad (días)", (& $ix "d_Seg"), $UNID, 10), @("Plazo (días)", (& $ix "d_Plazo"), $UNID, 9), @("Stock objetivo (u)", (& $ix "d_Objetivo"), $UNID, 11),
    @("A pedir (u)", (& $ix "d_APedirU"), $UNID, 10), @("Unid. por bulto", (& $ix "d_UB"), $UNID, 9), @("BULTOS A PEDIR", (& $ix "d_BultosPedir"), $UNID, 11),
    @("Unidades a pedir", '=INDEX(d_BultosPedir,$A{r})*INDEX(d_UB,$A{r})', $UNID, 11), @("Costo unit. s/IVA", (& $ix "d_Costo"), '$#.##0,00', 12), @("IMPORTE A PEDIR a costo (`$)", (& $ix "d_Importe"), $MONEY, 16),
    @("Compras desde 1/10 (u)", (& $ix "d_ComprasMes"), $UNID, 11),
    @("Última venta", ('=IF(INDEX(d_UltVenta,$A{r})="","sin ventas",INDEX(d_UltVenta,$A{r}))'), 'FECHA', 12), @("Días sin vender", (& $ix "d_DiasSinVender"), $UNID, 10)
)
$nG = Lista $wsG 4 $idxG $colsG
if ($nG -gt 0) {
    $uG = 4 + $nG
    CF-Dias $wsG.Range("I5:I$uG")
    $wsG.Range("O5:O$uG").Font.Bold = $true; $wsG.Range("R5:R$uG").Font.Bold = $true
    # resumen por proveedor a la derecha
    $wsG.Cells.Item(3,23) = "Resumen del pedido (solo productos con venta reciente)"; $wsG.Cells.Item(3,23).Font.Bold = $true
    $wsG.Cells.Item(4,23) = "Proveedor"; $wsG.Cells.Item(4,24) = "Importe a pedir (`$)"; $wsG.Cells.Item(4,25) = "Líneas a pedir"
    Encabezado $wsG.Range("W4:Y4")
    for ($i = 0; $i -lt $prov.Count; $i++) { $wsG.Cells.Item(5 + $i, 23) = Seguro $prov[$i] }
    $uR = 4 + $prov.Count
    $wsG.Range("X5:X$uR").Formula = '=SUMIFS(d_Importe,d_Prov,$W5,d_DiasSinVender,"<="&Dias_reciente)'
    $wsG.Range("Y5:Y$uR").Formula = '=COUNTIFS(d_Prov,$W5,d_BultosPedir,">0",d_DiasSinVender,"<="&Dias_reciente)'
    $wsG.Cells.Item($uR + 1, 23) = "TOTAL"; $wsG.Range("X$($uR+1)").Formula = "=SUM(X5:X$uR)"; $wsG.Range("Y$($uR+1)").Formula = "=SUM(Y5:Y$uR)"
    $wsG.Range("W$($uR+1):Y$($uR+1)").Font.Bold = $true; $wsG.Range("W$($uR+1):Y$($uR+1)").Interior.Color = $GRIS
    $wsG.Range("X5:X$($uR+1)").NumberFormatLocal = $MONEY; $wsG.Range("Y5:Y$($uR+1)").NumberFormatLocal = $UNID
    $wsG.Columns.Item(22).ColumnWidth = 3; $wsG.Columns.Item(23).ColumnWidth = 32; $wsG.Columns.Item(24).ColumnWidth = 18; $wsG.Columns.Item(25).ColumnWidth = 12
    [void]$wsG.Range("A4:U$uG").AutoFilter()
    [void]$wsG.Range("A4:U$uG").AutoFilter(15, ">0")
    [void]$wsG.Range("A4:U$uG").AutoFilter(21, "<=30")
}
Congelar $wsG "F5"

# ---------- Stock valorizado ----------
$wsV = $wb.Worksheets.Item("Stock valorizado")
Titulo $wsV "Stock valorizado a precio de costo sin impuestos" ("Costo del bulto sobre unidades del bulto, según Gescom. Stock del depósito principal al " + $hoy.ToString("d/M/yyyy") + ".")
$h = @("Proveedor","Stock (u)","Stock (bultos)","Valor a costo (`$)","% del total","% acumulado")
for ($c = 0; $c -lt $h.Count; $c++) { $wsV.Cells.Item(4, $c + 1) = $h[$c] }
Encabezado $wsV.Range("A4:F4"); $wsV.Rows.Item(4).RowHeight = 32
for ($i = 0; $i -lt $nP; $i++) { $wsV.Cells.Item(5 + $i, 1) = Seguro $provOrd[$i].p }
$uv = 4 + $nP
$wsV.Range("B5:B$uv").Formula = '=SUMIFS(d_StockU,d_Prov,$A5)'
$wsV.Range("C5:C$uv").Formula = '=SUMIFS(d_StockB,d_Prov,$A5)'
$wsV.Range("D5:D$uv").Formula = '=SUMIFS(d_StockV,d_Prov,$A5)'
$tv = $uv + 1
$wsV.Cells.Item($tv, 1) = "TOTAL"
foreach ($c in @("B","C","D")) { $wsV.Range("$c$tv").Formula = "=SUM(${c}5:${c}$uv)" }
$wsV.Range("E5:E$uv").Formula = "=IF(`$D`$$tv>0,D5/`$D`$$tv,0)"; $wsV.Range("E$tv").Formula = "=SUM(E5:E$uv)"
$wsV.Range("F5:F$uv").Formula = "=SUM(`$E`$5:E5)"
$wsV.Range("A${tv}:F$tv").Font.Bold = $true; $wsV.Range("A${tv}:F$tv").Interior.Color = $GRIS
$wsV.Range("B5:C$tv").NumberFormatLocal = $UNID; $wsV.Range("D5:D$tv").NumberFormatLocal = $MONEY; $wsV.Range("E5:F$tv").NumberFormatLocal = $PCT
# top 30 productos por valor
$top = @(0..($N - 1) | Sort-Object { -$m[$_].val } | Select-Object -First 30)
$wsV.Cells.Item(3, 8) = "Los 30 productos con más stock valorizado"; $wsV.Cells.Item(3, 8).Font.Bold = $true; $wsV.Cells.Item(3, 8).Font.Size = 12
$colsT = @(
    @("Fila en Datos", "", $null, 8),
    @("Proveedor", (& $ix "d_Prov"), $null, 28), @("Producto", (& $ix "d_Prod"), $null, 44), @("Stock (u)", (& $ix "d_StockU"), $UNID, 12), @("Valor a costo (`$)", (& $ix "d_StockV"), $MONEY, 18), @("Días de stock", (& $ix "d_DiasStock"), $DEC1, 12)
)
$nT = Lista $wsV 4 $top $colsT 8
Ancho $wsV @(36,13,13,18,11,12,3)
CF-Dias $wsV.Range("M5:M$(4 + $nT)")
Congelar $wsV "A5"

# ---------- Stock alto - venta lenta ----------
$wsL = $wb.Worksheets.Item("Stock alto - venta lenta")
Titulo $wsL "Stock alto con venta lenta" "Productos con stock cuyos días de stock superan el umbral de Parámetros (60 días = alto, 120 = muy lento) o que no vendieron nada en 3 meses. Ordenados por valor inmovilizado a costo."
$idxL = @(0..($N - 1) | Where-Object { $art[$_].stock -gt 0 -and ($art[$_].venta3mU -le 0 -or $m[$_].dst -gt 60) } | Sort-Object { -$m[$_].val })
$colsL = @(
    @("Fila en Datos", "", $null, 8),
    @("Proveedor", (& $ix "d_Prov"), $null, 28), @("Categoría", (& $ix "d_Cat"), $null, 22), @("Código", (& $ix "d_Cod"), $null, 11), @("Producto", (& $ix "d_Prod"), $null, 44),
    @("Clase", (& $ix "d_Clase"), $null, 9), @("Stock (u)", (& $ix "d_StockU"), $UNID, 11), @("Stock (bultos)", (& $ix "d_StockB"), $DEC1, 10), @("VALOR A COSTO (`$)", (& $ix "d_StockV"), $MONEY, 16),
    @("Venta diaria (u)", (& $ix "d_VDu"), $DEC1, 11), @("Días de stock", (& $ix "d_DiasStock"), $DEC1, 10),
    @("Última venta", ('=IF(INDEX(d_UltVenta,$A{r})="","sin ventas",INDEX(d_UltVenta,$A{r}))'), 'FECHA', 13),
    @("Estado", ('=IF(INDEX(d_VentaU,$A{r})<=0,"Sin venta en 3 meses",IF(K{r}>Dias_muy_lento,"Venta muy lenta",IF(K{r}>Dias_alto,"Stock alto","Normal")))'), $null, 22)
)
$nL = Lista $wsL 5 $idxL $colsL
if ($nL -gt 0) {
    $uL = 5 + $nL
    $wsL.Cells.Item(3,1) = "Valor total de esta lista (`$)"; $wsL.Cells.Item(3,1).Font.Bold = $true
    $wsL.Range("I3").Formula = "=SUBTOTAL(9,I6:I$uL)"; $wsL.Range("I3").NumberFormatLocal = $MONEY; $wsL.Range("I3").Font.Bold = $true
    $wsL.Cells.Item(3,10) = "% del stock total"; $wsL.Range("K3").Formula = "=I3/SUM(d_StockV)"; $wsL.Range("K3").NumberFormatLocal = $PCT; $wsL.Range("K3").Font.Bold = $true
    CF-Dias $wsL.Range("K6:K$uL")
    $fc = $wsL.Range("M6:M$uL").FormatConditions
    $y1 = $fc.Add(1, 3, '="Sin venta en 3 meses"'); $y1.Interior.Color = $MUYLENTO; $y1.Font.Bold = $true
    $y2 = $fc.Add(1, 3, '="Venta muy lenta"'); $y2.Interior.Color = $MUYLENTO
    $y3 = $fc.Add(1, 3, '="Stock alto"'); $y3.Interior.Color = $NARANJA
    [void]$wsL.Range("A5:M$uL").AutoFilter()
}
Congelar $wsL "F6"

# ---------- Notas ----------
Titulo $wsN "Cómo se armó este libro" $null
$notas = @(
    "SOLAPAS",
    "Días de stock / Por categoría: stock actual contra la venta diaria de los últimos 3 meses, por proveedor y abierto por categoría de Gescom.",
    "Sin compra del mes: productos que venden y no tuvieron ingreso de compra desde el 1/10. Prioridad: Urgente = quedan menos días de stock que el plazo de reposición.",
    "Sugerido de pedido: objetivo = venta diaria x (seguridad + plazo). Seguridad A 7 días, B 5, C 2. Pedir = objetivo - stock, redondeado a bultos.",
    "Stock valorizado: stock a costo sin impuestos. Stock alto - venta lenta: stock que supera los días de Parámetros o sin venta en 3 meses.",
    "Parámetros: plazos y días de seguridad (celdas azules). Datos: una fila por producto con todos los cálculos; las demás solapas salen de ahí.",
    "",
    "DE DÓNDE SALE CADA NÚMERO",
    "Stock: depósito principal (PRI) de Gescom, al momento del armado. No incluye el depósito de cambios ni mercadería pedida y todavía no recibida (no está en la base).",
    "Venta: facturas y notas de débito menos rechazos (los canjes por vencimiento no cuentan). Del 7/7 al 2/8 sale del export de Gescom; desde el 3/8, de la base compartida.",
    "Días con operación: " + $d.diasTotal + " (" + $d.diasBase + " de la base + " + $d.diasJulio + " de julio). Venta diaria = venta de 3 meses / días con operación.",
    "Productos que se pesan: la base los cuenta en gramos. Julio se convirtió con la proporción de agosto. Los que no tenían agosto se calcularon solo con los " + $d.diasBase + " días de la base.",
    "Costo: costo del bulto / unidades del bulto (sin impuestos). Venta neta: sin IVA. Se excluyen los vendedores 1176 y 43 y los renglones de servicios.",
    "Clase ABC: por proveedor, sobre la venta neta de 3 meses. A = hasta el 80% acumulado, B = hasta el 95%, C = el resto. Se calculó al armar el libro (no cambia sola).",
    "Categoría: la 'familia' del producto en Gescom. 'Sin categoría' son productos que no vendieron en 2026 o no la tienen cargada.",
    "",
    "LÍMITES A TENER EN CUENTA",
    "El plazo de reposición es 7 días para todos hasta que se cargue el real de cada proveedor (solapa Parámetros).",
    "Si ya hay un pedido hecho que todavía no llegó, el sugerido no lo descuenta: restalo a mano o cargá el ingreso apenas llegue.",
    "Los pedidos se calculan sobre el promedio de 3 meses: en productos de temporada (fiestas, Mantecol) conviene ajustar a mano.",
    "El historial de compras de la base empieza el 17/5/2026: 'sin registro' en Última compra significa que no hay ingreso cargado desde esa fecha, no que nunca se compró.",
    "Productos que dejaron de venderse: si no vendieron en los últimos 30 días (Parámetros) figuran 'Sin venta reciente' y el Sugerido de pedido los oculta por defecto.",
    "Quiebres: productos con venta en los 3 meses y stock cero hoy (cuadro rojo arriba de 'Días de stock'). Su venta diaria es la que tenían antes de quedarse sin stock."
)
for ($i = 0; $i -lt $notas.Count; $i++) {
    $c = $wsN.Cells.Item(3 + $i, 1); $c.Value2 = $notas[$i]; $c.WrapText = $true
    if ($notas[$i] -in @("SOLAPAS","DE DÓNDE SALE CADA NÚMERO","LÍMITES A TENER EN CUENTA")) { $c.Font.Bold = $true; $c.Interior.Color = $GRIS }
}
$wsN.Columns.Item(1).ColumnWidth = 150

# ---------- cierre ----------
$excel.ScreenUpdating = $true
$excel.Calculation = -4105
$excel.Calculate()
$wb.Worksheets.Item("Días de stock").Activate()
$wb.Worksheets.Item("Días de stock").Range("A1").Select()
if (Test-Path $salida) { Remove-Item $salida -Force }
$wb.SaveAs($salida, 51)
Write-Host ("Guardado: " + $salida)
Write-Host ("Articulos en Datos: $N | proveedores: $nP | pares proveedor-categoria: $nC | sin compra: $nS | sugerido: $nG | stock alto: $nL")
$wb.Close($true); $excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
