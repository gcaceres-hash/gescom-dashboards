<#
  Plata en juego por vencimientos (Resumen del Deposito).

  Valoriza al costo de Gescom lo que Digip tiene vencido y lo que esta por vencer, y agrega
  el bloque "perdidas" al data.json del Resumen del Deposito (docs/digip-resumen), que tiene
  que estar DESCIFRADO. No toca las tarjetas (bien / atencion / vigilar).

  - Unidades: las del export de Digip (scripts/digip/extraer-vencimientos.js).
  - Costo: costo_bulto / unidades_bulto, sin IVA. Los productos por peso (unidades_bulto 1000)
    Digip los cuenta por horma: costo por horma = costo del kg x peso promedio por horma del stock
    de Gescom (stock.cantidad / stock.piezas, deposito PRI).
  - Dias: los que le quedan a cada lote para poder despacharse segun Digip.
  - En riesgo: lo que no llega a venderse antes de esa fecha al ritmo de los ultimos 30 dias
    (VEN/DEB sin los vendedores 1176/43/16/37).
  - Liquidaciones: ventas con el codigo 16 (Vendedor Deposito) facturadas a menos del 80% del
    costo (ventas de $50K o mas al costo), con la parte de su saldo en cuenta corriente que les toca. Asi salieron los quesos vencidos de septiembre.

  Uso:
    .\pull-vencidos-valor.ps1 -ListasPath $env:TEMP\digip-vencimientos.json
#>
param(
    [Parameter(Mandatory = $true)][string]$ListasPath,
    [string]$ResumenPath = (Join-Path $PSScriptRoot "..\docs\digip-resumen\data.json"),
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$LiquidacionesDesde = "2026-08-01"
)
$ErrorActionPreference = "Stop"
function Write-Log($m) { Write-Host ("{0}  {1}" -f (Get-Date -Format "HH:mm:ss"), $m) }
$utf8 = New-Object System.Text.UTF8Encoding $false
$config = [IO.File]::ReadAllText($ConfigPath, $utf8) | ConvertFrom-Json
function Invoke-PanelSql([string]$Sql) {
    $uri = "$($config.baseUrl)/consulta?sql=" + [uri]::EscapeDataString($Sql)
    $r = Invoke-RestMethod -Uri $uri -Headers @{ Authorization = "Bearer $($config.clave)" } -Method Get
    if ($r.truncado) { Write-Log "AVISO: la consulta vino truncada -- revisar." }
    return $r.filas
}
$inv = [Globalization.CultureInfo]::InvariantCulture
$i_ = [string][char]0xED; $a_ = [string][char]0xE1; $u_ = [string][char]0xFA; $o_ = [string][char]0xF3
$dias_ = "d" + $i_ + "as"

if (-not (Test-Path $ResumenPath)) { throw "No existe $ResumenPath : descifrar digip-resumen antes (protect-tableros.ps1 -Modo Descifrar -Tableros digip-resumen)." }
$L = [IO.File]::ReadAllText($ListasPath, $utf8) | ConvertFrom-Json
if (-not $L.p) { throw "El archivo de Digip no trae proximos a vencer: volver a correr extraer-vencimientos.js." }
Write-Log ("Digip: {0} lotes vencidos, {1} por vencer (bajado {2})" -f @($L.v).Count, @($L.p).Count, $L.generado)

# --- costos, peso por horma y venta de 30 dias (Gescom) ---
$ART = @{}; foreach ($x in (Invoke-PanelSql "SELECT codigo, unidades_bulto ub, costo_bulto cb FROM articulos")) { $ART[[string]$x.codigo] = $x }
$STK = @{}; foreach ($x in (Invoke-PanelSql "SELECT articulo, SUM(cantidad) cant, SUM(piezas) pz FROM stock WHERE deposito='PRI' GROUP BY articulo")) { $STK[[string]$x.articulo] = $x }
$desde30 = (Get-Date).Date.AddDays(-30).ToString("yyyy-MM-dd")
$VTA = @{}; foreach ($x in (Invoke-PanelSql "SELECT i.articulo art, SUM(i.cantidad) q FROM venta_items i JOIN ventas v ON v.id=i.venta_id WHERE v.tipo IN ('VEN','DEB') AND v.vendedor NOT IN ('1176','43','16','37') AND v.fecha >= '$desde30' GROUP BY i.articulo")) { $VTA[[string]$x.art] = [double]$x.q }

function Get-CostoUnidad([string]$cod) {
    $a = $ART[$cod]; if (-not $a) { return $null }
    $ub = [double]$a.ub; if ($ub -le 0) { $ub = 1 }
    if ($ub -eq 1000) {
        $st = $STK[$cod]
        if (-not $st -or -not $st.pz -or [double]$st.pz -le 0) { return $null }
        $kg = [double]$st.cant / [double]$st.pz / 1000
        return @{ u = [double]$a.cb * $kg; kgPieza = $kg }
    }
    return @{ u = [double]$a.cb / $ub; kgPieza = $null }
}
$sinCosto = New-Object System.Collections.ArrayList
$ven = foreach ($r in $L.v) {
    $c = Get-CostoUnidad $r[0]; if (-not $c) { [void]$sinCosto.Add("$($r[0]) $($r[1])") }
    [pscustomobject]@{ cod = $r[0]; desc = $r[1]; venc = [datetime]::ParseExact($r[2], 'dd/MM/yyyy', $inv); unid = [double]$r[4]; valor = if ($c) { $c.u * [double]$r[4] } else { 0 } }
}
$pv = foreach ($r in $L.p) {
    $c = Get-CostoUnidad $r[0]; if (-not $c) { [void]$sinCosto.Add("$($r[0]) $($r[1])") }
    $u = [double]$r[4]; $dias = [int]$r[3]
    $vendido = $VTA[$r[0]]; if ($vendido -and $c -and $c.kgPieza) { $vendido = $vendido / 1000 / $c.kgPieza }   # gramos -> hormas
    $diaria = if ($vendido) { $vendido / 30 } else { 0 }
    $sobra = [math]::Max(0, $u - $diaria * [math]::Max(0, $dias))
    [pscustomobject]@{ cod = $r[0]; desc = $r[1]; venc = [datetime]::ParseExact($r[2], 'dd/MM/yyyy', $inv); dias = $dias; unid = $u
        venta30 = if ($vendido) { $vendido } else { 0 }; sobra = $sobra
        valor = if ($c) { $c.u * $u } else { 0 }; riesgo = if ($c) { $c.u * $sobra } else { 0 } }
}
if ($sinCosto.Count) { Write-Log ("AVISO sin costo en Gescom (quedan en 0): " + (($sinCosto | Select-Object -Unique) -join '; ')) }

function Sum($rows, $campo) { $s = ($rows | Measure-Object $campo -Sum).Sum; if ($s) { [double]$s } else { 0.0 } }
function Por-Articulo($rows, $campo) {
    $rows | Group-Object cod | ForEach-Object {
        $g = $_.Group
        [pscustomobject]@{ descripcion = $g[0].desc; vence = (($g | Sort-Object venc)[0].venc).ToString('dd/MM/yyyy')
            dias = if ($g[0].PSObject.Properties['dias']) { ($g | Measure-Object dias -Minimum).Minimum } else { $null }
            unidades = Sum $g 'unid'; venta30 = if ($g[0].PSObject.Properties['venta30']) { [math]::Round($g[0].venta30) } else { $null }
            sinVender = if ($g[0].PSObject.Properties['sobra']) { [math]::Round((Sum $g 'sobra')) } else { $null }
            costo = [math]::Round((Sum $g $campo)) }
    } | Sort-Object costo -Descending
}
# Variantes de un mismo producto (sabor o tamano) van en una sola fila: la clave son las 3 primeras
# palabras de la descripcion sin sabores ni medidas (los 4 Tetratop Proteina -> TETRATOP PROTEINA ILOLAY).
$SABORES = 'FRUTILLA','FRUT','VAINILLA','VAIN','CHOCOLATE','CHOCO','DDL','MANZ','MANZANA','UVA','CEREZA','DURAZNO','ARANDANOS','BAN','BANANA','LIMON','KW','KIWI','NARANJA'
$MEDIDAS = 'X','GR','GRS','G','ML','CC','KG','L','CM3','U'
function Get-Linea([string]$desc) {
    $t = ($desc.ToUpper() -replace '[\.\/,\(\)]', ' ').Split(' ', [StringSplitOptions]::RemoveEmptyEntries) |
        Where-Object { $_ -notmatch '\d' -and $SABORES -notcontains $_ -and $MEDIDAS -notcontains $_ }
    return (@($t) | Select-Object -First 3) -join ' '
}
function Por-Linea($rows, $campo) {
    $rows | Group-Object { Get-Linea $_.desc } | ForEach-Object {
        $g = $_.Group; $arts = @($g | Group-Object cod)
        $nombre = if ($arts.Count -gt 1) { "$($_.Name) ($($arts.Count) variantes)" } else { $g[0].desc }
        $detalle = if ($arts.Count -gt 1) { ($arts | Sort-Object { -(Sum $_.Group 'unid') } | ForEach-Object { "{0} ({1:N0} u)" -f $_.Group[0].desc, (Sum $_.Group 'unid') }) -join ' · ' } else { $null }
        [pscustomobject]@{ descripcion = $nombre; detalle = $detalle; vence = (($g | Sort-Object venc)[0].venc).ToString('dd/MM/yyyy')
            dias = ($g | Measure-Object dias -Minimum).Minimum; unidades = Sum $g 'unid'
            venta30 = [math]::Round((($arts | ForEach-Object { $_.Group[0].venta30 }) | Measure-Object -Sum).Sum)
            sinVender = [math]::Round((Sum $g 'sobra')); costo = [math]::Round((Sum $g $campo)) }
    } | Sort-Object costo -Descending
}
$tramos = foreach ($b in @(@{ n = "0 a 15 $dias_"; a = -9999; z = 15 }, @{ n = "16 a 30 $dias_"; a = 16; z = 30 }, @{ n = "M$($a_)s de 30 $dias_"; a = 31; z = 99999 })) {
    $x = @($pv | Where-Object { $_.dias -ge $b.a -and $_.dias -le $b.z })
    [ordered]@{ tramo = $b.n; unidades = Sum $x 'unid'; costo = [math]::Round((Sum $x 'valor')); enRiesgo = [math]::Round((Sum $x 'riesgo')) }
}
$vencCosto = [math]::Round((Sum $ven 'valor'))
$perd = [ordered]@{
    generado      = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    vencido       = [ordered]@{ unidades = Sum $ven 'unid'; lotes = @($ven).Count; costo = $vencCosto }
    porVencer     = @($tramos)
    totalEnRiesgo = $vencCosto + (($tramos | ForEach-Object { $_.enRiesgo }) | Measure-Object -Sum).Sum
    topVencidos   = @(Por-Articulo $ven 'valor' | Select-Object -First 8 | ForEach-Object { [ordered]@{ descripcion = $_.descripcion; vence = $_.vence; unidades = $_.unidades; costo = $_.costo } })
    topRiesgo     = @(Por-Linea @($pv | Where-Object { $_.riesgo -gt 0 }) 'riesgo' | Select-Object -First 10 | ForEach-Object { [ordered]@{ descripcion = $_.descripcion; detalle = $_.detalle; vence = $_.vence; dias = $_.dias; unidades = $_.unidades; venta30 = $_.venta30; sinVender = $_.sinVender; costo = $_.costo } })
    criterio      = "Valor al costo sin IVA de Gescom. Los $dias_ son los que le quedan a cada lote para poder despacharse seg$($u_)n Digip. 'En riesgo' = lo que no llega a venderse antes de esa fecha al ritmo de venta de los $($u_)ltimos 30 $dias_. Quesos y fiambres por peso: Digip cuenta hormas; se valorizan con el peso promedio por horma del stock de Gescom. Liquidaciones: ventas con el c$($o_)digo 16 facturadas a menos del 80% del costo actual."
}

# --- liquidaciones: codigo 16 por debajo del 80% del costo ---
$liq = Invoke-PanelSql @"
SELECT * FROM (
  SELECT v.id, v.fecha, v.neto vneto, v.cliente, c.razon_social cli, i.articulo art, a.descripcion des, a.unidades_bulto ub, i.cantidad q, i.neto nt,
         i.cantidad * a.costo_bulto / CASE WHEN a.unidades_bulto > 0 THEN a.unidades_bulto ELSE 1 END cto
  FROM venta_items i JOIN ventas v ON v.id = i.venta_id JOIN articulos a ON a.codigo = i.articulo LEFT JOIN clientes c ON c.codigo = v.cliente
  WHERE v.vendedor = '16' AND v.tipo IN ('VEN','DEB') AND v.fecha >= '$LiquidacionesDesde' AND i.neto > 0
) WHERE nt < 0.8 * cto
"@
$ventasLiq = @()
if ($liq) {
    $ids = (@($liq) | ForEach-Object { $_.id } | Select-Object -Unique) -join ','
    $SAL = @{}; foreach ($x in (Invoke-PanelSql "SELECT venta_id, SUM(saldo) s FROM ctacte_clientes WHERE venta_id IN ($ids) GROUP BY venta_id")) { $SAL[[string]$x.venta_id] = [double]$x.s }
    $ventasLiq = @($liq | Group-Object id | Where-Object { (Sum $_.Group 'cto') -ge 50000 } | ForEach-Object {
        $g = $_.Group; $f = [datetime]::ParseExact(([string]$g[0].fecha).Substring(0, 10), 'yyyy-MM-dd', $inv)
        $det = ($g | Sort-Object cto -Descending | ForEach-Object {
            $q = if ([double]$_.ub -eq 1000) { ("{0:N1} kg" -f ([double]$_.q / 1000)) } else { ("{0:N0}" -f [double]$_.q) }
            $d = (Get-Culture).TextInfo.ToTitleCase(([string]$_.des).ToLower()); "$d ($q)" }) -join ', '
        $cto = Sum $g 'cto'; $nt = Sum $g 'nt'
        [ordered]@{ fecha = $f.ToString('dd/MM'); orden = $f; cliente = "$($g[0].cliente) " + [char]0xB7 + " $($g[0].cli)"; detalle = $det
            costo = [math]::Round($cto); neto = [math]::Round($nt); recupero = [math]::Round($nt / $cto, 2); saldo = [math]::Round([double]$SAL[[string]$_.Name] * [math]::Min([double]1, $nt / [math]::Max([double]1, [double]$g[0].vneto)), 2) }
    } | Sort-Object { $_.orden } | ForEach-Object { $_.Remove('orden'); $_ })
}
$perd.liquidaciones = [ordered]@{
    vendedor       = "16 (Vendedor Dep$($o_)sito)"
    costo          = [math]::Round((($ventasLiq | ForEach-Object { $_.costo }) | Measure-Object -Sum).Sum)
    netoFacturado  = [math]::Round((($ventasLiq | ForEach-Object { $_.neto }) | Measure-Object -Sum).Sum)
    saldoSinCobrar = [math]::Round((($ventasLiq | ForEach-Object { $_.saldo }) | Measure-Object -Sum).Sum, 2)
    ventas         = $ventasLiq
}

$res = [IO.File]::ReadAllText($ResumenPath, $utf8) | ConvertFrom-Json
$res | Add-Member -NotePropertyName perdidas -NotePropertyValue $perd -Force
[IO.File]::WriteAllText($ResumenPath, ($res | ConvertTo-Json -Depth 8), $utf8)

Write-Log ("Vencido: {0:N0} u = `${1:N0} | por vencer: {2} | total en juego `${3:N0}" -f $perd.vencido.unidades, $perd.vencido.costo, (($tramos | ForEach-Object { "$($_.tramo) `$$('{0:N0}' -f $_.enRiesgo) en riesgo" }) -join ' / '), $perd.totalEnRiesgo)
Write-Log ("Liquidaciones codigo 16: {0} ventas, costo `${1:N0}, facturado `${2:N0}, sin cobrar `${3:N0}" -f @($ventasLiq).Count, $perd.liquidaciones.costo, $perd.liquidaciones.netoFacturado, $perd.liquidaciones.saldoSinCobrar)
Write-Log "Guardado: $ResumenPath"
