<#
Inteligencia Comercial de UN proveedor (pedido de Gisela, 28/10/2026; lo llama pull-softys-ic.ps1 para cada
solapa): los 9 indicadores de la tabla "Indicador / Que responde / Frecuencia / Senal de alerta".

  1 Tasa de crecimiento        mensual            ventas planas o cayendo
  2 Cumplimiento de objetivo   semanal / mensual  menos de 90% sostenido
  3 Tasa de conversion         semanal            de los clientes que compran a la empresa, cuantos llevan el proveedor
  4 Ticket promedio            semanal            mas tickets con menos valor
  5 UPT                        semanal            unidades y articulos por ticket (venta cruzada)
  6 Ventas por vendedor        semanal / mensual  brechas sin explicacion
  7 Rendimiento por producto   mensual            mucha venta con poco aporte de margen
  8 Retencion                  mensual            dependencia de clientes nuevos
  9 Devoluciones               semanal / mensual  rechazos + cambios en aumento

Venta = VEN + DEB - devoluciones (DEV-RE / DEV-CA), con IVA (como los objetivos); margen sobre el neto
sin IVA con costo = precio_costo x cantidad (en las ventas por bulto precio_costo ya es el del bulto).
Sin los vendedores 1176, 43, 16 y 37. Objetivo del proveedor: scripts/privado/objetivos (el del Director).
La base empieza el 03/08/2026.
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [Parameter(Mandatory = $true)][string]$OutPath,
    [string]$ObjetivosDir = (Join-Path $PSScriptRoot "privado"),
    [string]$Proveedor = "150",
    [string]$ProveedorObjetivo = "SOFTYS ARGENTINA S.A",
    [string]$Nombre = "Softys",
    [ValidateSet("familia","descripcion")][string]$LineaPor = "familia",   # Georgalos e Ilolay tienen una sola familia: la linea sale de la descripcion
    [int]$CantSemanas = 12
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$u8 = New-Object System.Text.UTF8Encoding $false
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
function Invoke-PanelSql([string]$Sql) {
    $uri = "$($config.baseUrl)/consulta?sql=" + [uri]::EscapeDataString($Sql)
    $r = Invoke-RestMethod -Uri $uri -Headers @{ Authorization = "Bearer $($config.clave)" } -Method Get
    if ($r.truncado) { throw "La consulta vino truncada (mas de 20000 filas): acotarla. SQL: $($Sql.Substring(0,[Math]::Min(120,$Sql.Length)))" }
    return $r.filas
}
function R0($x) { return [math]::Round([double]$x, 0) }
function R2($x) { return [math]::Round([double]$x, 2) }
function R4($x) { return [math]::Round([double]$x, 4) }
function Div($a, $b) { if ($null -eq $b -or [double]$b -eq 0) { return $null } return [double]$a / [double]$b }
function D0($x) { if ($null -eq $x) { 0.0 } else { [double]$x } }

$EXCL = "'1176','43','043','16','016','37','037'"
$FERIADOS = @('2026-08-17','2026-10-12','2026-11-23','2026-12-08','2026-12-25','2027-01-01')
$SIGNO = "CASE WHEN v.tipo IN ('VEN','DEB') THEN 1 WHEN v.tipo LIKE 'DEV%' THEN -1 ELSE 0 END"
$DESDE = "2026-08-01"
$BASE = "FROM venta_items i JOIN ventas v ON v.id = i.venta_id JOIN articulos a ON a.codigo = i.articulo WHERE a.proveedor = '$Proveedor' AND COALESCE(v.vendedor,'') NOT IN ($EXCL)"
$SEMANA = "date(v.fecha,'weekday 0','-6 days')"   # lunes de la semana

function Es-Habil([datetime]$d) { return ($d.DayOfWeek -ne 'Saturday' -and $d.DayOfWeek -ne 'Sunday' -and $FERIADOS -notcontains $d.ToString('yyyy-MM-dd')) }
function Habiles([datetime]$desde, [datetime]$hasta) { $n=0; for ($d=$desde.Date; $d -le $hasta.Date; $d=$d.AddDays(1)) { if (Es-Habil $d) { $n++ } }; return $n }
function F([datetime]$d) { $d.ToString('yyyy-MM-dd') }

# ---------------- periodo ----------------
$corte = [datetime]::ParseExact(((Invoke-PanelSql "SELECT MAX(fecha) AS f FROM ventas WHERE tipo='VEN'")[0].f), 'yyyy-MM-dd', $null)
$ini = (Get-Date -Year $corte.Year -Month $corte.Month -Day 1).Date; $fin = $ini.AddMonths(1).AddDays(-1)
$iniPrev = $ini.AddMonths(-1); $finPrev = $ini.AddDays(-1)
$mes = $ini.ToString('yyyy-MM'); $mesPrev = $iniPrev.ToString('yyyy-MM')
$habTot = Habiles $ini $fin; $habTrans = Habiles $ini $corte; if ($habTrans -lt 1) { $habTrans = 1 }
$k=0; $cortePrev = $finPrev
for ($d=$iniPrev; $d -le $finPrev; $d=$d.AddDays(1)) { if (Es-Habil $d) { $k++; if ($k -eq $habTrans) { $cortePrev = $d; break } } }
Write-Log "$Nombre $mes al $($corte.ToString('dd/MM')) | dias habiles $habTrans/$habTot | comparo con $mesPrev al $($cortePrev.ToString('dd/MM'))"

# ---------------- objetivo ----------------
$objMes = @{}
$objPath = Join-Path $ObjetivosDir "objetivos/data.json"
if (-not (Test-Path $objPath)) { & (Join-Path $PSScriptRoot "protect-tableros.ps1") -Modo Descifrar -DocsDir $ObjetivosDir -Tableros objetivos | Out-Null }
if (Test-Path $objPath) {
    $o = [IO.File]::ReadAllText($objPath, $u8) | ConvertFrom-Json
    $p = $o.proveedores.$ProveedorObjetivo
    if ($p) { $objMes[[string]$o.mes] = @{ eq = [double]$p.eq; ideal = [double]$p.ideal } }
}
if (-not $objMes.ContainsKey($mes)) { Write-Log "AVISO: no hay objetivo de $Nombre para $mes (se muestra sin objetivo)." }
# objetivo diario: el del mes repartido en sus dias habiles (null si algun dia habil del rango no tiene objetivo)
function Obj-Rango([datetime]$a, [datetime]$b, [string]$esc) {
    $s = 0.0; $hay = $false
    for ($d=$a.Date; $d -le $b.Date; $d=$d.AddDays(1)) {
        if (-not (Es-Habil $d)) { continue }
        $m = $d.ToString('yyyy-MM'); if (-not $objMes.ContainsKey($m)) { return $null }   # semana partida sin objetivo en una parte: no se compara
        $i0 = Get-Date -Year $d.Year -Month $d.Month -Day 1; $ht = Habiles $i0 $i0.AddMonths(1).AddDays(-1)
        $s += $objMes[$m][$esc] / $ht; $hay = $true
    }
    if ($hay) { return $s } else { return $null }
}

# ---------------- agregados por periodo ----------------
function Agregado([string]$clave, [string]$filtro) {
    Invoke-PanelSql @"
SELECT $clave AS k,
  COUNT(DISTINCT CASE WHEN v.tipo='VEN' THEN v.id END) AS tickets,
  COUNT(DISTINCT CASE WHEN v.tipo='VEN' THEN v.cliente END) AS clientes,
  SUM(CASE WHEN v.tipo='VEN' THEN i.total ELSE 0 END) AS ventaVen,
  SUM(CASE WHEN v.tipo='VEN' THEN i.cantidad ELSE 0 END) AS unidades,
  SUM(CASE WHEN v.tipo='VEN' THEN 1 ELSE 0 END) AS lineas,
  SUM(CASE WHEN v.tipo='DEV-RE' THEN i.total ELSE 0 END) AS rechazos,
  SUM(CASE WHEN v.tipo='DEV-CA' THEN i.total ELSE 0 END) AS cambios,
  SUM($SIGNO * i.total) AS venta, SUM($SIGNO * i.neto) AS neto, SUM($SIGNO * i.precio_costo * i.cantidad) AS costo
$BASE AND $filtro GROUP BY k
"@
}
function Empresa([string]$clave, [string]$filtro) {
    $h = @{}; foreach ($x in (Invoke-PanelSql "SELECT $clave AS k, COUNT(DISTINCT v.cliente) AS n FROM ventas v WHERE v.tipo='VEN' AND COALESCE(v.vendedor,'') NOT IN ($EXCL) AND $filtro GROUP BY k")) { $h[[string]$x.k] = [double]$x.n }; return $h
}
function Fila($x, $emp, $extra) {
    $t = D0 $x.tickets
    $o = [ordered]@{
        venta = R0 (D0 $x.venta); neto = R0 (D0 $x.neto); margen = R4 (Div ((D0 $x.neto) - (D0 $x.costo)) (D0 $x.neto))
        contribucion = R0 ((D0 $x.neto) - (D0 $x.costo))
        tickets = R0 $t; clientes = R0 (D0 $x.clientes); clientesEmpresa = $(if ($emp) { R0 $emp } else { $null })
        conversion = R4 (Div (D0 $x.clientes) $emp)
        ticketProm = R0 (Div (D0 $x.ventaVen) $t); upt = R2 (Div (D0 $x.unidades) $t); articulosPorTicket = R2 (Div (D0 $x.lineas) $t)
        rechazos = R0 (D0 $x.rechazos); cambios = R0 (D0 $x.cambios)
        devolucionPct = R4 (Div ((D0 $x.rechazos) + (D0 $x.cambios)) (D0 $x.ventaVen))
    }
    if ($extra) { foreach ($e in $extra.GetEnumerator()) { $o[$e.Key] = $e.Value } }
    return $o
}

# meses (desde agosto) + mes anterior a mismos dias
$mesesRaw = Agregado "substr(v.fecha,1,7)" "v.fecha >= '$DESDE' AND v.fecha <= '$(F $corte)'"
$empMes = Empresa "substr(v.fecha,1,7)" "v.fecha >= '$DESDE' AND v.fecha <= '$(F $corte)'"
$prevMtdRaw = @(Agregado "'prev'" "v.fecha >= '$(F $iniPrev)' AND v.fecha <= '$(F $cortePrev)'")[0]
$empPrevMtd = (Empresa "'prev'" "v.fecha >= '$(F $iniPrev)' AND v.fecha <= '$(F $cortePrev)'")['prev']
$meses = @($mesesRaw | Sort-Object k | ForEach-Object {
    $m = [string]$_.k; $i0 = [datetime]::ParseExact("$m-01", 'yyyy-MM-dd', $null); $f0 = $i0.AddMonths(1).AddDays(-1)
    $ht = Habiles $i0 $f0; $hc = if ($m -eq $mes) { $habTrans } else { $ht }
    $ob = if ($objMes.ContainsKey($m)) { $objMes[$m] } else { $null }
    Fila $_ $empMes[$m] ([ordered]@{ mes = $m; completo = ($m -ne $mes); habiles = $ht; habilesTransc = $hc
        proyeccion = R0 ((D0 $_.venta) / $hc * $ht); objEq = $(if ($ob) { R0 $ob.eq } else { $null }); objIdeal = $(if ($ob) { R0 $ob.ideal } else { $null }) })
})
$actual = $meses | Where-Object { $_.mes -eq $mes }
$prevMtd = Fila $prevMtdRaw $empPrevMtd ([ordered]@{ mes = $mesPrev; hasta = F $cortePrev })
$prevMes = $meses | Where-Object { $_.mes -eq $mesPrev }

# semanas (lunes a domingo)
$lunes = $corte.Date.AddDays(-(([int]$corte.DayOfWeek + 6) % 7)); $desdeSem = $lunes.AddDays(-7 * ($CantSemanas - 1))
if ($desdeSem -lt [datetime]'2026-08-03') { $desdeSem = [datetime]'2026-08-03' }
$filtroSem = "v.fecha >= '$(F $desdeSem)' AND v.fecha <= '$(F $corte)'"
$empSem = Empresa $SEMANA $filtroSem
$semanas = @(Agregado $SEMANA $filtroSem | Sort-Object k | ForEach-Object {
    $l = [datetime]::ParseExact([string]$_.k, 'yyyy-MM-dd', $null); $h = $l.AddDays(6); if ($h -gt $corte) { $h = $corte }
    $oe = Obj-Rango $l $h 'eq'; $oi = Obj-Rango $l $h 'ideal'; $hb = Habiles $l $h
    $e = [ordered]@{ desde = F $l; hasta = F $h; completa = ($l.AddDays(6) -le $corte); habiles = $hb
        objEq = $(if ($null -ne $oe) { R0 $oe } else { $null }); objIdeal = $(if ($null -ne $oi) { R0 $oi } else { $null }) }
    Fila $_ $empSem[[string]$_.k] $e
})
Write-Log "Meses: $(($meses | ForEach-Object { $_.mes + ' ' + [math]::Round($_.venta/1e6,1) + 'M' }) -join ' | ') | semanas: $($semanas.Count)"

# ---------------- vendedores ----------------
$nomVend = @{}; foreach ($x in (Invoke-PanelSql "SELECT codigo, nombre FROM vendedores")) { $nomVend[[string]$x.codigo] = [string]$x.nombre }
$ultLunes = $lunes; $semAnt = $lunes.AddDays(-7)
$vraw = Invoke-PanelSql @"
SELECT v.vendedor AS ve,
  SUM(CASE WHEN v.fecha >= '$(F $ini)' THEN $SIGNO * i.total ELSE 0 END) AS mtd,
  SUM(CASE WHEN v.fecha >= '$(F $iniPrev)' AND v.fecha <= '$(F $cortePrev)' THEN $SIGNO * i.total ELSE 0 END) AS prevMtd,
  SUM(CASE WHEN v.fecha >= '$(F $iniPrev)' AND v.fecha <= '$(F $finPrev)' THEN $SIGNO * i.total ELSE 0 END) AS prevTotal,
  SUM(CASE WHEN v.fecha >= '$(F $ini)' THEN $SIGNO * i.neto ELSE 0 END) AS neto,
  SUM(CASE WHEN v.fecha >= '$(F $ini)' THEN $SIGNO * i.precio_costo * i.cantidad ELSE 0 END) AS costo,
  SUM(CASE WHEN v.fecha >= '$(F $ultLunes)' THEN $SIGNO * i.total ELSE 0 END) AS semana,
  SUM(CASE WHEN v.fecha >= '$(F $semAnt)' AND v.fecha < '$(F $ultLunes)' THEN $SIGNO * i.total ELSE 0 END) AS semanaAnt,
  COUNT(DISTINCT CASE WHEN v.fecha >= '$(F $ini)' AND v.tipo='VEN' THEN v.id END) AS tickets,
  COUNT(DISTINCT CASE WHEN v.fecha >= '$(F $ini)' AND v.tipo='VEN' THEN v.cliente END) AS clientes,
  SUM(CASE WHEN v.fecha >= '$(F $ini)' AND v.tipo='VEN' THEN i.total ELSE 0 END) AS ventaVen,
  SUM(CASE WHEN v.fecha >= '$(F $ini)' AND v.tipo='VEN' THEN 1 ELSE 0 END) AS lineas,
  SUM(CASE WHEN v.fecha >= '$(F $ini)' AND v.tipo IN ('DEV-RE','DEV-CA') THEN i.total ELSE 0 END) AS devol
$BASE AND v.fecha >= '$(F $iniPrev)' AND v.fecha <= '$(F $corte)' GROUP BY v.vendedor
"@
$empVend = @{}; foreach ($x in (Invoke-PanelSql "SELECT v.vendedor AS ve, COUNT(DISTINCT v.cliente) AS n FROM ventas v WHERE v.tipo='VEN' AND v.fecha >= '$(F $ini)' AND v.fecha <= '$(F $corte)' AND COALESCE(v.vendedor,'') NOT IN ($EXCL) GROUP BY v.vendedor")) { $empVend[[string]$x.ve] = [double]$x.n }
$vendedores = @($vraw | Where-Object { [math]::Abs((D0 $_.mtd)) + [math]::Abs((D0 $_.prevTotal)) -gt 0 } | ForEach-Object {
    $c = [string]$_.ve; $t = D0 $_.tickets
    [ordered]@{ codigo = $c; nombre = $(if ($nomVend.ContainsKey($c)) { $nomVend[$c] } else { "Vendedor $c" })
        venta = R0 (D0 $_.mtd); ventaPrevMtd = R0 (D0 $_.prevMtd); ventaPrevTotal = R0 (D0 $_.prevTotal)
        var = R4 (Div ((D0 $_.mtd) - (D0 $_.prevMtd)) (D0 $_.prevMtd))
        semana = R0 (D0 $_.semana); semanaAnt = R0 (D0 $_.semanaAnt)
        margen = R4 (Div ((D0 $_.neto) - (D0 $_.costo)) (D0 $_.neto))
        tickets = R0 $t; clientes = R0 (D0 $_.clientes); clientesEmpresa = R0 (D0 $empVend[$c])
        conversion = R4 (Div (D0 $_.clientes) $empVend[$c])
        ticketProm = R0 (Div (D0 $_.ventaVen) $t); articulosPorTicket = R2 (Div (D0 $_.lineas) $t)
        devolucionPct = R4 (Div (D0 $_.devol) (D0 $_.ventaVen)) }
} | Sort-Object { -$_.venta })
$totVend = ($vendedores | ForEach-Object { $_.venta } | Measure-Object -Sum).Sum
foreach ($v in $vendedores) { $v.participacion = R4 (Div $v.venta $totVend) }

# ---------------- productos ----------------
$fam = @{}; foreach ($x in (Invoke-PanelSql "SELECT codigo, descripcion FROM familias")) { $fam[[string]$x.codigo] = [string]$x.descripcion }
$praw = Invoke-PanelSql @"
SELECT i.articulo AS a, a.descripcion AS d, a.familia AS f,
  SUM(CASE WHEN v.fecha >= '$(F $ini)' THEN $SIGNO * i.total ELSE 0 END) AS t,
  SUM(CASE WHEN v.fecha >= '$(F $ini)' THEN $SIGNO * i.neto ELSE 0 END) AS n,
  SUM(CASE WHEN v.fecha >= '$(F $ini)' THEN $SIGNO * i.precio_costo * i.cantidad ELSE 0 END) AS c,
  SUM(CASE WHEN v.fecha >= '$(F $ini)' THEN $SIGNO * i.cantidad ELSE 0 END) AS q,
  SUM(CASE WHEN v.fecha <= '$(F $cortePrev)' THEN $SIGNO * i.total ELSE 0 END) AS tp,
  SUM(CASE WHEN v.fecha <= '$(F $cortePrev)' THEN $SIGNO * i.neto ELSE 0 END) AS np,
  SUM(CASE WHEN v.fecha <= '$(F $cortePrev)' THEN $SIGNO * i.precio_costo * i.cantidad ELSE 0 END) AS cp
$BASE AND ((v.fecha >= '$(F $ini)' AND v.fecha <= '$(F $corte)') OR (v.fecha >= '$(F $iniPrev)' AND v.fecha <= '$(F $cortePrev)')) GROUP BY i.articulo, a.descripcion, a.familia
"@
$NOMBRE_LINEA = @{ 'PAPEL HIG' = 'Papel higiénico'; 'TOALLA HUMEDA' = 'Toalla húmeda'; 'SOFTYS' = 'Otros Softys'; 'AERO POETT' = 'Aerosol Poett'; 'AERO AYUDIN' = 'Aerosol Ayudín'; 'SELTON' = 'Selton'; 'MORTIMER' = 'Mortimer' }
# por descripcion: la primera palabra (YOG. / YOGUR -> Yogur, MANTECOL, TODDY...)
$PALABRA_LINEA = @{ 'YOG' = 'Yogur'; 'YOGUR' = 'Yogur'; 'DULCE' = 'Dulce de leche'; 'DISPLAY' = 'Queso'; 'FLYNN' = 'Flynn Paff'; 'CEREAL' = 'Flynn Paff'; 'TURRÓN' = 'Turrón'; 'TURRON' = 'Turrón'; 'NAMUR' = 'Turrón'; 'P' = 'Pasta de maní'; 'FULL' = 'Full Maní'; 'HP' = 'Full Maní'; 'ALFAJOR' = 'Full Maní'; 'MANÍ' = 'Maní'; 'CARAM' = 'Caramelos'; 'CARAMELO' = 'Caramelos'; 'CARAMELOS' = 'Caramelos'; 'PAN' = 'Pan dulce'; 'GEORGALOS' = 'Pan dulce' }
function Linea($x) {
    if ($LineaPor -eq 'descripcion') { $w = (([string]$x.d).ToUpper() -replace '[\.\/,]', ' ').Split(' ', [StringSplitOptions]::RemoveEmptyEntries)[0]
        if ($PALABRA_LINEA.ContainsKey($w)) { return $PALABRA_LINEA[$w] }; return $w.Substring(0,1) + $w.Substring(1).ToLower() }
    $f = [string]$x.f; if (-not $fam.ContainsKey($f)) { return 'Sin familia' }; $s = $fam[$f]; if ($NOMBRE_LINEA.ContainsKey($s)) { return $NOMBRE_LINEA[$s] }; $s.Substring(0,1).ToUpper() + $s.Substring(1).ToLower() }
$porLinea = @{}
foreach ($x in $praw) {
    $l = Linea $x; if (-not $porLinea.ContainsKey($l)) { $porLinea[$l] = @{ t=0.0; n=0.0; c=0.0; tp=0.0; np=0.0; cp=0.0; skus=0 } }
    $g = $porLinea[$l]; foreach ($c in 't','n','c','tp','np','cp') { $g[$c] += D0 $x.$c }; if ((D0 $x.t) -gt 0) { $g.skus++ }
}
if ($LineaPor -eq 'descripcion') {   # las lineas que salen de la descripcion son muchas: las de menos de 1,5% de la venta van a "Otros"
    $tt = ($porLinea.Values | ForEach-Object { $_.t } | Measure-Object -Sum).Sum
    foreach ($k in @($porLinea.Keys)) { if ($k -ne 'Otros' -and $porLinea[$k].t -lt 0.015 * $tt) {
        if (-not $porLinea.ContainsKey('Otros')) { $porLinea['Otros'] = @{ t=0.0; n=0.0; c=0.0; tp=0.0; np=0.0; cp=0.0; skus=0 } }
        foreach ($c in 't','n','c','tp','np','cp','skus') { $porLinea['Otros'][$c] += $porLinea[$k][$c] }; $porLinea.Remove($k) } }
}
$totT = ($praw | ForEach-Object { D0 $_.t } | Measure-Object -Sum).Sum; $totN = ($praw | ForEach-Object { D0 $_.n } | Measure-Object -Sum).Sum
$totC = ($praw | ForEach-Object { D0 $_.c } | Measure-Object -Sum).Sum; $totContrib = $totN - $totC
$lineas = @($porLinea.GetEnumerator() | Where-Object { $_.Value.t -ne 0 -or $_.Value.tp -ne 0 } | ForEach-Object {
    $g = $_.Value
    [ordered]@{ nombre = $_.Key; skus = $g.skus; venta = R0 $g.t; ventaPrevMtd = R0 $g.tp; var = R4 (Div ($g.t - $g.tp) $g.tp)
        margen = R4 (Div ($g.n - $g.c) $g.n); margenPrevMtd = R4 (Div ($g.np - $g.cp) $g.np)
        contribucion = R0 ($g.n - $g.c); pesoVenta = R4 (Div $g.t $totT); pesoContribucion = R4 (Div ($g.n - $g.c) $totContrib) }
} | Sort-Object { -$_.venta })
$skus = @($praw | Where-Object { (D0 $_.t) -gt 0 } | Sort-Object { -(D0 $_.t) } | Select-Object -First 30 | ForEach-Object {
    [ordered]@{ codigo = [string]$_.a; nombre = [string]$_.d; linea = Linea $_; venta = R0 (D0 $_.t); ventaPrevMtd = R0 (D0 $_.tp)
        var = R4 (Div ((D0 $_.t) - (D0 $_.tp)) (D0 $_.tp)); margen = R4 (Div ((D0 $_.n) - (D0 $_.c)) (D0 $_.n))
        contribucion = R0 ((D0 $_.n) - (D0 $_.c)); pesoVenta = R4 (Div (D0 $_.t) $totT); pesoContribucion = R4 (Div ((D0 $_.n) - (D0 $_.c)) $totContrib) }
})

# ---------------- retencion ----------------
$craw = Invoke-PanelSql "SELECT substr(v.fecha,1,7) AS m, v.cliente AS cl, MIN(v.fecha) AS primera, SUM(i.total) AS t $BASE AND v.tipo='VEN' AND v.fecha >= '$DESDE' AND v.fecha <= '$(F $corte)' GROUP BY m, v.cliente"
$nomCli = @{}; foreach ($x in (Invoke-PanelSql "SELECT c.codigo, c.razon_social, c.vendedor FROM clientes c")) { $nomCli[[string]$x.codigo] = $x }
$porMes = @{}; foreach ($x in $craw) { $m = [string]$x.m; if (-not $porMes.ContainsKey($m)) { $porMes[$m] = @{} }; $porMes[$m][[string]$x.cl] = $x }
$mesesCli = @($porMes.Keys | Sort-Object)
$retencion = @()
for ($i = 1; $i -lt $mesesCli.Count; $i++) {
    $m = $mesesCli[$i]; $mp = $mesesCli[$i-1]; $cur = $porMes[$m]; $ant = $porMes[$mp]
    $antes = @{}; for ($j = 0; $j -lt $i - 1; $j++) { foreach ($c in $porMes[$mesesCli[$j]].Keys) { $antes[$c] = $true } }
    $ret = 0; $rec = 0; $nue = 0; $vRet = 0.0; $vRec = 0.0; $vNue = 0.0
    foreach ($c in $cur.Keys) { $t = D0 $cur[$c].t
        if ($ant.ContainsKey($c)) { $ret++; $vRet += $t } elseif ($antes.ContainsKey($c)) { $rec++; $vRec += $t } else { $nue++; $vNue += $t } }
    # para el mes en curso: comparar contra el mes anterior al mismo dia habil
    $retMismoDia = $null
    if ($m -eq $mes -and $i -ge 2) {
        $mpp = $mesesCli[$i-2]; $n = 0
        foreach ($c in $porMes[$mpp].Keys) { if ($ant.ContainsKey($c) -and ([string]$ant[$c].primera) -le (F $cortePrev)) { $n++ } }
        $retMismoDia = R4 (Div $n $porMes[$mpp].Count)
    }
    $retencion += [ordered]@{ mes = $m; mesAnterior = $mp; clientesMesAnterior = $ant.Count; compradores = $cur.Count
        retenidos = $ret; recuperados = $rec; nuevos = $nue; tasaRetencion = R4 (Div $ret $ant.Count)
        tasaRetencionMesAnteriorMismoDia = $retMismoDia
        ventaRetenidos = R0 $vRet; ventaRecuperados = R0 $vRec; ventaNuevos = R0 $vNue; pesoNuevos = R4 (Div $vNue ($vRet + $vRec + $vNue)) }
}
# clientes grandes del mes anterior que todavia no volvieron
$sinVolver = @()
if ($porMes.ContainsKey($mesPrev) -and $porMes.ContainsKey($mes)) {
    $sinVolver = @($porMes[$mesPrev].GetEnumerator() | Where-Object { -not $porMes[$mes].ContainsKey($_.Key) } | Sort-Object { -(D0 $_.Value.t) } | Select-Object -First 15 | ForEach-Object {
        $c = $nomCli[$_.Key]; $ve = if ($c) { [string]$c.vendedor } else { '' }
        [ordered]@{ codigo = $_.Key; nombre = $(if ($c) { [string]$c.razon_social } else { "Cliente $($_.Key)" }); vendedor = $(if ($nomVend.ContainsKey($ve)) { "$ve $($nomVend[$ve])" } else { $ve }); ventaMesAnterior = R0 (D0 $_.Value.t) } })
}

# ---------------- senales de alerta ----------------
$completas = @($semanas | Where-Object { $_.completa })
$ult = if ($completas.Count) { $completas[-1] } else { $null }; $ant = if ($completas.Count -gt 1) { $completas[-2] } else { $null }
function Prom($rows, $campo) { $v = @($rows | ForEach-Object { $_.$campo } | Where-Object { $null -ne $_ }); if ($v.Count) { ($v | Measure-Object -Average).Average } else { $null } }
$prev4 = if ($completas.Count -gt 1) { @($completas[([math]::Max(0, $completas.Count - 5))..($completas.Count - 2)]) } else { @() }
$pct = { param($x) if ($null -eq $x) { '-' } else { ('{0:N1}%' -f ($x * 100)) } }
$alertas = [ordered]@{}

$varMtd = Div ($actual.venta - $prevMtd.venta) $prevMtd.venta
$alertas.crecimiento = [ordered]@{ estado = $(if ($null -eq $varMtd) { 'n' } elseif ($varMtd -lt -0.02) { 'r' } elseif ($varMtd -lt 0.02) { 'a' } else { 'g' })
    texto = $(if ($null -eq $varMtd) { 'Sin mes anterior para comparar.' } else { "Contra el mes anterior a los mismos días hábiles: $(& $pct $varMtd)." }) }

$avance = Div $actual.venta $actual.objEq; $tiempo = $habTrans / $habTot
$cumplSem = @($completas | Where-Object { $_.objEq } | ForEach-Object { Div $_.venta $_.objEq })
$bajas = @($cumplSem | Select-Object -Last 2 | Where-Object { $_ -lt 0.9 })
$alertas.cumplimiento = [ordered]@{ estado = $(if ($null -eq $avance) { 'n' } elseif ($bajas.Count -ge 2 -or (Div $avance $tiempo) -lt 0.9) { 'r' } elseif ((Div $avance $tiempo) -lt 1) { 'a' } else { 'g' })
    texto = $(if ($null -eq $avance) { 'Sin objetivo cargado para el mes.' } else { "Lleva $(& $pct $avance) del objetivo de equilibrio con $(& $pct $tiempo) del mes." + $(if ($bajas.Count -ge 2) { ' Las dos últimas semanas quedaron debajo del 90%.' } else { '' }) }) }

$convProm = Prom $prev4 'conversion'
$alertas.conversion = [ordered]@{ estado = $(if (-not $ult -or $null -eq $convProm) { 'n' } elseif ($ult.conversion -lt $convProm - 0.03) { 'r' } elseif ($ult.conversion -lt $convProm) { 'a' } else { 'g' })
    texto = $(if ($ult) { "Última semana completa: $(& $pct $ult.conversion) de los clientes que compraron llevaron $Nombre (promedio de las 4 anteriores $(& $pct $convProm))." } else { '' }) }

$alertas.ticket = [ordered]@{ estado = $(if (-not $ult -or -not $ant) { 'n' } elseif ($ult.tickets -gt $ant.tickets -and $ult.ticketProm -lt $ant.ticketProm) { 'r' } elseif ($ult.ticketProm -lt (Prom $prev4 'ticketProm')) { 'a' } else { 'g' })
    texto = $(if ($ult -and $ant) { "Última semana: $($ult.tickets) tickets de `$$('{0:N0}' -f $ult.ticketProm) (la anterior $($ant.tickets) de `$$('{0:N0}' -f $ant.ticketProm))." } else { '' }) }

$aptProm = Prom $prev4 'articulosPorTicket'
$alertas.upt = [ordered]@{ estado = $(if (-not $ult -or $null -eq $aptProm) { 'n' } elseif ($ult.articulosPorTicket -lt $aptProm * 0.95) { 'r' } elseif ($ult.articulosPorTicket -lt $aptProm) { 'a' } else { 'g' })
    texto = $(if ($ult) { "Última semana: $('{0:N2}' -f $ult.articulosPorTicket) artículos $Nombre distintos y $('{0:N1}' -f $ult.upt) unidades por ticket (promedio anterior $('{0:N2}' -f $aptProm) artículos)." } else { '' }) }

$activos = @($vendedores | Where-Object { $_.ventaPrevMtd -gt 1000000 })
# brecha sin explicacion = cae 20 puntos mas que el promedio del proveedor contra el mes pasado
$caidas = @($activos | Where-Object { $_.venta -gt 0 -and $_.var -ne $null -and $null -ne $varMtd -and $_.var -lt $varMtd - 0.20 } | Sort-Object { $_.var })
$alertas.vendedores = [ordered]@{ estado = $(if ($caidas.Count -ge 3) { 'r' } elseif ($caidas.Count -gt 0) { 'a' } else { 'g' })
    texto = $(if ($caidas.Count) { "$Nombre va $(& $pct $varMtd) contra el mes pasado a los mismos días; $($caidas.Count) vendedores quedan 20 puntos o más por debajo: " + (($caidas | Select-Object -First 4 | ForEach-Object { "$($_.nombre) ($(& $pct $_.var))" }) -join ', ') + '.' } else { "Ningún vendedor queda 20 puntos por debajo de $Nombre en total ($(& $pct $varMtd))." }) }

$margenTot = Div ($totN - $totC) $totN
$flojas = @($lineas | Where-Object { $_.pesoVenta -gt 0.05 -and $_.margen -ne $null -and $_.margen -lt $margenTot - 0.03 })
$alertas.productos = [ordered]@{ estado = $(if ($flojas.Count -ge 2) { 'r' } elseif ($flojas.Count) { 'a' } else { 'g' })
    texto = $(if ($flojas.Count) { "Venden mucho con poco margen: " + (($flojas | ForEach-Object { "$($_.nombre) ($(& $pct $_.pesoVenta) de la venta, margen $(& $pct $_.margen))" }) -join ', ') + ". Margen $Nombre del mes: $(& $pct $margenTot)." } else { "Ninguna línea grande queda 3 puntos debajo del margen de $Nombre ($(& $pct $margenTot))." }) }

$rAct = $retencion | Where-Object { $_.mes -eq $mes }
$alertas.retencion = [ordered]@{ estado = $(if (-not $rAct) { 'n' } elseif ($rAct.pesoNuevos -gt 0.25) { 'r' } elseif ($null -ne $rAct.tasaRetencionMesAnteriorMismoDia -and $rAct.tasaRetencion -lt $rAct.tasaRetencionMesAnteriorMismoDia - 0.03) { 'a' } else { 'g' })
    texto = $(if ($rAct) { "Volvieron $($rAct.retenidos) de los $($rAct.clientesMesAnterior) clientes de $(([datetime]::ParseExact($rAct.mesAnterior + '-01', 'yyyy-MM-dd', $null)).ToString('MMMM', [Globalization.CultureInfo]'es-AR')) ($(& $pct $rAct.tasaRetencion))" + $(if ($null -ne $rAct.tasaRetencionMesAnteriorMismoDia) { "; el mes pasado a esta altura iba $(& $pct $rAct.tasaRetencionMesAnteriorMismoDia)" } else { '' }) + ". Los nuevos aportan $(& $pct $rAct.pesoNuevos) de la venta." } else { '' }) }

$devProm = Prom $prev4 'devolucionPct'
$sube3 = $false; if ($completas.Count -ge 3) { $z = $completas[-3..-1]; $sube3 = ($z[1].devolucionPct -gt $z[0].devolucionPct -and $z[2].devolucionPct -gt $z[1].devolucionPct) }
$alertas.devoluciones = [ordered]@{ estado = $(if (-not $ult) { 'n' } elseif ($sube3) { 'r' } elseif ($ult.devolucionPct -gt $devProm) { 'a' } else { 'g' })
    texto = $(if ($ult) { "Última semana: $(& $pct $ult.devolucionPct) de lo facturado volvió (rechazos + cambios); promedio de las 4 anteriores $(& $pct $devProm)." + $(if ($sube3) { ' Sube tres semanas seguidas.' } else { '' }) } else { '' }) }

# ---------------- salida ----------------
$out = [ordered]@{
    generadoEn = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    proveedor = $ProveedorObjetivo; nombre = $Nombre; codigo = $Proveedor
    periodo = [ordered]@{ mes = $mes; mesAnterior = $mesPrev; corte = F $corte; cortePrev = F $cortePrev; habilesTransc = $habTrans; habilesTotal = $habTot; baseDesde = '2026-08-03' }
    actual = $actual; mesAnteriorMismoDia = $prevMtd; mesAnteriorCompleto = $prevMes
    meses = $meses; semanas = $semanas
    vendedores = $vendedores
    productos = [ordered]@{ margen = R4 $margenTot; lineas = $lineas; skus = $skus }
    retencion = [ordered]@{ meses = $retencion; sinVolver = $sinVolver }
    alertas = $alertas
}
[IO.File]::WriteAllText($OutPath, ($out | ConvertTo-Json -Depth 8), $u8)
Write-Log ("Guardado: $OutPath | venta $mes {0:N0} (obj eq {1:N0}) | vs mes anterior {2} | alertas: {3}" -f $actual.venta, $actual.objEq, (& $pct $varMtd), (($alertas.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value.estado)" }) -join ' '))
