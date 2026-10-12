<#
Tablero del Director Comercial (uso personal de Gisela, 12/10/2026): los 8 indicadores de
"Lunes 8:00 - que deberia mirar un director comercial":
  1 Ventas vs objetivo   2 Margen   3 Volumen y mix   4 Clientes   5 Pipeline y conversion
  6 Canales              7 Forecast 8 Ejecucion (prioridades de la semana)

Todo sale de la base compartida (ventas, venta_items, clientes, articulos, digip_*), asi se
actualiza solo. Venta neta = VEN + DEB - devoluciones (DEV-RE / DEV-CA), sin los vendedores
1176, 43, 16 y 37 (no son vendedores reales). Septiembre da $1.007M, igual que Rentabilidad.

Objetivos del mes: scripts/privado/objetivos/data.json (cifrado en el repo; si no esta en claro
se descifra con protect-tableros.ps1). Sale del Excel "Objetivos <Mes> <Anio>.xlsx".
Dias habiles = lunes a viernes sin los feriados de $FERIADOS.
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/director"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "director-template.html"),
    [string]$ObjetivosDir = (Join-Path $PSScriptRoot "privado"),
    [string]$CategoriasPath = (Join-Path $PSScriptRoot "categorias-articulos.json")
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
function R4($x) { return [math]::Round([double]$x, 4) }
function Div($a, $b) { if ([double]$b -eq 0) { return $null } return [double]$a / [double]$b }

$EXCL = "'1176','43','043','16','016','37','037'"
$MAYORISTAS = @('76','38','50','18','40')
$FERIADOS = @('2026-08-17','2026-10-12','2026-11-23','2026-12-08','2026-12-25','2027-01-01')
$SIGNO = "CASE WHEN v.tipo IN ('VEN','DEB') THEN 1 WHEN v.tipo LIKE 'DEV%' THEN -1 ELSE 0 END"

function Es-Habil([datetime]$d) { return ($d.DayOfWeek -ne 'Saturday' -and $d.DayOfWeek -ne 'Sunday' -and $FERIADOS -notcontains $d.ToString('yyyy-MM-dd')) }
function Habiles([datetime]$desde, [datetime]$hasta) { $n=0; for ($d=$desde.Date; $d -le $hasta.Date; $d=$d.AddDays(1)) { if (Es-Habil $d) { $n++ } }; return $n }

# ---------------- periodo ----------------
$corte = [datetime]::ParseExact(((Invoke-PanelSql "SELECT MAX(fecha) AS f FROM ventas WHERE tipo='VEN'")[0].f), 'yyyy-MM-dd', $null)
$ini = (Get-Date -Year $corte.Year -Month $corte.Month -Day 1).Date
$fin = $ini.AddMonths(1).AddDays(-1)
$iniPrev = $ini.AddMonths(-1); $finPrev = $ini.AddDays(-1)
$iniPrev2 = $ini.AddMonths(-2)
$mes = $ini.ToString('yyyy-MM'); $mesPrev = $iniPrev.ToString('yyyy-MM'); $mesPrev2 = $iniPrev2.ToString('yyyy-MM')
$habTot = Habiles $ini $fin; $habTrans = Habiles $ini $corte; if ($habTrans -lt 1) { $habTrans = 1 }
$habTotPrev = Habiles $iniPrev $finPrev
# mismo dia habil del mes anterior
$k=0; $cortePrev = $finPrev
for ($d=$iniPrev; $d -le $finPrev; $d=$d.AddDays(1)) { if (Es-Habil $d) { $k++; if ($k -eq $habTrans) { $cortePrev = $d; break } } }
Write-Log "Mes $mes al $($corte.ToString('dd/MM')) | dias habiles $habTrans/$habTot | comparo con $mesPrev al $($cortePrev.ToString('dd/MM'))"

# ---------------- objetivos ----------------
$objPath = Join-Path $ObjetivosDir "objetivos/data.json"
if (-not (Test-Path $objPath)) { & (Join-Path $PSScriptRoot "protect-tableros.ps1") -Modo Descifrar -DocsDir $ObjetivosDir -Tableros objetivos | Out-Null }
$obj = $null
if (Test-Path $objPath) { $o = [IO.File]::ReadAllText($objPath, $u8) | ConvertFrom-Json; if ($o.mes -eq $mes) { $obj = $o } }
if (-not $obj) { Write-Log "AVISO: no hay objetivos cargados para $mes (se muestra sin objetivo)." }

# ---------------- datos ----------------
Write-Log "Venta diaria..."
$diario = Invoke-PanelSql "SELECT v.fecha AS f, SUM($SIGNO * v.total) AS n FROM ventas v WHERE v.fecha >= '$($iniPrev2.ToString('yyyy-MM-dd'))' AND COALESCE(v.vendedor,'') NOT IN ($EXCL) GROUP BY v.fecha ORDER BY v.fecha"
Write-Log "Venta por articulo (mes y mes anterior)..."
$arts = Invoke-PanelSql @"
SELECT substr(v.fecha,1,7) AS m, i.articulo AS a,
  SUM($SIGNO * i.neto) AS n, SUM($SIGNO * i.total) AS t,
  SUM($SIGNO * i.precio_costo * i.cantidad * COALESCE(i.factor,1)) AS c,
  SUM(CASE WHEN v.tipo='VEN' THEN i.neto ELSE 0 END) AS nv,
  SUM(CASE WHEN v.tipo='VEN' THEN i.precio_unitario * i.cantidad ELSE 0 END) AS lv,
  SUM($SIGNO * i.cantidad * COALESCE(i.factor,1)) AS u
FROM ventas v JOIN venta_items i ON i.venta_id = v.id
WHERE v.fecha >= '$($iniPrev.ToString('yyyy-MM-dd'))' AND COALESCE(v.vendedor,'') NOT IN ($EXCL)
GROUP BY m, i.articulo
"@
# mismo periodo del mes anterior (para comparar margen y mix con los mismos dias)
$artsPrevMtd = Invoke-PanelSql @"
SELECT i.articulo AS a, SUM($SIGNO * i.neto) AS n, SUM($SIGNO * i.total) AS t, SUM($SIGNO * i.precio_costo * i.cantidad * COALESCE(i.factor,1)) AS c
FROM ventas v JOIN venta_items i ON i.venta_id = v.id
WHERE v.fecha >= '$($iniPrev.ToString('yyyy-MM-dd'))' AND v.fecha <= '$($cortePrev.ToString('yyyy-MM-dd'))' AND COALESCE(v.vendedor,'') NOT IN ($EXCL)
GROUP BY i.articulo
"@
Write-Log "Clientes..."
$cliMes = Invoke-PanelSql @"
SELECT substr(v.fecha,1,7) AS m, v.cliente AS cl, SUM($SIGNO * v.total) AS n,
  COUNT(DISTINCT CASE WHEN v.tipo='VEN' THEN v.fecha END) AS ped
FROM ventas v WHERE v.fecha >= '$($iniPrev2.ToString('yyyy-MM-dd'))' AND COALESCE(v.vendedor,'') NOT IN ($EXCL)
GROUP BY m, v.cliente
"@
$cliPrevMtd = Invoke-PanelSql "SELECT v.cliente AS cl, SUM($SIGNO * v.total) AS n, COUNT(DISTINCT CASE WHEN v.tipo='VEN' THEN v.fecha END) AS ped FROM ventas v WHERE v.fecha >= '$($iniPrev.ToString('yyyy-MM-dd'))' AND v.fecha <= '$($cortePrev.ToString('yyyy-MM-dd'))' AND COALESCE(v.vendedor,'') NOT IN ($EXCL) GROUP BY v.cliente"
$primera = @{}; Invoke-PanelSql "SELECT cliente AS cl, MIN(fecha) AS f FROM ventas WHERE tipo='VEN' GROUP BY cliente" | ForEach-Object { $primera[[string]$_.cl] = [string]$_.f }
Write-Log "Vendedores..."
$vendMes = Invoke-PanelSql "SELECT v.vendedor AS ve, SUM($SIGNO * v.total) AS n FROM ventas v WHERE v.fecha >= '$($ini.ToString('yyyy-MM-dd'))' AND COALESCE(v.vendedor,'') NOT IN ($EXCL) GROUP BY v.vendedor"
$vendPrevMtd = Invoke-PanelSql "SELECT v.vendedor AS ve, SUM($SIGNO * v.total) AS n FROM ventas v WHERE v.fecha >= '$($iniPrev.ToString('yyyy-MM-dd'))' AND v.fecha <= '$($cortePrev.ToString('yyyy-MM-dd'))' AND COALESCE(v.vendedor,'') NOT IN ($EXCL) GROUP BY v.vendedor"
$vendNom = @{}; Invoke-PanelSql "SELECT codigo, nombre FROM vendedores" | ForEach-Object { $vendNom[([string]$_.codigo).TrimStart('0')] = [string]$_.nombre }
Write-Log "Catalogos, cartera, Digip..."
$catArt = @{}; Invoke-PanelSql "SELECT a.codigo, a.descripcion, p.nombre AS prov FROM articulos a LEFT JOIN proveedores p ON p.codigo = a.proveedor" | ForEach-Object { $catArt[[string]$_.codigo] = $_ }
$categ = @{}; if (Test-Path $CategoriasPath) { ([IO.File]::ReadAllText($CategoriasPath, $u8) | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $categ[$_.Name] = [string]$_.Value } }
$clientes = Invoke-PanelSql "SELECT codigo, nombre, vendedor, subramo FROM clientes WHERE activo = 1"
$subNom = @{}; Invoke-PanelSql "SELECT codigo, descripcion FROM subramos" | ForEach-Object { $subNom[[string]$_.codigo] = ([string]$_.descripcion).Trim() }
$pend = Invoke-PanelSql "SELECT importe, fecha, vendedor FROM digip_pedidos WHERE estado = 'Pendiente'"
$stockDigip = @{}; Invoke-PanelSql "SELECT articulo, SUM(cantidad) AS q FROM digip_stock GROUP BY articulo" | ForEach-Object { $stockDigip[[string]$_.articulo] = [double]$_.q }
$digipAct = (Invoke-PanelSql "SELECT MAX(actualizado) AS a FROM digip_stock")[0].a

# ---------------- 1 + 7: ventas y forecast ----------------
$serie = @{}; foreach ($x in $diario) { $serie[[string]$x.f] = [double]$x.n }
function Acum($desde, $hasta) { $s=0.0; for ($d=$desde; $d -le $hasta; $d=$d.AddDays(1)) { $s += [double]$serie[$d.ToString('yyyy-MM-dd')] }; return $s }
$mtd = Acum $ini $corte; $prevMtd = Acum $iniPrev $cortePrev; $prevTotal = Acum $iniPrev $finPrev
$curvaAct = @(); $curvaPrev = @(); $a=0.0; $hab=0
for ($d=$ini; $d -le $corte; $d=$d.AddDays(1)) { $a += [double]$serie[$d.ToString('yyyy-MM-dd')]; if (Es-Habil $d) { $hab++; $curvaAct += [ordered]@{ dia=$hab; fecha=$d.ToString('yyyy-MM-dd'); acum=R0 $a } } }
$a=0.0; $hab=0; $prevPorDia = @{}
for ($d=$iniPrev; $d -le $finPrev; $d=$d.AddDays(1)) { $a += [double]$serie[$d.ToString('yyyy-MM-dd')]; if (Es-Habil $d) { $hab++; $curvaPrev += [ordered]@{ dia=$hab; fecha=$d.ToString('yyyy-MM-dd'); acum=R0 $a }; $prevPorDia[$hab] = $a } }
$forecastLineal = $mtd / $habTrans * $habTot
# ritmo de los ultimos 5 dias habiles
$ult = @($curvaAct | Select-Object -Last 6)
$ritmo5 = if ($ult.Count -ge 2) { ([double]$ult[-1].acum - [double]$ult[0].acum) / ($ult.Count - 1) } else { $mtd / $habTrans }
$forecastRitmo = $mtd + $ritmo5 * ($habTot - $habTrans)
# precision: que hubiera dicho el forecast lineal el mes pasado en este mismo dia habil
$fcPrevMismoDia = if ($prevPorDia.ContainsKey($habTrans)) { $prevPorDia[$habTrans] / $habTrans * $habTotPrev } else { $null }
$errorPrev = if ($fcPrevMismoDia) { $fcPrevMismoDia / $prevTotal - 1 } else { $null }
$restantes = $habTot - $habTrans
$ventas = [ordered]@{
    mtd = R0 $mtd; prevMtd = R0 $prevMtd; prevTotal = R0 $prevTotal; varVsPrev = R4 (Div ($mtd - $prevMtd) $prevMtd)
    objEq = $(if ($obj) { R0 $obj.equilibrio } else { $null }); objIdeal = $(if ($obj) { R0 $obj.ideal } else { $null })
    avanceEq = $(if ($obj) { R4 (Div $mtd $obj.equilibrio) } else { $null }); avanceTiempo = R4 ($habTrans / $habTot)
    ritmoActual = R0 ($mtd / $habTrans)
    ritmoNecEq = $(if ($obj -and $restantes -gt 0) { R0 (([double]$obj.equilibrio - $mtd) / $restantes) } else { $null })
    ritmoNecIdeal = $(if ($obj -and $restantes -gt 0) { R0 (([double]$obj.ideal - $mtd) / $restantes) } else { $null })
}
$forecast = [ordered]@{ lineal = R0 $forecastLineal; ritmo5 = R0 $forecastRitmo; ritmoDiario5 = R0 $ritmo5; prevTotal = R0 $prevTotal
    fcPrevMismoDia = $(if ($fcPrevMismoDia) { R0 $fcPrevMismoDia } else { $null }); errorPrevMismoDia = $(if ($null -ne $errorPrev) { R4 $errorPrev } else { $null })
    curvaActual = $curvaAct; curvaPrev = $curvaPrev }

# ---------------- 2 + 3: margen, mix, SKUs ----------------
$porProv = @{}; $porCat = @{}; $skus = @{}
$tot = @{ n=0.0; c=0.0; nv=0.0; lv=0.0 }; $totP = @{ n=0.0; c=0.0; nv=0.0; lv=0.0 }
foreach ($x in $arts) {
    $a = [string]$x.a; $info = $catArt[$a]; $prov = if ($info -and $info.prov) { [string]$info.prov } else { 'Sin proveedor' }
    $cat = if ($categ.ContainsKey($a)) { $categ[$a] } else { 'Sin categoria' }
    $esAct = ($x.m -eq $mes)
    if (-not $esAct -and $x.m -ne $mesPrev) { continue }
    $t = $totP; if ($esAct) { $t = $tot }
    $t.n += [double]$x.n; $t.c += [double]$x.c; $t.nv += [double]$x.nv; $t.lv += [double]$x.lv
    if ($esAct) {
        # venta = con IVA (t), como los objetivos; margen = sobre el neto sin IVA (n y c)
        if (-not $porProv[$prov]) { $porProv[$prov] = @{ n=0.0; c=0.0; nv=0.0; lv=0.0; pn=0.0; pc=0.0; pt=0.0; vt=0.0; pvt=0.0 } }
        $pp = $porProv[$prov]; $pp.n += [double]$x.n; $pp.c += [double]$x.c; $pp.nv += [double]$x.nv; $pp.lv += [double]$x.lv; $pp.vt += [double]$x.t
        if (-not $porCat[$cat]) { $porCat[$cat] = @{ n=0.0; pn=0.0 } }; $porCat[$cat].n += [double]$x.t
        $skus[$a] = @{ n=[double]$x.n; c=[double]$x.c; u=[double]$x.u; t=[double]$x.t; pn=0.0 }
    } else {
        if (-not $porProv[$prov]) { $porProv[$prov] = @{ n=0.0; c=0.0; nv=0.0; lv=0.0; pn=0.0; pc=0.0; pt=0.0; vt=0.0; pvt=0.0 } }
        $porProv[$prov].pt += [double]$x.t
    }
}
foreach ($x in $artsPrevMtd) {
    $a = [string]$x.a; $info = $catArt[$a]; $prov = if ($info -and $info.prov) { [string]$info.prov } else { 'Sin proveedor' }
    $cat = if ($categ.ContainsKey($a)) { $categ[$a] } else { 'Sin categoria' }
    if (-not $porProv[$prov]) { $porProv[$prov] = @{ n=0.0; c=0.0; nv=0.0; lv=0.0; pn=0.0; pc=0.0; pt=0.0; vt=0.0; pvt=0.0 } }
    $porProv[$prov].pn += [double]$x.n; $porProv[$prov].pc += [double]$x.c; $porProv[$prov].pvt += [double]$x.t
    if (-not $porCat[$cat]) { $porCat[$cat] = @{ n=0.0; pn=0.0 } }; $porCat[$cat].pn += [double]$x.t
    if ($skus.ContainsKey($a)) { $skus[$a].pn = [double]$x.t }
}
$prevMtdN = ($artsPrevMtd | Measure-Object n -Sum).Sum; $prevMtdC = ($artsPrevMtd | Measure-Object c -Sum).Sum
$margen = [ordered]@{
    margen = R4 (Div ($tot.n - $tot.c) $tot.n); margenPrevMtd = R4 (Div ($prevMtdN - $prevMtdC) $prevMtdN); margenPrevTotal = R4 (Div ($totP.n - $totP.c) $totP.n)
    contribucion = R0 ($tot.n - $tot.c); contribucionPrevMtd = R0 ($prevMtdN - $prevMtdC)
    descuento = R4 (1 - (Div $tot.nv $tot.lv)); descuentoPrev = R4 (1 - (Div $totP.nv $totP.lv))
    proveedores = @($porProv.GetEnumerator() | Where-Object { $_.Value.n -ne 0 -or $_.Value.pn -ne 0 } | Sort-Object { -$_.Value.vt } | ForEach-Object {
        $v = $_.Value; $o = if ($obj -and $obj.proveedores.($_.Key)) { $obj.proveedores.($_.Key) } else { $null }
        [ordered]@{ nombre=$_.Key; venta=R0 $v.vt; ventaPrevMtd=R0 $v.pvt; ventaPrevTotal=R0 $v.pt; margen=R4 (Div ($v.n-$v.c) $v.n); margenPrevMtd=R4 (Div ($v.pn-$v.pc) $v.pn)
            descuento=R4 (1 - (Div $v.nv $v.lv)); objEq=$(if ($o) { R0 $o.eq } else { $null }); objIdeal=$(if ($o) { R0 $o.ideal } else { $null }) } })
}
$mix = [ordered]@{
    categorias = @($porCat.GetEnumerator() | Sort-Object { -$_.Value.n } | ForEach-Object { [ordered]@{ nombre=$_.Key; venta=R0 $_.Value.n; ventaPrevMtd=R0 $_.Value.pn } })
    skus = @($skus.GetEnumerator() | Sort-Object { -$_.Value.t } | Select-Object -First 25 | ForEach-Object {
        $a=$_.Key; $info=$catArt[$a]
        [ordered]@{ codigo=$a; nombre=$(if ($info) { [string]$info.descripcion } else { $a }); proveedor=$(if ($info) { [string]$info.prov } else { '' }); venta=R0 $_.Value.t; ventaPrevMtd=R0 $_.Value.pn; margen=R4 (Div ($_.Value.n-$_.Value.c) $_.Value.n); unidades=R0 $_.Value.u; stockDigip=$(if ($stockDigip.ContainsKey($a)) { R0 $stockDigip[$a] } else { 0 }) } })
    skusActivos = @($skus.Values | Where-Object { $_.n -gt 0 }).Count
}

# ---------------- 4 + 5 + 6: clientes, pipeline, canales ----------------
$cli = @{}
foreach ($x in $cliMes) { $c=[string]$x.cl; if (-not $cli[$c]) { $cli[$c] = @{} }; $cli[$c][[string]$x.m] = @{ n=[double]$x.n; ped=[int]$x.ped } }
$cliPM = @{}; $cliPMped = @{}; foreach ($x in $cliPrevMtd) { $cliPM[[string]$x.cl] = [double]$x.n; $cliPMped[[string]$x.cl] = [int]$x.ped }
$infoCli = @{}; foreach ($c in $clientes) { $infoCli[[string]$c.codigo] = $c }
$compAct = @($cli.Keys | Where-Object { $cli[$_][$mes] -and $cli[$_][$mes].ped -gt 0 })
$compPrev = @($cli.Keys | Where-Object { $cli[$_][$mesPrev] -and $cli[$_][$mesPrev].ped -gt 0 })
$compPrevMtdL = @($cliPMped.Keys | Where-Object { $cliPMped[$_] -gt 0 }); $compPrevMtd = $compPrevMtdL.Count
$nuevos = @($compAct | Where-Object { $primera[$_] -and $primera[$_] -ge $ini.ToString('yyyy-MM-dd') })
$recurrentes = @($compAct | Where-Object { $cli[$_][$mesPrev] -and $cli[$_][$mesPrev].ped -gt 0 })
$pedidos = ($compAct | ForEach-Object { $cli[$_][$mes].ped } | Measure-Object -Sum).Sum
$ventaComp = ($compAct | ForEach-Object { $cli[$_][$mes].n } | Measure-Object -Sum).Sum
$pedPrev = ($compPrevMtdL | ForEach-Object { $cliPMped[$_] } | Measure-Object -Sum).Sum
$ventaPrevComp = ($compPrevMtdL | ForEach-Object { $cliPM[$_] } | Measure-Object -Sum).Sum
$sinVolver = @($compPrev | Where-Object { -not ($cli[$_][$mes] -and $cli[$_][$mes].ped -gt 0) })
$perdidos = @($sinVolver | Where-Object { $cli[$_][$mesPrev2] -and $cli[$_][$mesPrev2].ped -gt 0 })   # compraban 2 meses seguidos y este no
function NomCli($c) { $i = $infoCli[$c]; if ($i) { return [string]$i.nombre } return "Cliente $c" }
function VendCli($c) { $i = $infoCli[$c]; if ($i -and $i.vendedor) { $v = ([string]$i.vendedor).TrimStart('0'); return "$v $($vendNom[$v])" } return '' }
$clientesOut = [ordered]@{
    compradores = $compAct.Count; compradoresPrevMtd = $compPrevMtd; compradoresPrevTotal = $compPrev.Count
    nuevos = $nuevos.Count; recurrentes = $recurrentes.Count; recurrencia = R4 (Div $recurrentes.Count $compAct.Count)
    frecuencia = [math]::Round((Div $pedidos $compAct.Count), 2); frecuenciaPrev = [math]::Round((Div $pedPrev $compPrevMtd), 2)
    ticket = R0 (Div $ventaComp $pedidos); ticketPrev = R0 (Div $ventaPrevComp $pedPrev)
    sinVolver = $sinVolver.Count; ventaSinVolverPrev = R0 (($sinVolver | ForEach-Object { $cli[$_][$mesPrev].n } | Measure-Object -Sum).Sum)
    enRiesgo = $perdidos.Count
    topSinVolver = @($sinVolver | Sort-Object { -$cli[$_][$mesPrev].n } | Select-Object -First 15 | ForEach-Object { [ordered]@{ codigo=$_; nombre=NomCli $_; vendedor=VendCli $_; ventaPrev=R0 $cli[$_][$mesPrev].n; dosMeses=$(if ($cli[$_][$mesPrev2] -and $cli[$_][$mesPrev2].ped -gt 0) { $true } else { $false }) } })
    nuevosLista = @($nuevos | Sort-Object { -$cli[$_][$mes].n } | Select-Object -First 10 | ForEach-Object { [ordered]@{ codigo=$_; nombre=NomCli $_; vendedor=VendCli $_; venta=R0 $cli[$_][$mes].n } })
}
# pipeline: cartera, cobertura y clientes clave (80% de la venta de los 2 meses anteriores) sin compra
$cartera = @($clientes | Where-Object { $_.vendedor -and ("'" + ([string]$_.vendedor) + "'") -notin ($EXCL -split ',') })
$cartCod = @{}; foreach ($c in $cartera) { $cartCod[[string]$c.codigo] = $true }
$cartCompr = @($compAct | Where-Object { $cartCod.ContainsKey($_) }).Count
$base2m = @{}; foreach ($c in $cli.Keys) { $s = 0.0; foreach ($m in $mesPrev,$mesPrev2) { if ($cli[$c][$m]) { $s += $cli[$c][$m].n } }; if ($s -gt 0) { $base2m[$c] = $s } }
$tot2m = ($base2m.Values | Measure-Object -Sum).Sum; $acc = 0.0; $clave = @()
foreach ($e in ($base2m.GetEnumerator() | Sort-Object { -$_.Value })) { if ($acc / $tot2m -ge 0.8) { break }; $acc += $e.Value; $clave += $e.Key }
$claveSin = @($clave | Where-Object { -not ($cli[$_][$mes] -and $cli[$_][$mes].ped -gt 0) })
$pendHoy = (Get-Date)
$pAnt = @{ masDe7=0; de3a7=0; de1a3=0; menosDe1=0 }
foreach ($p in $pend) { $f = [datetime]::Parse([string]$p.fecha, [Globalization.CultureInfo]::InvariantCulture); $d = ($pendHoy - $f).TotalDays; if ($d -gt 7) { $pAnt.masDe7++ } elseif ($d -gt 3) { $pAnt.de3a7++ } elseif ($d -gt 1) { $pAnt.de1a3++ } else { $pAnt.menosDe1++ } }
$pipeline = [ordered]@{
    cartera = $cartera.Count; carteraCompro = $cartCompr; cobertura = R4 (Div $cartCompr $cartera.Count)
    clientesClave = $clave.Count; claveSinCompra = $claveSin.Count; ventaMensualClaveSin = R0 ((($claveSin | ForEach-Object { $base2m[$_] }) | Measure-Object -Sum).Sum / 2)
    topClaveSin = @($claveSin | Sort-Object { -$base2m[$_] } | Select-Object -First 15 | ForEach-Object { [ordered]@{ codigo=$_; nombre=NomCli $_; vendedor=VendCli $_; promedioMensual=R0 ($base2m[$_]/2) } })
    pendientes = @($pend).Count; pendientesImporte = R0 (($pend | Measure-Object importe -Sum).Sum); pendientesAntiguedad = $pAnt; digipActualizado = [string]$digipAct
}
# canales
$canal = @{ Preventa=@{n=0.0;pn=0.0}; Mayorista=@{n=0.0;pn=0.0} }
$vendRows = @{}
foreach ($x in $vendMes) { $v=([string]$x.ve).TrimStart('0'); $k = if ($MAYORISTAS -contains $v) { 'Mayorista' } else { 'Preventa' }; $canal[$k].n += [double]$x.n; $vendRows[$v] = @{ n=[double]$x.n; pn=0.0 } }
foreach ($x in $vendPrevMtd) { $v=([string]$x.ve).TrimStart('0'); $k = if ($MAYORISTAS -contains $v) { 'Mayorista' } else { 'Preventa' }; $canal[$k].pn += [double]$x.n; if (-not $vendRows[$v]) { $vendRows[$v] = @{ n=0.0; pn=0.0 } }; $vendRows[$v].pn = [double]$x.n }
$ramos = @{}
foreach ($c in $cli.Keys) {
    $sr = if ($infoCli[$c] -and $infoCli[$c].subramo) { [string]$infoCli[$c].subramo } else { '' }
    $nom = if ($subNom.ContainsKey($sr)) { $subNom[$sr] } else { 'Sin ramo' }
    if ($nom -match '^(?i)autoserv') { $nom = 'Autoservicio' } elseif ($nom -match '^(?i)almac') { $nom = 'Almacen' } elseif ($nom -match '^(?i)mayorist') { $nom = 'Mayorista' } elseif ($nom -match '^(?i)kiosc') { $nom = 'Kiosco' }
    if (-not $ramos[$nom]) { $ramos[$nom] = @{ n=0.0; pn=0.0; cli=0 } }
    if ($cli[$c][$mes]) { $ramos[$nom].n += $cli[$c][$mes].n; if ($cli[$c][$mes].ped -gt 0) { $ramos[$nom].cli++ } }
    if ($cliPM.ContainsKey($c)) { $ramos[$nom].pn += $cliPM[$c] }
}
$canales = [ordered]@{
    canales = @('Preventa','Mayorista' | ForEach-Object { $o = if ($obj -and $obj.canales.$_) { $obj.canales.$_ } else { $null }; [ordered]@{ nombre=$_; venta=R0 $canal[$_].n; ventaPrevMtd=R0 $canal[$_].pn; objEq=$(if ($o) { R0 $o.eq } else { $null }); objIdeal=$(if ($o) { R0 $o.ideal } else { $null }) } })
    ramos = @($ramos.GetEnumerator() | Where-Object { $_.Value.n -ne 0 -or $_.Value.pn -ne 0 } | Sort-Object { -$_.Value.n } | ForEach-Object { [ordered]@{ nombre=$_.Key; venta=R0 $_.Value.n; ventaPrevMtd=R0 $_.Value.pn; compradores=$_.Value.cli } })
    vendedores = @($vendRows.GetEnumerator() | ForEach-Object {
        $v=$_.Key; $o = if ($obj -and $obj.vendedores.$v) { $obj.vendedores.$v } else { $null }
        [ordered]@{ codigo=$v; nombre=$(if ($o) { [string]$o.nombre } else { [string]$vendNom[$v] }); canal=$(if ($MAYORISTAS -contains $v) { 'Mayorista' } else { 'Preventa' }); venta=R0 $_.Value.n; ventaPrevMtd=R0 $_.Value.pn; objEq=$(if ($o) { R0 $o.eq } else { $null }); objIdeal=$(if ($o) { R0 $o.ideal } else { $null }) } } | Sort-Object { -[double]$_.venta })
}

# ---------------- 8: ejecucion (prioridades armadas con reglas) ----------------
$prio = @()
$tiempo = $habTrans / $habTot
if ($obj) {
    $falta = [double]$obj.equilibrio - $mtd
    $prio += [ordered]@{ tipo='ventas'; titulo=("Vender `$" + [math]::Round($ventas.ritmoNecEq/1e6,1) + "M por dia habil para llegar al equilibrio"); detalle=("Hoy el ritmo es `$" + [math]::Round($ventas.ritmoActual/1e6,1) + "M por dia. Faltan `$" + [math]::Round($falta/1e6,0) + "M en $restantes dias habiles.") }
    $brechas = @($margen.proveedores | Where-Object { $_.objEq } | ForEach-Object { [pscustomobject]@{ n=$_.nombre; gap=([double]$_.objEq * $tiempo - [double]$_.venta) } } | Where-Object { $_.gap -gt 0 } | Sort-Object { -$_.gap } | Select-Object -First 3)
    foreach ($b in $brechas) { $prio += [ordered]@{ tipo='proveedor'; titulo=("$($b.n): `$" + [math]::Round($b.gap/1e6,1) + "M atrasado contra el ritmo del objetivo"); detalle='Es lo que deberia llevar vendido a esta altura del mes para cumplir su objetivo de equilibrio.' } }
    $vb = @($canales.vendedores | Where-Object { $_.objEq -and $_.canal -eq 'Preventa' } | ForEach-Object { [pscustomobject]@{ n="$($_.codigo) $($_.nombre)"; av=(Div $_.venta $_.objEq) } } | Sort-Object av | Select-Object -First 5)
    if ($vb.Count) { $prio += [ordered]@{ tipo='vendedores'; titulo='Vendedores con menor avance contra su objetivo'; detalle=(($vb | ForEach-Object { "$($_.n) ($([math]::Round($_.av*100))%)" }) -join ', ') + (" -- el mes va " + [math]::Round($tiempo*100) + "% transcurrido.") } }
}
$quiebres = @($mix.skus | Where-Object { $_.stockDigip -le 0 } | Select-Object -First 5)
if ($quiebres.Count) { $prio += [ordered]@{ tipo='stock'; titulo="$($quiebres.Count) de los 25 articulos que mas venden no tienen stock en Digip"; detalle=(($quiebres | ForEach-Object { $_.nombre }) -join ', ') } }
if ($claveSin.Count) { $prio += [ordered]@{ tipo='clientes'; titulo="$($claveSin.Count) clientes clave todavia no compraron este mes"; detalle=("Promedian `$" + [math]::Round($pipeline.ventaMensualClaveSin/1e6,1) + "M por mes. Los primeros: " + (($pipeline.topClaveSin | Select-Object -First 5 | ForEach-Object { $_.nombre }) -join ', ') + '.') } }
if ($pAnt.masDe7 -gt 0) { $prio += [ordered]@{ tipo='pedidos'; titulo="$($pAnt.masDe7) pedidos con mas de 7 dias sin despachar"; detalle='Revisarlos uno por uno con el deposito: stock, direccion o credito.' } }

$out = [ordered]@{
    generadoEn = (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
    periodo = [ordered]@{ mes=$mes; mesPrev=$mesPrev; corte=$corte.ToString('yyyy-MM-dd'); cortePrev=$cortePrev.ToString('yyyy-MM-dd'); habilesTransc=$habTrans; habilesTotal=$habTot; habilesTotalPrev=$habTotPrev }
    objetivosFuente = $(if ($obj) { [string]$obj.fuente } else { $null })
    ventas = $ventas; margen = $margen; mix = $mix; clientes = $clientesOut; pipeline = $pipeline; canales = $canales; forecast = $forecast; prioridades = $prio
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
[IO.File]::WriteAllText($OutPath, ($out | ConvertTo-Json -Depth 8 -Compress), $u8)
Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force
Write-Log ("Guardado: $OutPath | venta MTD " + [math]::Round($mtd) + " | margen " + $margen.margen + " | compradores " + $compAct.Count)
