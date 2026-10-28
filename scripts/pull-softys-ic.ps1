<#
Inteligencia Comercial (docs/softys-ic, tablero personal de Gisela). Empezo solo con Softys (28/10/2026)
y el mismo dia sumo solapas de Georgalos, Clorox (Grupo Ayudin) e Ilolay, mas una de recomendaciones
"para atacar".

Corre pull-ic-proveedor.ps1 para cada proveedor (los 9 indicadores) y agrega:
  - oportunidades por cliente: los que este mes ya compraron a la empresa pero no al proveedor,
    separados en "lo compraban el mes pasado" (recuperar) y "compran los otros proveedores y este no" (venta cruzada);
  - recomendaciones armadas con las senales de alerta de cada solapa.
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/softys-ic"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "softys-ic-template.html")
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$u8 = New-Object System.Text.UTF8Encoding $false
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
function Invoke-PanelSql([string]$Sql) {
    $uri = "$($config.baseUrl)/consulta?sql=" + [uri]::EscapeDataString($Sql)
    $r = Invoke-RestMethod -Uri $uri -Headers @{ Authorization = "Bearer $($config.clave)" } -Method Get
    if ($r.truncado) { throw "La consulta vino truncada (mas de 20000 filas): acotarla." }
    return $r.filas
}
function D0($x) { if ($null -eq $x) { 0.0 } else { [double]$x } }
function Div($a, $b) { if ($null -eq $b -or [double]$b -eq 0) { return $null } return [double]$a / [double]$b }
function Mill($x) { '$' + ('{0:N1}' -f ([double]$x / 1e6)) + 'M' }
function Pct($x) { if ($null -eq $x) { '-' } else { ('{0:N1}%' -f ($x * 100)) } }

$PROVEEDORES = @(
    @{ clave = 'softys';    codigo = '150'; objetivo = 'SOFTYS ARGENTINA S.A';       nombre = 'Softys';    linea = 'familia' }
    @{ clave = 'georgalos'; codigo = '42';  objetivo = 'GEORGALOS';                  nombre = 'Georgalos'; linea = 'descripcion' }
    @{ clave = 'clorox';    codigo = '247'; objetivo = 'GRUPO AYUDIN ARGENTINA S.A'; nombre = 'Clorox';    linea = 'familia' }
    @{ clave = 'ilolay';    codigo = '2';   objetivo = 'ILOLAY';                     nombre = 'Ilolay';    linea = 'descripcion' }
)
$EXCL = "'1176','43','043','16','016','37','037'"
$tmpDir = Join-Path ([IO.Path]::GetTempPath()) "ic-proveedores"; if (-not (Test-Path $tmpDir)) { New-Item -ItemType Directory -Path $tmpDir | Out-Null }

$datos = [ordered]@{}
foreach ($p in $PROVEEDORES) {
    $out = Join-Path $tmpDir "$($p.clave).json"
    & (Join-Path $PSScriptRoot "pull-ic-proveedor.ps1") -ConfigPath $ConfigPath -OutPath $out -Proveedor $p.codigo -ProveedorObjetivo $p.objetivo -Nombre $p.nombre -LineaPor $p.linea
    $datos[$p.clave] = [IO.File]::ReadAllText($out, $u8) | ConvertFrom-Json
}
$per = $datos['softys'].periodo
$ini = "$($per.mes)-01"; $iniPrev = "$($per.mesAnterior)-01"; $corte = $per.corte

# ---------------- oportunidades por cliente ----------------
$codigos = ($PROVEEDORES | ForEach-Object { "'$($_.codigo)'" }) -join ','
$porCli = @{}
foreach ($x in (Invoke-PanelSql @"
SELECT v.cliente AS cl, a.proveedor AS p,
  SUM(CASE WHEN v.fecha >= '$ini' THEN i.total ELSE 0 END) AS act,
  SUM(CASE WHEN v.fecha < '$ini' THEN i.total ELSE 0 END) AS ant
FROM venta_items i JOIN ventas v ON v.id = i.venta_id JOIN articulos a ON a.codigo = i.articulo
WHERE v.tipo = 'VEN' AND COALESCE(v.vendedor,'') NOT IN ($EXCL) AND v.fecha >= '$iniPrev' AND v.fecha <= '$corte' AND a.proveedor IN ($codigos)
GROUP BY v.cliente, a.proveedor
"@)) { $c = [string]$x.cl; if (-not $porCli.ContainsKey($c)) { $porCli[$c] = @{} }; $porCli[$c][[string]$x.p] = @{ act = D0 $x.act; ant = D0 $x.ant } }
$empresa = @{}; foreach ($x in (Invoke-PanelSql "SELECT v.cliente AS cl, SUM(v.total) AS t, MAX(v.vendedor) AS ve FROM ventas v WHERE v.tipo='VEN' AND COALESCE(v.vendedor,'') NOT IN ($EXCL) AND v.fecha >= '$ini' AND v.fecha <= '$corte' GROUP BY v.cliente")) { $empresa[[string]$x.cl] = $x }
$nomCli = @{}; foreach ($x in (Invoke-PanelSql "SELECT codigo, razon_social, vendedor FROM clientes")) { $nomCli[[string]$x.codigo] = $x }
$nomVend = @{}; foreach ($x in (Invoke-PanelSql "SELECT codigo, nombre FROM vendedores")) { $nomVend[[string]$x.codigo] = [string]$x.nombre }
function Vend([string]$cl) { $ve = if ($nomCli.ContainsKey($cl) -and $nomCli[$cl].vendedor) { [string]$nomCli[$cl].vendedor } elseif ($empresa.ContainsKey($cl)) { [string]$empresa[$cl].ve } else { '' }; if ($nomVend.ContainsKey($ve)) { "$ve $($nomVend[$ve])" } else { $ve } }
function NomCli([string]$cl) { if ($nomCli.ContainsKey($cl)) { [string]$nomCli[$cl].razon_social } else { "Cliente $cl" } }

$oportunidades = [ordered]@{}
foreach ($p in $PROVEEDORES) {
    $cod = $p.codigo; $otros = @($PROVEEDORES | Where-Object { $_.codigo -ne $cod } | ForEach-Object { $_.codigo })
    $recuperar = @(); $cruzada = @()
    foreach ($cl in $empresa.Keys) {
        $m = if ($porCli.ContainsKey($cl)) { $porCli[$cl] } else { @{} }
        $esteAct = if ($m.ContainsKey($cod)) { $m[$cod].act } else { 0 }; $esteAnt = if ($m.ContainsKey($cod)) { $m[$cod].ant } else { 0 }
        if ($esteAct -gt 0) { continue }
        $otrosAct = @($otros | Where-Object { $m.ContainsKey($_) -and $m[$_].act -gt 0 })
        if ($esteAnt -gt 0) { $recuperar += [pscustomobject]@{ cl = $cl; ant = $esteAnt; emp = D0 $empresa[$cl].t } }
        elseif ($otrosAct.Count -ge 2) { $cruzada += [pscustomobject]@{ cl = $cl; emp = D0 $empresa[$cl].t; otros = ($otrosAct | ForEach-Object { $c = $_; ($PROVEEDORES | Where-Object { $_.codigo -eq $c }).nombre }) -join ', ' } }
    }
    $lista = { param($rows, $orden, $tipo) @($rows | Sort-Object { -$_.$orden } | Select-Object -First 25 | ForEach-Object {
        $o = [ordered]@{ codigo = $_.cl; nombre = NomCli $_.cl; vendedor = Vend $_.cl; compraEmpresaMes = [math]::Round($_.emp) }
        if ($tipo -eq 'r') { $o.compraMesAnterior = [math]::Round($_.ant) } else { $o.compraOtros = $_.otros }; $o }) }
    $porVend = { param($rows) @($rows | Group-Object { Vend $_.cl } | Sort-Object { -$_.Count } | Select-Object -First 6 | ForEach-Object { [ordered]@{ vendedor = $_.Name; clientes = $_.Count } }) }
    $oportunidades[$p.clave] = [ordered]@{
        recuperar = [ordered]@{ clientes = $recuperar.Count; ventaMesAnterior = [math]::Round((($recuperar | Measure-Object ant -Sum).Sum)); lista = (& $lista $recuperar 'ant' 'r'); porVendedor = (& $porVend $recuperar) }
        cruzada = [ordered]@{ clientes = $cruzada.Count; lista = (& $lista $cruzada 'emp' 'c'); porVendedor = (& $porVend $cruzada) }
    }
}

# ---------------- recomendaciones ----------------
$rec = @()
function Rec($prov, $prio, $titulo, $detalle, $accion, $monto = 0, $lista = $null) { [ordered]@{ proveedor = $prov; prioridad = $prio; titulo = $titulo; detalle = $detalle; accion = $accion; monto = [math]::Round([double]$monto); lista = $lista } }
foreach ($p in $PROVEEDORES) {
    $d = $datos[$p.clave]; $n = $p.nombre; $A = $d.actual; $pe = $d.periodo; $op = $oportunidades[$p.clave]
    $tiempo = $pe.habilesTransc / $pe.habilesTotal; $rest = $pe.habilesTotal - $pe.habilesTransc
    if ($A.objEq) {
        $av = Div $A.venta $A.objEq; $falta = $A.objEq - $A.venta
        if ($falta -gt 0 -and $av -lt $tiempo) {
            $ritmo = $A.venta / $pe.habilesTransc; $nec = if ($rest -gt 0) { $falta / $rest } else { $falta }
            $rec += Rec $n $(if ((Div $av $tiempo) -lt 0.9) { 'alta' } else { 'media' }) "$($n): faltan $(Mill $falta) para el objetivo de equilibrio" "Lleva $(Pct $av) del objetivo con $(Pct $tiempo) del mes. Vende $(Mill $ritmo) por día hábil y necesita $(Mill $nec) en los $rest días que quedan." "Priorizar en la ruta los clientes de recuperar y venta cruzada de esta solapa; si la brecha es muy grande, negociar con $n una acción de cierre (bonificación por volumen o combo)." $falta
        }
    }
    if ($op.recuperar.clientes -gt 0) {
        $top = ($op.recuperar.porVendedor | Select-Object -First 3 | ForEach-Object { "$($_.vendedor) ($($_.clientes))" }) -join ', '
        $rec += Rec $n $(if ($op.recuperar.ventaMesAnterior -gt 0.1 * [math]::Max(1, $d.mesAnteriorCompleto.venta)) { 'alta' } else { 'media' }) "$($n): $($op.recuperar.clientes) clientes lo compraban y este mes no" "Ya compraron otra cosa a la empresa este mes, así que el vendedor los está visitando. El mes pasado le compraron $(Mill $op.recuperar.ventaMesAnterior) de $n. Más casos: $top." "Pasar la lista a cada vendedor para que ofrezca $n en la próxima visita; empezar por los de mayor compra del mes pasado." $op.recuperar.ventaMesAnterior "recuperar"
    }
    if ($op.cruzada.clientes -gt 0) {
        $top = ($op.cruzada.porVendedor | Select-Object -First 3 | ForEach-Object { "$($_.vendedor) ($($_.clientes))" }) -join ', '
        $rec += Rec $n 'media' "$($n): $($op.cruzada.clientes) clientes compran otros dos proveedores del tablero y no $n" "Compran este mes al menos dos de los otros proveedores y no le compraron a $n ni este mes ni el anterior. Más casos: $top." "Ofrecer un pedido de prueba con los 3 artículos que más venden de ${n}: $((@($d.productos.skus) | Select-Object -First 3 | ForEach-Object { $_.nombre }) -join ' · ')." (($op.cruzada.lista | Select-Object -First 10 | ForEach-Object { $_.compraEmpresaMes }) | Measure-Object -Sum).Sum "cruzada"
    }
    $al = $d.alertas
    if ($al.vendedores.estado -ne 'g') { $rec += Rec $n 'media' "$($n): vendedores con brecha" $al.vendedores.texto "Revisar con cada uno si es un cliente grande que no repitió o la ruta; acompañarlo una visita esta semana." }
    if ($al.productos.estado -ne 'g') { $rec += Rec $n 'media' "$($n): líneas que venden mucho con poco margen" $al.productos.texto "Revisar precio de lista y dinámicas de esas líneas; empujar en la venta las de mejor margen." }
    if ($al.upt.estado -ne 'g' -or $al.ticket.estado -ne 'g') { $rec += Rec $n 'media' "$($n): tickets más chicos o con menos artículos" ($al.ticket.texto + ' ' + $al.upt.texto) "Armar un surtido sugerido de $n (los más vendidos de cada línea) para que el vendedor sume un artículo más por pedido." }
    if ($al.conversion.estado -eq 'r') { $rec += Rec $n 'alta' "$($n): baja la conversión" $al.conversion.texto "Revisar los vendedores con menor conversión de la solapa y fijarles un mínimo de clientes con $n por semana." }
    if ($al.devoluciones.estado -eq 'r') { $rec += Rec $n 'alta' "$($n): suben las devoluciones" $al.devoluciones.texto "Revisar con logística los motivos de rechazo de $n de las últimas tres semanas." }
}
$orden = @{ alta = 0; media = 1 }
$rec = @($rec | Sort-Object @{ Expression = { $orden[$_.prioridad] } }, @{ Expression = { $_.monto }; Descending = $true })

$out = [ordered]@{
    generadoEn = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    periodo = $per
    solapas = @($PROVEEDORES | ForEach-Object { [ordered]@{ clave = $_.clave; nombre = $_.nombre; proveedor = $_.objetivo } })
    proveedores = $datos
    oportunidades = $oportunidades
    recomendaciones = $rec
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir | Out-Null }
[IO.File]::WriteAllText($OutPath, ($out | ConvertTo-Json -Depth 10), $u8)
if (Test-Path $TemplatePath) { Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force }
Write-Log ("Guardado: $OutPath | " + (($PROVEEDORES | ForEach-Object { "$($_.nombre) $(Mill $datos[$_.clave].actual.venta) recuperar $($oportunidades[$_.clave].recuperar.clientes) cruzada $($oportunidades[$_.clave].cruzada.clientes)" }) -join ' | ') + " | recomendaciones: $($rec.Count)")
