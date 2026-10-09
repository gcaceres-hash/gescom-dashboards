<#
Rastreador de precios mayoristas -- solapa "80% de la venta".

Para los articulos que hacen el 80% de la venta neta de los ultimos 30 dias (Lago Puelo + Elebes)
compara, por unidad y con IVA:
  - nuestra Lista 1 (y Lista 3) de la tabla precios de la base compartida,
  - el precio CON DINAMICA: lo que efectivamente se cobro en esos 30 dias (total / unidades), o sea
    la lista menos las dinamicas y descuentos que aplico Gescom. Tambien el escalon mas bajo cobrado
    ("mejor dinamica"), ignorando bonificaciones (lineas a menos del 60% de la lista),
  - el precio mas bajo de los folletos/revistas de Nini, Vital, Yaguar, Diarco, Maxiconsumo, Belen,
    Carrefour Maxi, Tornado y Lamadrid (scripts/rastreador/mercado.json).

Reglas:
  - Venta = comprobantes VEN de los ultimos 30 dias; se excluyen los vendedores 1176, 43, 16 y 37
    (no son vendedores reales).
  - mercado.json sale del Rastreador de Lucas (ventas.panelempresas.workers.dev/panel/rastreador);
    sus codigos son "LPE-<codigo Gescom>" y trae cuantas unidades tiene el bulto (k) para llevar
    nuestro precio a precio por unidad como en el folleto. Se refresca a mano en el ciclo
    "actualizame todo" (requiere la sesion de Gisela en ese panel).
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/rastreador"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "rastreador-template.html"),
    [string]$MercadoPath = (Join-Path $PSScriptRoot "rastreador/mercado.json"),
    [int]$Dias = 30,
    [double]$Corte = 0.80
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$EXCLUDED = @("43","043","1176","16","016","37","037")
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
function Invoke-PanelSql([string]$Sql) {
    $uri = "$($config.baseUrl)/consulta?sql=" + [uri]::EscapeDataString($Sql)
    $headers = @{ Authorization = "Bearer $($config.clave)" }
    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    if ($r.truncado) { Write-Log "AVISO: la consulta vino truncada (mas de 20000 filas) -- revisar y acotar." }
    return $r.filas
}
function R2($x) { return [math]::Round([double]$x, 2) }
$EMPRESAS = @{ "1"="Elebes"; "3"="Lago Puelo" }
$excl = ($EXCLUDED | ForEach-Object { "'$_'" }) -join ","

# --- 1) venta por articulo, ultimos $Dias dias ---
Write-Log "Descargando venta de los ultimos $Dias dias por articulo..."
$desde = (Get-Date).AddDays(-$Dias).ToString("yyyy-MM-dd")
$venta = Invoke-PanelSql @"
SELECT i.articulo AS c, a.descripcion AS d, a.empresa AS e, p.nombre AS pv,
       SUM(i.neto) AS neto, SUM(i.total) AS total, SUM(i.cantidad * COALESCE(i.factor,1)) AS uni,
       SUM(i.precio_costo * i.cantidad * COALESCE(i.factor,1)) AS costo
FROM ventas v JOIN venta_items i ON i.venta_id = v.id
LEFT JOIN articulos a ON a.codigo = i.articulo
LEFT JOIN proveedores p ON p.codigo = a.proveedor
WHERE v.tipo = 'VEN' AND v.fecha >= '$desde' AND COALESCE(v.vendedor,'') NOT IN ($excl) AND i.total > 0
GROUP BY i.articulo ORDER BY neto DESC
"@
$totalNeto = ($venta | Measure-Object neto -Sum).Sum
$top = @(); $acum = 0.0
foreach ($v in $venta) {
    if ($acum / $totalNeto -ge $Corte) { break }
    $acum += [double]$v.neto
    $v | Add-Member -NotePropertyName acum -NotePropertyValue ($acum / $totalNeto)
    $top += $v
}
Write-Log "Articulos con venta: $($venta.Count) | venta neta $([math]::Round($totalNeto)) | el $([math]::Round($Corte*100))% lo hacen $($top.Count) articulos"
$codigos = ($top | ForEach-Object { "'" + ([string]$_.c).Replace("'","''") + "'" }) -join ","

# --- 2) listas de precio (unidad base) ---
$precios = Invoke-PanelSql "SELECT articulo, lista, neto, alicuota_iva, impuesto_interno FROM precios WHERE articulo IN ($codigos) AND lista IN ('1','3') AND factor = 1"
$lista = @{}
foreach ($p in $precios) {
    $iva = if ($null -ne $p.alicuota_iva) { [double]$p.alicuota_iva } else { 0.21 }
    $lista["$($p.articulo)|$($p.lista)"] = [double]$p.neto * (1 + $iva) + [double]$p.impuesto_interno
}

# --- 3) escalon mas bajo cobrado (mejor dinamica) y unidades vendidas con descuento ---
# (agregado en la base: linea a linea son mas de 20000 filas y la consulta se trunca)
# las lineas a menos del 60% de la lista son bonificaciones / regalos, no una dinamica de precio
$lineas = Invoke-PanelSql @"
SELECT c, MIN(CASE WHEN pu >= 0.6 * lp THEN pu END) AS mn,
       SUM(CASE WHEN pu >= 0.6 * lp AND pu < 0.995 * lp THEN u ELSE 0 END) AS ud
FROM (SELECT i.articulo AS c, i.total / (i.cantidad * COALESCE(i.factor,1)) AS pu, i.cantidad * COALESCE(i.factor,1) AS u,
             pr.neto * (1 + COALESCE(pr.alicuota_iva,0.21)) + COALESCE(pr.impuesto_interno,0) AS lp
      FROM ventas v JOIN venta_items i ON i.venta_id = v.id
      JOIN precios pr ON pr.articulo = i.articulo AND pr.lista = '1' AND pr.factor = 1
      WHERE v.tipo = 'VEN' AND v.fecha >= '$desde' AND COALESCE(v.vendedor,'') NOT IN ($excl) AND i.total > 0 AND i.cantidad > 0 AND i.articulo IN ($codigos))
GROUP BY c
"@
$minPu = @{}; $uDesc = @{}
foreach ($l in $lineas) {
    if ($null -ne $l.mn) { $minPu[[string]$l.c] = [double]$l.mn }
    $uDesc[[string]$l.c] = [double]$l.ud
}

# --- 4) precios de folletos ---
$mercado = [System.IO.File]::ReadAllText($MercadoPath, (New-Object System.Text.UTF8Encoding $false)) | ConvertFrom-Json
$mk = @{}
foreach ($f in $mercado.filas) { $mk[([string]$f[0]) -replace '^LPE-',''] = $f }

# --- 5) armar filas ---
$rows = @()
foreach ($v in $top) {
    $c = [string]$v.c
    $f = $mk[$c]
    $k = if ($f) { [double]$f[2] } else { 1.0 }
    $uni = [double]$v.uni
    $l1 = $lista["$c|1"]; $l3 = $lista["$c|3"]
    $din = if ($uni -gt 0) { [double]$v.total / $uni } else { $null }
    $m = @()
    if ($f) { foreach ($x in $f[3]) { $m += [ordered]@{ f = [string]$x[0]; p = R2 $x[1]; t = [string]$x[2]; n = [string]$x[3]; u = [string]$x[4] } } }
    $ivaF = if ($l1 -and $uni -gt 0 -and [double]$v.neto -gt 0) { [double]$v.total / [double]$v.neto } else { 1.21 }
    $rows += [ordered]@{
        c = $c; d = [string]$v.d; e = $(if ($EMPRESAS.ContainsKey([string]$v.e)) { $EMPRESAS[[string]$v.e] } else { "Empresa " + $v.e }); pv = [string]$v.pv
        k = $k
        l1 = $(if ($l1) { R2 ($l1 / $k) } else { $null })
        l3 = $(if ($l3) { R2 ($l3 / $k) } else { $null })
        din = $(if ($din) { R2 ($din / $k) } else { $null })
        dinMin = $(if ($minPu.ContainsKey($c)) { R2 ($minPu[$c] / $k) } else { $null })
        pctUniDesc = $(if ($uni -gt 0) { R2 ([double]$uDesc[$c] / $uni) } else { 0 })
        cu = $(if ($uni -gt 0) { R2 ([double]$v.costo * $ivaF / $uni / $k) } else { $null })
        n30 = R2 $v.neto; u30 = [math]::Round($uni); acum = [math]::Round([double]$v.acum, 4)
        m = $m
    }
}
$conM = @($rows | Where-Object { $_.m.Count -gt 0 }).Count
Write-Log "Articulos del 80% con precio de folleto: $conM de $($rows.Count)"

$out = [ordered]@{
    generadoEn = (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
    ventana = [ordered]@{ desde = $desde; dias = $Dias; corte = $Corte; ventaNetaTotal = R2 $totalNeto; articulosConVenta = $venta.Count }
    mercado = [ordered]@{ generado = [string]$mercado.generado; fuentes = $mercado.fuentes }
    rows = $rows
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
[System.IO.File]::WriteAllText($OutPath, ($out | ConvertTo-Json -Depth 8 -Compress), (New-Object System.Text.UTF8Encoding $false))
Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force
Write-Log "Guardado: $OutPath"
