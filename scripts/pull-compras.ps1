<#
Dashboard de Compras: cuanto facturaron los proveedores (neto de IVA) y que
% del total representa cada uno, guardando cada mes por separado (igual que
la Pizarra de Rentabilidad) para poder navegar meses hacia adelante.

MIGRADO el 30/9/2026 de la API de Gescom a la base compartida de Lucas
(datos-gescom.panelempresas.workers.dev) -- ver pull-cobertura-general.ps1
para el detalle de por que.

Reglas de negocio:
  - Fuente: compras_arca + compra_netos (desglose armado por Lucas para la
    conciliacion con ARCA), NO la tabla `compras` simple -- esa solo tiene el
    total CON impuesto. Se decidio con la usuaria (30/9/2026) mantener el
    mismo criterio historico ("compra sin impuesto") aunque compras_arca
    cubra ~95% de las facturas del mes (se sincroniza con unos dias de
    demora respecto a `compras`) en vez de cambiar a un numero con impuesto
    con cobertura completa.
  - Se cuentan SOLO comprobantes tipo FAC-A y FAC-C (facturas de compra
    reales). Los remitos (REMI/REMI-I/REME/REME-I) y notas de credito/debito
    (NCR-A/NCR-C/NDB-A/NDB-C) quedan afuera -- las notas de credito/debito NO
    se netean todavia (misma salvedad que la version anterior: pendiente
    confirmar el tratamiento con la usuaria).
  - Solo se muestran proveedores de MERCADERIA (a pedido explicito de la
    usuaria): se excluyen los que no tienen NINGUN articulo cargado a su
    nombre en el catalogo de inventario (tabla articulos), señal objetiva de
    que no venden mercaderia que se revenda.
  - agrupa por proveedor; calcula % del total que representa cada proveedor.

Uso:
  powershell -File pull-compras.ps1                       # mes actual
  powershell -File pull-compras.ps1 -MesDesde 2026-07-01  # un mes especifico
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/compras"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "compras-template.html"),
    [string]$MesDesde = ""
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$TIPOS_FACTURA = @("FAC-A", "FAC-C")

function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
function Invoke-PanelSql([string]$Sql) {
    $uri = "$($config.baseUrl)/consulta?sql=" + [uri]::EscapeDataString($Sql)
    $headers = @{ Authorization = "Bearer $($config.clave)" }
    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    if ($r.truncado) { Write-Log "AVISO: la consulta vino truncada (mas de 20000 filas) -- revisar y acotar." }
    return $r.filas
}

Write-Log "Descargando catalogo de proveedores y articulos..."
$proveedores = Invoke-PanelSql "SELECT codigo, nombre FROM proveedores"
$provNombre = @{}
foreach ($p in $proveedores) { $provNombre[[string]$p.codigo] = $p.nombre }
$articulos = Invoke-PanelSql "SELECT DISTINCT proveedor FROM articulos WHERE proveedor IS NOT NULL AND proveedor <> ''"
$provsMercaderia = @{}
foreach ($a in $articulos) { $provsMercaderia[[string]$a.proveedor] = $true }
Write-Log "Proveedores con articulos en el catalogo (mercaderia real): $($provsMercaderia.Count)"

if ($MesDesde -ne "") { $inicioMes = Get-Date $MesDesde } else { $inicioMes = Get-Date -Day 1 }
$inicioMes = Get-Date -Year $inicioMes.Year -Month $inicioMes.Month -Day 1 -Hour 0 -Minute 0 -Second 0
$finMesCompleto = $inicioMes.AddMonths(1)
$hoy = (Get-Date).Date
$esMesActual = ($inicioMes.Year -eq $hoy.Year -and $inicioMes.Month -eq $hoy.Month)
$mesKey = $inicioMes.ToString("yyyy-MM")
$fechaDesde = $inicioMes.ToString("yyyy-MM-dd")
$fechaHasta = $finMesCompleto.ToString("yyyy-MM-dd")
Write-Log "Procesando mes $mesKey ($fechaDesde a $fechaHasta)..."

$tiposIn = ($TIPOS_FACTURA | ForEach-Object { "'$_'" }) -join ","
$sql = @"
SELECT ca.proveedor AS proveedor, COUNT(DISTINCT ca.id) AS facturas, ROUND(SUM(cn.neto),2) AS neto
FROM compras_arca ca
JOIN compra_netos cn ON cn.compra_id = ca.id
WHERE ca.tipo IN ($tiposIn) AND ca.fecha >= '$fechaDesde' AND ca.fecha < '$fechaHasta'
GROUP BY ca.proveedor
"@
$filas = Invoke-PanelSql $sql
Write-Log "Proveedores con factura de compra (antes del filtro de mercaderia): $($filas.Count)"

$porProveedor = @{}
foreach ($f in $filas) {
    $prov = [string]$f.proveedor
    if (-not $provsMercaderia.ContainsKey($prov)) { continue }
    $porProveedor[$prov] = @{ importe = [double]$f.neto; comprobantes = [int]$f.facturas }
}

$totalGeneral = 0.0
foreach ($v in $porProveedor.Values) { $totalGeneral += $v.importe }

$proveedoresOut = foreach ($cod in $porProveedor.Keys) {
    $p = $porProveedor[$cod]
    $nombre = $provNombre[$cod]
    if (-not $nombre) { $nombre = "Proveedor $cod" }
    [pscustomobject]@{
        codigo = $cod
        nombre = $nombre
        importe = [math]::Round($p.importe, 2)
        comprobantes = $p.comprobantes
        pct = if ($totalGeneral -ne 0) { [math]::Round(100.0 * $p.importe / $totalGeneral, 2) } else { 0.0 }
    }
}

$mesData = [pscustomobject]@{
    periodoDesde = $inicioMes.ToString("yyyy-MM-dd")
    periodoHasta = $(if ($esMesActual) { $hoy.ToString("yyyy-MM-dd") } else { $finMesCompleto.AddDays(-1).ToString("yyyy-MM-dd") })
    cerrado = -not $esMesActual
    totalCompraSinImpuesto = [math]::Round($totalGeneral, 2)
    proveedores = @($proveedoresOut | Sort-Object importe -Descending)
}

$meses = [ordered]@{}
try {
    if (Test-Path $OutPath) {
        $rawText = [System.IO.File]::ReadAllText($OutPath, [System.Text.Encoding]::UTF8)
        $prev = $rawText | ConvertFrom-Json
        if ($prev.meses) {
            foreach ($prop in $prev.meses.PSObject.Properties) { $meses[$prop.Name] = $prop.Value }
            Write-Log "Historico cargado desde $OutPath : $($meses.Keys -join ', ')"
        }
    } else {
        Write-Log "No existe data.json previo en $OutPath, se arranca sin historico."
    }
} catch {
    Write-Log "Aviso: no se pudo leer el data.json existente, se arranca sin historico previo ($($_.Exception.Message))"
}
foreach ($k in @($meses.Keys)) {
    if ($k -ne $mesKey -and -not $meses[$k].cerrado) { $meses[$k] | Add-Member -NotePropertyName cerrado -NotePropertyValue $true -Force }
}
$meses[$mesKey] = $mesData

$out = [pscustomobject]@{
    generatedAt = (Get-Date).ToString("o")
    mesActual = (Get-Date -Format "yyyy-MM")
    meses = $meses
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
$jsonText = $out | ConvertTo-Json -Depth 10 -Compress
[System.IO.File]::WriteAllText($OutPath, $jsonText, (New-Object System.Text.UTF8Encoding $false))
Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force
Write-Log "Guardado: $OutPath (meses en historico: $($meses.Keys -join ', '))"
Write-Log "Mes $mesKey -> Compra sin impuesto: $($mesData.totalCompraSinImpuesto) | Proveedores: $($proveedoresOut.Count)"
