$ErrorActionPreference = "Stop"
$sp = if ($env:COMPRAS_TMP) { $env:COMPRAS_TMP } else { Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "tmp" }; if (-not (Test-Path $sp)) { New-Item -ItemType Directory $sp -Force | Out-Null }
$cfg = Get-Content (Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) "panel-config.json") -Raw | ConvertFrom-Json
$h = @{ Authorization = "Bearer " + $cfg.clave }
function Q($q) { $u = $cfg.baseUrl + "/consulta?sql=" + [uri]::EscapeDataString($q); $r = Invoke-RestMethod -Uri $u -Headers $h; if ($r.truncado) { Write-Host "AVISO truncado: $q" }; return $r.filas }
$hoy = (Get-Date).ToString("yyyy-MM-dd")
Write-Host "Articulos..."
$arts = Q "SELECT codigo, descripcion, proveedor, unidades_bulto, costo_bulto, bloqueado, marca FROM articulos"
Write-Host "  $($arts.Count)"
$provs = Q "SELECT codigo, nombre FROM proveedores"
Write-Host "Stock (deposito PRI)..."
$stock = Q "SELECT articulo, SUM(cantidad) AS cantidad FROM stock WHERE deposito = 'PRI' GROUP BY articulo"
Write-Host "  $($stock.Count)"
Write-Host "Ventas desde 3/8 por articulo, tipo y mes..."
$ventas = Q "SELECT i.articulo AS articulo, v.tipo AS tipo, substr(v.fecha,1,7) AS mes, SUM(i.cantidad * i.factor) AS unidades, SUM(i.neto) AS neto, MAX(v.fecha) AS ultima FROM venta_items i JOIN ventas v ON v.id = i.venta_id WHERE v.fecha >= '2026-08-03' AND v.tipo IN ('VEN','DEB','DEV-RE') GROUP BY i.articulo, v.tipo, substr(v.fecha,1,7)"
Write-Host "  $($ventas.Count)"
Write-Host "Compras por articulo..."
$compras = Q "SELECT i.articulo AS articulo, MAX(c.carga) AS ultima, SUM(CASE WHEN c.carga >= '2026-10-01' THEN i.cantidad * i.factor ELSE 0 END) AS u_mes, SUM(i.cantidad * i.factor) AS u_total FROM compra_items i JOIN compras c ON c.clave = i.compra WHERE c.tipo NOT LIKE 'NCR%' AND c.tipo NOT LIKE 'NDB%' AND c.tipo NOT LIKE 'LIQ%' GROUP BY i.articulo"
Write-Host "  $($compras.Count)"
Write-Host "Dias con operacion en la base..."
$dias = Q "SELECT fecha, COUNT(*) AS n FROM ventas WHERE tipo = 'VEN' AND fecha >= '2026-08-03' GROUP BY fecha ORDER BY fecha"
Write-Host "  $($dias.Count)"
$out = [ordered]@{ hoy = $hoy; articulos = $arts; proveedores = $provs; stock = $stock; ventas = $ventas; compras = $compras; dias = $dias }
[System.IO.File]::WriteAllText("$sp\compras-base.json", (ConvertTo-Json -InputObject $out -Depth 5 -Compress), (New-Object System.Text.UTF8Encoding $false))
Write-Host "Guardado compras-base.json"
