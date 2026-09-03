<#
Genera stock-dashboard.html con el estado de stock de todas las empresas,
apertura por proveedor y ranking de SKUs mas vendidos / sin movimiento (30 dias).
Pensado para correr cada hora desde el Programador de tareas de Windows.

Uso manual:
  powershell -File stock-dashboard-refresh.ps1
#>
param(
    [int]$PeriodoDias = 30,
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "stock-dashboard-template.html"),
    [string]$OutputPath = (Join-Path $PSScriptRoot "stock-dashboard.html"),
    [string]$ShareTemplatePath = (Join-Path $PSScriptRoot "stock-dashboard-artifact-template.html"),
    [string]$ShareOutputPath = (Join-Path $PSScriptRoot "stock-dashboard-share.html"),
    [string]$LogPath = (Join-Path $PSScriptRoot "stock-dashboard-refresh.log")
)

$ErrorActionPreference = "Stop"

function Write-Log {
    param([string]$Mensaje)
    $linea = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Mensaje"
    Write-Host $linea
    Add-Content -Path $LogPath -Value $linea
}

try {
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

    function Get-GescomToken {
        $tokenUrl = "$($config.authUrl)/realms/$($config.realm)/protocol/openid-connect/token"
        $body = @{ client_id = $config.clientId; username = $config.usuario; password = $config.clave; grant_type = "password" }
        (Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded").access_token
    }
    function Invoke-GescomApi {
        param([string]$Path, [hashtable]$Query = @{})
        $qs = ($Query.GetEnumerator() | ForEach-Object { "$($_.Key)=$([uri]::EscapeDataString([string]$_.Value))" }) -join "&"
        $url = "$($config.baseUrl)$Path"
        if ($qs) { $url += "?$qs" }
        Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $script:token" } -Method Get
    }

    Write-Log "Autenticando contra Gescom ($($config.realm))..."
    $script:token = Get-GescomToken

    Write-Log "Descargando empresas, proveedores, articulos y stock..."
    $empresasRaw    = Invoke-GescomApi -Path "/data/cmd/config/api/v1/get-empresas"
    $proveedoresRaw = Invoke-GescomApi -Path "/data/cmd/compras/api/v1/get-proveedores"
    $articulosRaw   = Invoke-GescomApi -Path "/data/cmd/inventario/api/v2/get-articulos"
    $stockRaw       = Invoke-GescomApi -Path "/data/cmd/inventario/api/v1/get-stock" -Query @{ includeZero = "false"; disponibleVentas = "true" }

    $proveedoresPorCodigo = @{}
    foreach ($p in $proveedoresRaw) { $proveedoresPorCodigo[[string]$p.codigo] = $p.nombre }

    $articulosPorCodigo = @{}
    foreach ($a in $articulosRaw) { $articulosPorCodigo[[string]$a.codigo] = $a }

    $stockPorItem = @{}
    foreach ($s in $stockRaw) {
        $total = 0.0
        foreach ($d in $s.stock) { $total += [double]$d.cantidad }
        $stockPorItem[[string]$s.codigoItem] = $total
    }

    $fechaDesde = (Get-Date).AddDays(-$PeriodoDias).ToString("yyyy-MM-dd")
    $fechaHasta = (Get-Date).AddDays(1).ToString("yyyy-MM-dd")

    Write-Log "Descargando ventas de los ultimos $PeriodoDias dias ($fechaDesde a $fechaHasta)..."
    $ventasNetasPorItem = @{}
    $pagestoskip = 0
    $pagesize = 500
    $totalVentas = 0
    while ($true) {
        $page = Invoke-GescomApi -Path "/data/cmd/ventas/api/v2/get" -Query @{
            fechadesde  = $fechaDesde
            fechahasta  = $fechaHasta
            pagesize    = $pagesize
            pagestoskip = $pagestoskip
        }
        if (-not $page -or $page.Count -eq 0) { break }
        $totalVentas += $page.Count
        foreach ($venta in $page) {
            if ($venta.cerrada -ne $true) { continue }
            $signo = if ($venta.codigoTipoVenta -like "DEV*") { -1.0 } else { 1.0 }
            foreach ($item in $venta.items) {
                $codigo = [string]$item.codigoItem
                if (-not $ventasNetasPorItem.ContainsKey($codigo)) { $ventasNetasPorItem[$codigo] = 0.0 }
                $ventasNetasPorItem[$codigo] += ($signo * [double]$item.cantidad)
            }
        }
        if ($page.Count -lt $pagesize) { break }
        $pagestoskip++
    }
    Write-Log "Ventas descargadas: $totalVentas (cerrada=true, neto de devoluciones)"

    Write-Log "Calculando dataset de articulos con stock..."
    $articulosDataset = foreach ($codigo in $stockPorItem.Keys) {
        $art = $articulosPorCodigo[$codigo]
        if (-not $art) { continue }
        $unidades = [math]::Round($stockPorItem[$codigo], 2)
        if ($unidades -le 0) { continue }
        $costo = if ($art.precioCosto) { [double]$art.precioCosto } else { 0.0 }
        $vendido = if ($ventasNetasPorItem.ContainsKey($codigo)) { [math]::Round($ventasNetasPorItem[$codigo], 2) } else { 0.0 }
        if ($vendido -lt 0) { $vendido = 0.0 }
        $bulto = if ($art.unidadesPorBulto -and $art.unidadesPorBulto -gt 1) { [int]$art.unidadesPorBulto } else { 1 }
        [pscustomobject]@{
            codigo            = $codigo
            descripcion       = $art.descripcion
            empresaCodigo     = [string]$art.codigoEmpresa
            proveedorCodigo   = [string]$art.codigoProveedor
            stockUnidades     = $unidades
            precioCosto       = [math]::Round($costo, 4)
            valorStock        = [math]::Round($unidades * $costo, 2)
            vendidoUnidades   = $vendido
            unidadesPorBulto  = $bulto
        }
    }

    $empresasUsadas = $articulosDataset | Select-Object -ExpandProperty empresaCodigo -Unique
    $empresas = foreach ($codigo in $empresasUsadas) {
        if ([string]::IsNullOrWhiteSpace($codigo)) {
            [pscustomobject]@{ codigo = $codigo; nombre = "Sin empresa asignada" }
            continue
        }
        $nombre = ($empresasRaw | Where-Object { [string]$_.codigo -eq $codigo } | Select-Object -First 1 -ExpandProperty nombre)
        [pscustomobject]@{ codigo = $codigo; nombre = if ($nombre) { $nombre } else { "Empresa $codigo" } }
    }

    $proveedoresUsados = $articulosDataset | Select-Object -ExpandProperty proveedorCodigo -Unique
    $proveedores = foreach ($codigo in $proveedoresUsados) {
        $nombre = $proveedoresPorCodigo[$codigo]
        [pscustomobject]@{ codigo = $codigo; nombre = if ($nombre) { $nombre } else { "Proveedor $codigo" } }
    }

    $dataset = [pscustomobject]@{
        generatedAt = (Get-Date).ToString("o")
        periodDays  = $PeriodoDias
        empresas    = @($empresas | Sort-Object nombre)
        proveedores = @($proveedores | Sort-Object nombre)
        articulos   = @($articulosDataset)
    }

    $json = $dataset | ConvertTo-Json -Depth 6 -Compress

    Write-Log "SKUs con stock: $($articulosDataset.Count) | Empresas: $($empresas.Count) | Proveedores: $($proveedores.Count)"

    if (-not (Test-Path $TemplatePath)) { throw "No se encuentra el template: $TemplatePath" }
    $template = Get-Content $TemplatePath -Raw -Encoding UTF8

    $generadoTexto = (Get-Date).ToString("dd/MM/yyyy HH:mm")
    $html = $template.Replace("__DASHBOARD_JSON__", $json).Replace("__GENERATED_AT_TEXT__", $generadoTexto)

    Set-Content -Path $OutputPath -Value $html -Encoding UTF8
    Write-Log "Dashboard generado: $OutputPath"

    if (Test-Path $ShareTemplatePath) {
        $shareTemplate = Get-Content $ShareTemplatePath -Raw -Encoding UTF8
        $shareHtml = $shareTemplate.Replace("__DASHBOARD_JSON__", $json).Replace("__GENERATED_AT_TEXT__", $generadoTexto)
        Set-Content -Path $ShareOutputPath -Value $shareHtml -Encoding UTF8
        Write-Log "Version para compartir generada: $ShareOutputPath"
    }
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    throw
}
