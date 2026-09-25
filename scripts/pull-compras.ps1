<#
Dashboard de Compras: cuanto facturaron los proveedores (neto de IVA) y que
% del total representa cada uno, guardando cada mes por separado (igual que
la Pizarra de Rentabilidad) para poder navegar meses hacia adelante.

Reglas de negocio:
  - Fuente: /data/cmd/compras/api/v1/get (comprobantes de compra de Gescom).
  - Se cuentan SOLO comprobantes tipo FAC-A y FAC-C (facturas de compra reales).
    Se descubrio (caso concreto: SOFTYS ARGENTINA S.A, sept 2026) que cada
    factura de compra aparece DUPLICADA en este endpoint como un "REMI-I"
    (remito interno) con el MISMO monto y los MISMOS items -- son la misma
    operacion registrada dos veces con distinta etiqueta. Sumar REMI/REMI-I/
    REME/REME-I junto con las facturas duplicaria el total al doble. Los
    REMI simples (sin "-I") ademas suelen venir en $0 (son solo el remito
    fisico del camion, no la factura).
  - NO se netean notas de credito (NCR-A/NCR-C) ni notas de debito (NDB-A)
    todavia: se encontraron con "codigoItem" generico (no un articulo real)
    y montos que no se pudieron verificar contra la pantalla de Gescom antes
    de esta primera version -- mejor mostrar "facturado bruto" con esta
    salvedad explicita que arriesgar un neteo mal hecho. Pendiente confirmar
    con la usuaria como tratarlas.
  - El filtro fechaDesde/fechaHasta de la API es poco preciso (devuelve un
    rango mas ancho del pedido, mismo comportamiento que ya se documento en
    la API de ventas) -- se pide con margen y se filtra despues por
    fechaComprobante real, client-side.
  - agrupa por proveedor; calcula % del total que representa cada proveedor.

Uso:
  powershell -File pull-compras.ps1                       # mes actual
  powershell -File pull-compras.ps1 -MesDesde 2026-07-01  # un mes especifico
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/compras"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "compras-template.html"),
    [string]$MesDesde = ""
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$TIPOS_FACTURA = @("FAC-A", "FAC-C")

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
    for ($intento = 1; $intento -le 4; $intento++) {
        try {
            return Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $script:token" } -Method Get
        } catch {
            $esAuthError = ($_.ErrorDetails.Message -match "Acceso denegado") -or ($_.Exception.Response.StatusCode.value__ -in 401, 403)
            if ($esAuthError) { $script:token = Get-GescomToken }
            if ($intento -eq 4) { throw }
            if (-not $esAuthError) { Start-Sleep -Seconds ($intento * 5) }
        }
    }
}
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }

Write-Log "Autenticando..."
$script:token = Get-GescomToken

Write-Log "Descargando catalogo de proveedores..."
$proveedores = Invoke-GescomApi -Path "/data/cmd/compras/api/v1/get-proveedores"
$provNombre = @{}
foreach ($p in $proveedores) { $provNombre[[string]$p.codigo] = $p.nombre }

if ($MesDesde -ne "") {
    $inicioMes = Get-Date $MesDesde
} else {
    $inicioMes = Get-Date -Day 1
}
$inicioMes = Get-Date -Year $inicioMes.Year -Month $inicioMes.Month -Day 1 -Hour 0 -Minute 0 -Second 0
$finMesCompleto = $inicioMes.AddMonths(1)
$hoy = (Get-Date).Date
$esMesActual = ($inicioMes.Year -eq $hoy.Year -and $inicioMes.Month -eq $hoy.Month)
$mesKey = $inicioMes.ToString("yyyy-MM")

$BUFFER_DIAS = 20
$fechaDesdeQuery = $inicioMes.AddDays(-$BUFFER_DIAS).ToString("yyyy-MM-dd")
$fechaHastaQuery = $finMesCompleto.AddDays($BUFFER_DIAS).ToString("yyyy-MM-dd")
Write-Log "Procesando mes $mesKey (comprobantes con fecha entre $($inicioMes.ToString('yyyy-MM-dd')) y $($finMesCompleto.AddDays(-1).ToString('yyyy-MM-dd')), consultando desde $fechaDesdeQuery hasta $fechaHastaQuery)..."

# 2 pasadas independientes, union por (proveedor+tipo+numero) -- mismo motivo
# que en pull-venta-rentabilidad.ps1: esta familia de API de Gescom ya
# demostro perder comprobantes reales de forma inconsistente entre corridas.
function Get-ComprobantesCalificados {
    param([int]$Intento)
    $script:token = Get-GescomToken
    $vistos = @{}
    # OJO: el parametro "page" NO pagina de verdad en este endpoint (page=1 y
    # page=2 devuelven exactamente lo mismo) -- el que si funciona es
    # "pagestoskip" (mismo nombre que la API de ventas). Sin este loop, una
    # ventana con mas de 500 comprobantes se trunca en silencio (se
    # confirmo perdiendo 64 facturas reales de septiembre con el buffer de
    # +/-20 dias antes de este fix).
    $pagestoskip = 0
    $totalLocal = 0
    while ($true) {
        if ($pagestoskip -gt 0 -and $pagestoskip % 5 -eq 0) { $script:token = Get-GescomToken }
        $page = Invoke-GescomApi -Path "/data/cmd/compras/api/v1/get" -Query @{
            fechaDesde = $fechaDesdeQuery; fechaHasta = $fechaHastaQuery; pagesize = 500; pagestoskip = $pagestoskip
        }
        if (-not $page -or $page.Count -eq 0) { break }
        $totalLocal += $page.Count
        foreach ($comp in $page) {
            if ($TIPOS_FACTURA -notcontains $comp.codigoTipoDeComprobante) { continue }
            if (-not $comp.fechaComprobante) { continue }
            $fecha = ([datetime]$comp.fechaComprobante).Date
            if ($fecha -lt $inicioMes -or $fecha -ge $finMesCompleto) { continue }
            $idUnico = "$($comp.codigoProveedor)|$($comp.codigoTipoDeComprobante)|$($comp.codigoPuntoVenta)|$($comp.numeroComprobante)"
            $vistos[$idUnico] = $comp
        }
        if ($page.Count -lt 500) { break }
        $pagestoskip++
    }
    Write-Log "  [intento $Intento] comprobantes recibidos: $totalLocal | facturas del mes: $($vistos.Count)"
    return $vistos
}

$PASADAS_UNION = 2
$union = @{}
for ($intento = 1; $intento -le $PASADAS_UNION; $intento++) {
    $vistos = Get-ComprobantesCalificados -Intento $intento
    foreach ($id in $vistos.Keys) { $union[$id] = $vistos[$id] }
}
Write-Log "Facturas de compra (union de $PASADAS_UNION pasadas): $($union.Count)"

$porProveedor = @{}
foreach ($comp in $union.Values) {
    $prov = [string]$comp.codigoProveedor
    if (-not $porProveedor.ContainsKey($prov)) { $porProveedor[$prov] = @{ importe = 0.0; comprobantes = 0 } }
    $importeComp = 0.0
    foreach ($it in $comp.items) { $importeComp += [double]$it.importeTotal }
    $porProveedor[$prov].importe += $importeComp
    $porProveedor[$prov].comprobantes += 1
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
