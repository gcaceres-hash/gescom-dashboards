<#
Descarga la cuenta corriente de clientes de Gescom, resuelve cada comprobante
pendiente a su vendedor (via ventaId) y arma la estructura:
  Responsable de cobro -> Clientes -> (Comprobantes + apertura por vendedor)

Reglas de negocio (confirmadas con el usuario):
  - Vendedores 1176 y 43 (ONCE SETENTA Y SEIS): no son reales, se excluyen SIEMPRE.
  - Vendedor 16 (VENDEDOR DEPOSITO) y 37 (LOGISTICA): se excluyen tambien.
  - Vendedores 076 (Martin Di Giorno), 038 (Gisela Caceres), 050 (Guillermo
    Zeballos): son mayoristas, se gestionan solos (responsable = ellos mismos).
  - Codigo de vendedor de 3 digitos (ej "063"): responsable = Bruno.
  - Cualquier otro codigo (1-2 digitos, ej "7", "28"): responsable = Johana.
  - Comprobantes sin ventaId (deuda por otro motivo): responsable = "Sin vendedor asignado".
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string]$OutPath = (Join-Path $PSScriptRoot "cuentas-corrientes-data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "cuentas-corrientes-template.html"),
    [string]$NetlifyToken = $env:NETLIFY_TOKEN,
    [string]$NetlifySiteId = "980d2e04-124d-41aa-b4c0-a7222bffe82d",
    [string]$NetlifySiteUrl = "https://cobranzas-fusionlpelebes.netlify.app"
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

function Get-GescomToken {
    $tokenUrl = "$($config.authUrl)/realms/$($config.realm)/protocol/openid-connect/token"
    $body = @{ client_id = $config.clientId; username = $config.usuario; password = $config.clave; grant_type = "password" }
    (Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded").access_token
}
function Invoke-GescomApi {
    param([string]$Path, [string]$Query = "")
    $url = "$($config.baseUrl)$Path"
    if ($Query) { $url += "?$Query" }
    Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $script:token" } -Method Get
}
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }

$EXCLUDED = @("43","1176","16","37")
$MAYORISTA_SOLO = @{ "076" = "Mart$([char]0xED)n Di Giorno"; "038" = "Gisela Caceres"; "050" = "Guillermo Zeballos" }

function Get-Responsable($codigoVendedor) {
    if (-not $codigoVendedor) { return "_SIN_VENDEDOR_" }
    if ($EXCLUDED -contains $codigoVendedor) { return $null }
    if ($MAYORISTA_SOLO.ContainsKey($codigoVendedor)) { return $codigoVendedor }
    if ($codigoVendedor.Length -eq 3) { return "_BRUNO_" }
    return "_JOHANA_"
}

Write-Log "Autenticando..."
$script:token = Get-GescomToken

Write-Log "Descargando clientes, vendedores y cuenta corriente detalle..."
$clientes = Invoke-GescomApi -Path "/data/cmd/ventas/api/v1/get-clientes"
$vendedores = Invoke-GescomApi -Path "/data/cmd/ventas/api/v1/get-vendedores"
$detalle = Invoke-GescomApi -Path "/data/cmd/ctacte/api/v3/get-ctacte-clientes-detalle"
Write-Log "Clientes: $($clientes.Count) | Vendedores: $($vendedores.Count) | Cuentas corrientes: $($detalle.Count)"

$clienteNombre = @{}
foreach ($c in $clientes) { $clienteNombre[[string]$c.codigo] = $c.nombre }
$vendedorNombre = @{}
foreach ($v in $vendedores) { $vendedorNombre[[string]$v.codigo] = $v.nombre }

Write-Log "Resolviendo ventaId -> vendedor..."
$ventaIds = New-Object System.Collections.Generic.HashSet[string]
foreach ($cli in $detalle) { foreach ($c in $cli.comprobantes) { if ($c.ventaId) { [void]$ventaIds.Add([string]$c.ventaId) } } }
$ventaIdsList = @($ventaIds)
$ventaVendedor = @{}
$batchSize = 50
for ($i = 0; $i -lt $ventaIdsList.Count; $i += $batchSize) {
    if ($i -gt 0 -and ($i / $batchSize) % 10 -eq 0) { $script:token = Get-GescomToken }
    $batch = $ventaIdsList[$i..([math]::Min($i+$batchSize-1, $ventaIdsList.Count-1))]
    $r = Invoke-GescomApi -Path "/data/cmd/ventas/api/v2/get" -Query "ids=$($batch -join ',')&pagesize=50"
    foreach ($v in $r) { $ventaVendedor[[string]$v.id] = [string]$v.codigoVendedor }
}
Write-Log "Ventas resueltas: $($ventaVendedor.Count) / $($ventaIdsList.Count)"

$hoy = (Get-Date).Date
function Bucket($diasVencido) {
    if ($diasVencido -le 0) { return "vigente" }
    if ($diasVencido -le 30) { return "d1_30" }
    if ($diasVencido -le 60) { return "d31_60" }
    if ($diasVencido -le 90) { return "d61_90" }
    return "d90mas"
}

# estructura: responsables[respKey] = { clientes[codigoCliente] = { comprobantes=[...] } }
$responsables = @{}
$excluidoTotal = 0.0
$excluidoDetalle = @{}

foreach ($cli in $detalle) {
    $codCliente = [string]$cli.codigoCliente
    foreach ($c in $cli.comprobantes) {
        $vendCod = if ($c.ventaId -and $ventaVendedor.ContainsKey([string]$c.ventaId)) { $ventaVendedor[[string]$c.ventaId] } else { $null }
        $resp = Get-Responsable $vendCod
        $esCredito = $null -ne $c.creditoId
        $saldo = [double]$c.saldo * $(if ($esCredito) { -1 } else { 1 })
        if (-not $resp) {
            $excluidoTotal += $saldo
            if (-not $excluidoDetalle.ContainsKey($vendCod)) { $excluidoDetalle[$vendCod] = 0.0 }
            $excluidoDetalle[$vendCod] += $saldo
            continue
        }
        $fv = if ($c.fechaVencimiento) { [datetime]$c.fechaVencimiento } else { $hoy }
        $diasVencido = ($hoy - $fv.Date).Days
        $bucket = Bucket $diasVencido

        if (-not $responsables.ContainsKey($resp)) { $responsables[$resp] = @{} }
        if (-not $responsables[$resp].ContainsKey($codCliente)) { $responsables[$resp][$codCliente] = @() }
        $responsables[$resp][$codCliente] += [pscustomobject]@{
            comprobante = $c.comprobante
            saldo = [math]::Round($saldo,2)
            esCredito = $esCredito
            fechaEmision = if ($c.fechaEmision) { ([datetime]$c.fechaEmision).ToString("yyyy-MM-dd") } else { $null }
            fechaVencimiento = if ($c.fechaVencimiento) { $fv.ToString("yyyy-MM-dd") } else { $null }
            diasVencido = $diasVencido
            bucket = $bucket
            codigoVendedor = $vendCod
            nombreVendedor = if ($vendCod) { $vendedorNombre[$vendCod] } else { $null }
            codigoEmpresa = $c.codigoEmpresa
        }
    }
}

function NombreResponsable($key) {
    switch ($key) {
        "_BRUNO_" { return "Bruno" }
        "_JOHANA_" { return "Johana" }
        "_SIN_VENDEDOR_" { return "Sin vendedor asignado" }
        default { return $MAYORISTA_SOLO[$key] }
    }
}

Write-Log "Armando estructura final..."
$responsablesOut = foreach ($respKey in $responsables.Keys) {
    $clientesOut = foreach ($codCliente in $responsables[$respKey].Keys) {
        $comps = $responsables[$respKey][$codCliente]
        $saldoCliente = [math]::Round((($comps | Measure-Object saldo -Sum).Sum),2)
        $peorDias = ($comps | Measure-Object diasVencido -Maximum).Maximum
        $vendedoresCliente = $comps | Group-Object codigoVendedor | ForEach-Object {
            [pscustomobject]@{
                codigo = $_.Name
                nombre = if ($_.Name -and $_.Name -ne "") { $vendedorNombre[$_.Name] } else { "Sin asignar" }
                saldo = [math]::Round((($_.Group | Measure-Object saldo -Sum).Sum),2)
            }
        } | Sort-Object saldo -Descending
        [pscustomobject]@{
            codigo = $codCliente
            nombre = if ($clienteNombre.ContainsKey($codCliente)) { $clienteNombre[$codCliente] } else { "Cliente $codCliente" }
            saldoTotal = $saldoCliente
            peorDiasVencido = $peorDias
            repartidoEntreVendedores = ($vendedoresCliente.Count -gt 1)
            vendedores = @($vendedoresCliente)
            comprobantes = @($comps | Sort-Object diasVencido -Descending)
        }
    }
    $saldoResp = [math]::Round((($clientesOut | Measure-Object saldoTotal -Sum).Sum),2)
    [pscustomobject]@{
        clave = $respKey
        nombre = NombreResponsable $respKey
        saldoTotal = $saldoResp
        clientes = @($clientesOut | Sort-Object saldoTotal -Descending)
    }
}
$responsablesOut = @($responsablesOut | Sort-Object saldoTotal -Descending)

$todosComprobantes = $responsablesOut | ForEach-Object { $_.clientes } | ForEach-Object { $_.comprobantes }
$totales = [pscustomobject]@{
    saldoTotal = [math]::Round((($todosComprobantes | Measure-Object saldo -Sum).Sum),2)
    vigente = [math]::Round(((($todosComprobantes | Where-Object bucket -eq "vigente") | Measure-Object saldo -Sum).Sum),2)
    d1_30 = [math]::Round(((($todosComprobantes | Where-Object bucket -eq "d1_30") | Measure-Object saldo -Sum).Sum),2)
    d31_60 = [math]::Round(((($todosComprobantes | Where-Object bucket -eq "d31_60") | Measure-Object saldo -Sum).Sum),2)
    d61_90 = [math]::Round(((($todosComprobantes | Where-Object bucket -eq "d61_90") | Measure-Object saldo -Sum).Sum),2)
    d90mas = [math]::Round(((($todosComprobantes | Where-Object bucket -eq "d90mas") | Measure-Object saldo -Sum).Sum),2)
}

$out = [pscustomobject]@{
    generatedAt = (Get-Date).ToString("o")
    totales = $totales
    excluido = [pscustomobject]@{
        total = [math]::Round($excluidoTotal,2)
        detalle = @($excluidoDetalle.Keys | ForEach-Object { [pscustomobject]@{ codigo=$_; nombre=$vendedorNombre[$_]; monto=[math]::Round($excluidoDetalle[$_],2) } })
    }
    responsables = $responsablesOut
}
$jsonText = $out | ConvertTo-Json -Depth 12 -Compress
[System.IO.File]::WriteAllText($OutPath, $jsonText, (New-Object System.Text.UTF8Encoding $false))
Write-Log "Guardado local: $OutPath"
Write-Log "Saldo total (excl. no-reales): $($totales.saldoTotal) | Excluido: $($excluidoTotal)"
$responsablesOut | ForEach-Object { Write-Log "  $($_.nombre): `$$($_.saldoTotal) ($($_.clientes.Count) clientes)" }

if ($NetlifySiteId -and $NetlifySiteUrl) {
    Write-Log "Desplegando a Netlify..."
    $deployDir = Join-Path $env:TEMP "cuentascorrientes-netlify-deploy"
    if (Test-Path $deployDir) { Remove-Item $deployDir -Recurse -Force }
    New-Item -ItemType Directory -Path $deployDir | Out-Null
    Copy-Item $TemplatePath (Join-Path $deployDir "index.html")
    Copy-Item $OutPath (Join-Path $deployDir "data.json")
    $zipPath = Join-Path $deployDir "site.zip"
    Compress-Archive -Path (Join-Path $deployDir "index.html"), (Join-Path $deployDir "data.json") -DestinationPath $zipPath -Force
    $netlifyHeaders = @{ Authorization = "Bearer $NetlifyToken" }
    $deployResp = Invoke-RestMethod -Uri "https://api.netlify.com/api/v1/sites/$NetlifySiteId/deploys" -Headers $netlifyHeaders -Method Post -InFile $zipPath -ContentType "application/zip"
    Write-Log "Deploy Netlify: id=$($deployResp.id) state=$($deployResp.state)"
    Write-Log "Publicado en: $NetlifySiteUrl"
}
