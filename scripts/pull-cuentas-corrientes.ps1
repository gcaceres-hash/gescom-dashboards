<#
Descarga la cuenta corriente de clientes y arma la estructura:
  Responsable de cobro -> Clientes -> (Comprobantes + apertura por vendedor)

MIGRADO el 1/10/2026 de la API de Gescom a la base compartida de Lucas
(datos-gescom.panelempresas.workers.dev): agrego las tablas ctacte_clientes
(detalle, 1 fila por comprobante, YA con vendedor resuelto -- no hace falta
el batch lento de ventaId->vendedor que causaba los OutOfMemoryException/502
de la version anterior) y ctacte_saldos (agregado por pagador, usado solo
para verificar el total). El campo `saldo` ya viene con el signo correcto
(negativo en las notas de credito) -- no hay que invertirlo a mano.

OJO -- hallazgo al migrar (1/10/2026): el saldo CRUDO de ctacte_clientes da
$831,7M, muy por encima de los ~$190M que mostraba el dashboard viejo. Se
confirmo que $692,6M de eso es vendedor "16" (deposito, YA excluido por la
regla de negocio de siempre) -- en gran parte facturas internas entre
empresas del mismo grupo (Lago Puelo, Elebes, Tienda Perfecta, Primeros
Productos Pehuenia) fechadas en 2024, no deuda real de clientes externos.
Con las exclusiones de siempre aplicadas, el total real da ~$139M. Vale la
pena confirmar con Lucas si ese bloque de $692,6M bajo vendedor 16 esconde
algo real ademas de los saldos entre empresas.

Reglas de negocio (confirmadas con el usuario, sin cambios):
  - Vendedores 1176 y 43 (ONCE SETENTA Y SEIS): no son reales, se excluyen SIEMPRE.
  - Vendedor 16 (VENDEDOR DEPOSITO) y 37 (LOGISTICA): se excluyen tambien.
  - Vendedores 076 (Martin Di Giorno), 038 (Gisela Caceres), 050 (Guillermo
    Zeballos): son mayoristas, se gestionan solos (responsable = ellos mismos).
  - Codigo de vendedor de 3 digitos (ej "063"): responsable = Bruno.
  - Cualquier otro codigo (1-2 digitos, ej "7", "28"): responsable = Johana.
  - Comprobantes sin vendedor asignado: responsable = "Sin vendedor asignado".
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/cobranzas"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "cuentas-corrientes-template.html")
)
$ErrorActionPreference = "Stop"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$EXCLUDED = @("43","1176","16","37")
$MAYORISTA_SOLO = @{ "076" = "Mart$([char]0xED)n Di Giorno"; "038" = "Gisela Caceres"; "050" = "Guillermo Zeballos" }

function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
function Invoke-PanelSql([string]$Sql) {
    $uri = "$($config.baseUrl)/consulta?sql=" + [uri]::EscapeDataString($Sql)
    $headers = @{ Authorization = "Bearer $($config.clave)" }
    $r = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    if ($r.truncado) { Write-Log "AVISO: la consulta vino truncada (mas de 20000 filas) -- revisar y acotar." }
    return $r.filas
}
function Get-Responsable($codigoVendedor) {
    if (-not $codigoVendedor) { return "_SIN_VENDEDOR_" }
    if ($EXCLUDED -contains $codigoVendedor) { return $null }
    if ($MAYORISTA_SOLO.ContainsKey($codigoVendedor)) { return $codigoVendedor }
    if ($codigoVendedor.Length -eq 3) { return "_BRUNO_" }
    return "_JOHANA_"
}
function Bucket($diasVencido) {
    if ($diasVencido -le 0) { return "vigente" }
    if ($diasVencido -le 30) { return "d1_30" }
    if ($diasVencido -le 60) { return "d31_60" }
    if ($diasVencido -le 90) { return "d61_90" }
    return "d90mas"
}

Write-Log "Descargando clientes, vendedores y cuenta corriente..."
$clientesRaw = Invoke-PanelSql "SELECT codigo, nombre FROM clientes"
$clienteNombre = @{}
foreach ($c in $clientesRaw) { $clienteNombre[[string]$c.codigo] = $c.nombre }
$vendedoresRaw = Invoke-PanelSql "SELECT codigo, nombre FROM vendedores"
$vendedorNombre = @{}
foreach ($v in $vendedoresRaw) { $vendedorNombre[[string]$v.codigo] = $v.nombre }
$detalle = Invoke-PanelSql "SELECT cliente, vendedor, comprobante, numero, fecha, vence, saldo, empresa FROM ctacte_clientes"
Write-Log "Comprobantes en cuenta corriente: $($detalle.Count)"

$hoy = (Get-Date).Date
$responsables = @{}
$excluidoTotal = 0.0
$excluidoDetalle = @{}

foreach ($c in $detalle) {
    $codCliente = [string]$c.cliente
    $vendCod = if ($c.vendedor) { [string]$c.vendedor } else { $null }
    $resp = Get-Responsable $vendCod
    $esCredito = [double]$c.saldo -lt 0
    $saldo = [double]$c.saldo
    if (-not $resp) {
        $excluidoTotal += $saldo
        if (-not $excluidoDetalle.ContainsKey($vendCod)) { $excluidoDetalle[$vendCod] = 0.0 }
        $excluidoDetalle[$vendCod] += $saldo
        continue
    }
    $fv = if ($c.vence) { [datetime]$c.vence } else { $hoy }
    $diasVencido = ($hoy - $fv.Date).Days
    $bucket = Bucket $diasVencido

    if (-not $responsables.ContainsKey($resp)) { $responsables[$resp] = @{} }
    if (-not $responsables[$resp].ContainsKey($codCliente)) { $responsables[$resp][$codCliente] = @() }
    $responsables[$resp][$codCliente] += [pscustomobject]@{
        comprobante = "$($c.comprobante) $($c.numero)"
        saldo = [math]::Round($saldo,2)
        esCredito = $esCredito
        fechaEmision = if ($c.fecha) { [string]$c.fecha } else { $null }
        fechaVencimiento = if ($c.vence) { [string]$c.vence } else { $null }
        diasVencido = $diasVencido
        bucket = $bucket
        codigoVendedor = $vendCod
        nombreVendedor = if ($vendCod) { $vendedorNombre[$vendCod] } else { $null }
        codigoEmpresa = [string]$c.empresa
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
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
$jsonText = $out | ConvertTo-Json -Depth 12 -Compress
[System.IO.File]::WriteAllText($OutPath, $jsonText, (New-Object System.Text.UTF8Encoding $false))
Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force
Write-Log "Guardado: $OutPath"
Write-Log "Saldo total (excl. no-reales): $($totales.saldoTotal) | Excluido: $($excluidoTotal)"
$responsablesOut | ForEach-Object { Write-Log "  $($_.nombre): `$$($_.saldoTotal) ($($_.clientes.Count) clientes)" }
