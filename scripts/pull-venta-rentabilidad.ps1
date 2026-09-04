<#
Descarga las ventas de un mes desde Gescom y arma/actualiza el dataset para el
dashboard "Pizarra de Rentabilidad", guardando cada mes por separado (DATA.meses)
para poder navegar meses hacia adelante sin perder los anteriores.
  - excluye vendedor 1176 (no es venta real)
  - excluye items con precioCosto=1 (servicios/transporte)
  - neto de notas de credito (esCredito=true resta)
  - items sin proveedor asignado: se agrupan en "_SIN_PROVEEDOR_" (no se excluyen)
  - CMV = cantidad x precioCosto DEL ITEM (Gescom ya lo entrega en la misma
    unidad que "cantidad" -- para articulos pesables cantidad viene en gramos
    y este precioCosto ya viene por gramo, NO hay que usar el precioCosto del
    catalogo de articulos ahi porque ese esta expresado por kilo)
  - descuentos = suma de descuentoNeto de los items de venta normales (no NC)
  - agrupa por proveedor y por proveedor+familia

Uso:
  powershell -File pull-venta-rentabilidad.ps1                  # mes actual, al dia de hoy
  powershell -File pull-venta-rentabilidad.ps1 -MesDesde 2026-07-01   # un mes completo especifico
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/rentabilidad"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "venta-rentabilidad-template.html"),
    [string]$MesDesde = ""   # yyyy-MM-01; si se omite, usa el 1ro del mes actual
)
$ErrorActionPreference = "Stop"
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
    for ($intento = 1; $intento -le 4; $intento++) {
        try {
            return Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $script:token" } -Method Get
        } catch {
            if ($intento -eq 4) { throw }
            Start-Sleep -Seconds ($intento * 5)
        }
    }
}
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }

Write-Log "Autenticando..."
$script:token = Get-GescomToken

Write-Log "Descargando catalogos (articulos, proveedores, familias)..."
$articulos = Invoke-GescomApi -Path "/data/cmd/inventario/api/v2/get-articulos"
$proveedores = Invoke-GescomApi -Path "/data/cmd/compras/api/v1/get-proveedores"
$familias = Invoke-GescomApi -Path "/data/cmd/inventario/api/v1/get-familias"

$artInfo = @{}
foreach ($a in $articulos) { $artInfo[[string]$a.codigo] = @{ prov = [string]$a.codigoProveedor; fam = [string]$a.codigoFamilia } }
$provNombre = @{}
foreach ($p in $proveedores) { $provNombre[[string]$p.codigo] = $p.nombre }
$famNombre = @{}
foreach ($f in $familias) { $famNombre[[string]$f.codigo] = $f.descripcion }

# --- determinar rango del mes a procesar ---
if ($MesDesde -ne "") {
    $inicioMes = Get-Date $MesDesde
} else {
    $inicioMes = Get-Date -Day 1
}
$inicioMes = Get-Date -Year $inicioMes.Year -Month $inicioMes.Month -Day 1 -Hour 0 -Minute 0 -Second 0
$finMesCompleto = $inicioMes.AddMonths(1)
$hoy = (Get-Date).Date
$esMesActual = ($inicioMes.Year -eq $hoy.Year -and $inicioMes.Month -eq $hoy.Month)
# el filtro de fechaEntrega va hasta fin de mes SIEMPRE (incluye entregas ya
# facturadas con fecha futura dentro del mes en curso); el limite de la
# CONSULTA a la API (fecha de creacion) si se recorta a hoy, porque no puede
# haber ventas creadas a futuro.
$fechaHastaReal = $finMesCompleto

# La API solo filtra por fecha de creacion de la venta, pero el criterio real
# que queremos es fecha de ENTREGA -- una venta creada semanas antes puede
# entregarse recien este mes. Se pide un rango mas amplio hacia atras (buffer)
# y se filtra despues por fechaEntrega real de cada venta. El buffer se acorto
# de 30 a 12 dias: en la practica el circuito entrega/facturacion tarda mucho
# menos que eso, y una consulta mas angosta reduce el riesgo de que el servidor
# de Gescom devuelva paginas inconsistentes en consultas muy pesadas (ver nota
# de confiabilidad mas abajo).
$BUFFER_DIAS = 12
$fechaDesdeQuery = $inicioMes.AddDays(-$BUFFER_DIAS).ToString("yyyy-MM-dd")
$fechaHastaQuery = $fechaHastaReal.ToString("yyyy-MM-dd")
$mesKey = $inicioMes.ToString("yyyy-MM")
Write-Log "Procesando mes $mesKey (entregas entre $($inicioMes.ToString('yyyy-MM-dd')) y $($fechaHastaReal.ToString('yyyy-MM-dd')), consultando desde $fechaDesdeQuery)..."

# --- fetch+agregacion como funcion, para poder correrla mas de una vez y
# verificar que dos pasadas independientes coincidan antes de confiar en el
# resultado (la API de Gescom mostro ser inestable en consultas anchas: la
# misma consulta, corrida minutos aparte, puede devolver totales muy distintos) ---
function Get-VentasAgregadas {
    param([int]$Intento)
    $script:token = Get-GescomToken
    $aggLocal = @{}
    $pagestoskip = 0
    $totalLocal = 0
    $runningTotal = 0.0
    while ($true) {
        if ($pagestoskip -gt 0 -and $pagestoskip % 5 -eq 0) { $script:token = Get-GescomToken }
        $page = Invoke-GescomApi -Path "/data/cmd/ventas/api/v2/get" -Query @{
            fechadesde = $fechaDesdeQuery; fechahasta = $fechaHastaQuery; pagesize = 500; pagestoskip = $pagestoskip
        }
        if (-not $page -or $page.Count -eq 0) { break }
        $totalLocal += $page.Count
        foreach ($venta in $page) {
            if (-not $venta.fechaEntrega) { continue }
            if (-not $venta.comprobantePrincipal -or -not $venta.comprobantePrincipal.fechaComprobante) { continue }
            $fechaEntrega = ([datetime]$venta.fechaEntrega).Date
            $fechaComprobante = ([datetime]$venta.comprobantePrincipal.fechaComprobante).Date
            if ($fechaEntrega -ne $fechaComprobante) { continue }
            if ($fechaEntrega -lt $inicioMes -or $fechaEntrega -ge $fechaHastaReal) { continue }
            if ([string]$venta.codigoVendedor -eq "1176") { continue }
            $esCredito = $venta.esCredito -eq $true
            $signo = if ($esCredito) { -1.0 } else { 1.0 }
            foreach ($item in $venta.items) {
                if ([double]$item.precioCosto -eq 1.0) { continue }
                $info = $artInfo[[string]$item.codigoItem]
                $prov = if ($info) { $info.prov } else { $null }
                if (-not $prov) {
                    $prov = "_SIN_PROVEEDOR_"
                    $fam = "_SIN_FAMILIA_"
                } else {
                    $fam = if ($info.fam) { $info.fam } else { "_SIN_FAMILIA_" }
                }
                $neto = $signo * [double]$item.importeNeto
                $conImp = $signo * [double]$item.importeTotal
                $cmvItem = $signo * ([double]$item.cantidad * [double]$item.precioCosto)
                $desc = if (-not $esCredito) { [double]$item.descuentoNeto } else { 0.0 }
                $runningTotal += $neto

                $key = "$prov|$fam"
                if (-not $aggLocal.ContainsKey($key)) { $aggLocal[$key] = @{ prov=$prov; fam=$fam; ventaNeta=0.0; ventaConImp=0.0; cmv=0.0; descuentos=0.0 } }
                $aggLocal[$key].ventaNeta += $neto
                $aggLocal[$key].ventaConImp += $conImp
                $aggLocal[$key].cmv += $cmvItem
                $aggLocal[$key].descuentos += $desc
            }
        }
        if ($page.Count -lt 500) { break }
        $pagestoskip++
    }
    Write-Log "  [intento $Intento] ventas procesadas: $totalLocal | venta neta: $([math]::Round($runningTotal,2))"
    return @{ agg = $aggLocal; total = $totalLocal; ventaNeta = $runningTotal }
}

# --- correr hasta 4 veces y aceptar el resultado apenas dos pasadas coincidan
# en venta neta (tolerancia minima, solo para redondeos de punto flotante) ---
$MAX_INTENTOS = 4
$TOLERANCIA = 1.0
$pasadas = @()
$agg = $null
$total = $null
for ($intento = 1; $intento -le $MAX_INTENTOS; $intento++) {
    $r = Get-VentasAgregadas -Intento $intento
    $pasadas += $r
    $coincide = $pasadas | Where-Object { [math]::Abs($_.ventaNeta - $r.ventaNeta) -le $TOLERANCIA -and $_ -ne $r }
    if ($coincide) {
        Write-Log "Dos pasadas coincidieron en venta neta ($([math]::Round($r.ventaNeta,2))) -- se acepta el resultado."
        $agg = $r.agg
        $total = $r.total
        break
    }
}
if (-not $agg) {
    $ultimasVentaNeta = ($pasadas | ForEach-Object { [math]::Round($_.ventaNeta,0) }) -join ", "
    throw "La API de Gescom devolvio resultados distintos en las $MAX_INTENTOS pasadas ($ultimasVentaNeta) -- no se pudo verificar un numero confiable para $mesKey. No se sobreescribe el data.json existente."
}
Write-Log "Ventas procesadas: $total"

$porProveedor = @{}
foreach ($k in $agg.Keys) {
    $row = $agg[$k]
    if (-not $porProveedor.ContainsKey($row.prov)) {
        $porProveedor[$row.prov] = @{ ventaNeta=0.0; ventaConImp=0.0; cmv=0.0; descuentos=0.0; familias=@() }
    }
    $porProveedor[$row.prov].ventaNeta += $row.ventaNeta
    $porProveedor[$row.prov].ventaConImp += $row.ventaConImp
    $porProveedor[$row.prov].cmv += $row.cmv
    $porProveedor[$row.prov].descuentos += $row.descuentos
    if ($row.ventaNeta -ne 0 -or $row.cmv -ne 0) {
        $porProveedor[$row.prov].familias += [pscustomobject]@{
            codigo = $row.fam
            nombre = if ($row.fam -eq "_SIN_FAMILIA_") { "Sin familia asignada" } else { $famNombre[$row.fam] }
            ventaNeta = [math]::Round($row.ventaNeta,2)
            cmv = [math]::Round($row.cmv,2)
        }
    }
}

$proveedoresOut = foreach ($cod in $porProveedor.Keys) {
    $p = $porProveedor[$cod]
    $nombre = if ($cod -eq "_SIN_PROVEEDOR_") { "Sin proveedor asignado" } else { $provNombre[$cod] }
    if (-not $nombre) { $nombre = "Proveedor $cod" }
    [pscustomobject]@{
        codigo = $cod
        nombre = $nombre
        ventaNeta = [math]::Round($p.ventaNeta,2)
        ventaConImp = [math]::Round($p.ventaConImp,2)
        cmv = [math]::Round($p.cmv,2)
        descuentos = [math]::Round($p.descuentos,2)
        familias = @($p.familias | Sort-Object ventaNeta -Descending)
    }
}

$totales = [pscustomobject]@{
    ventaNeta = [math]::Round((($proveedoresOut | Measure-Object ventaNeta -Sum).Sum),2)
    ventaConImp = [math]::Round((($proveedoresOut | Measure-Object ventaConImp -Sum).Sum),2)
    cmv = [math]::Round((($proveedoresOut | Measure-Object cmv -Sum).Sum),2)
    descuentos = [math]::Round((($proveedoresOut | Measure-Object descuentos -Sum).Sum),2)
}

$mesData = [pscustomobject]@{
    periodoDesde = $inicioMes.ToString("yyyy-MM-dd")
    periodoHasta = $(if ($esMesActual) { $hoy.ToString("yyyy-MM-dd") } else { $finMesCompleto.AddDays(-1).ToString("yyyy-MM-dd") })
    cerrado = -not $esMesActual
    totales = $totales
    proveedores = @($proveedoresOut | Sort-Object ventaNeta -Descending)
}

# --- cargar historico existente (del data.json ya commiteado en el repo) y fusionar el mes procesado ---
$meses = [ordered]@{}
try {
    if (Test-Path $OutPath) {
        $rawText = [System.IO.File]::ReadAllText($OutPath, [System.Text.Encoding]::UTF8)  # detecta y quita BOM si esta
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
# cualquier mes anterior que no sea el que estamos procesando ahora queda cerrado
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
Write-Log "Mes $mesKey -> Venta neta: $($totales.ventaNeta) | CMV: $($totales.cmv) | Descuentos: $($totales.descuentos) | Proveedores: $($proveedoresOut.Count)"
