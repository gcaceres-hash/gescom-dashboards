<#
Arma/actualiza el dataset de la Pizarra de Rentabilidad a partir de un archivo
"Detallado de venta extendido" exportado manualmente desde Gescom (CSV con ;
como separador, codificacion Windows-1252), en vez de consultar la API en
vivo. Se uso porque la API (ventas/api/v2/get) mostro ser poco confiable en
consultas anchas: omitia ventas reales de forma inconsistente entre corridas
(confirmado con casos concretos). El export de Gescom es la fuente de verdad.

Reglas de negocio (mismas que la version por API):
  - excluye vendedores 1176, 43, 16, 37 (no son ventas reales)
  - excluye items con PrecioCosto=1 (servicios/transporte)
  - solo cuenta filas donde FechaComprobante = FechaEntrega
  - ImporteNetoItem ya viene con el signo correcto en el archivo (negativo
    para devoluciones/notas de credito, no hay que invertirlo)
  - items sin proveedor asignado: se agrupan en "_SIN_PROVEEDOR_" (no se excluyen)
  - CMV = PrecioCosto tal cual viene (en este archivo es el costo total de
    la linea, no hay que multiplicarlo por CantBase) EXCEPTO para RIOSMA:
    ahi los articulos son pesables, PrecioCosto viene por GRAMO pero
    escalado por CantBase (no es el costo total de la linea ni un puro
    costo por gramo) -- para RIOSMA con PesoKgReal<>0:
      CMV = (PrecioCosto / CantBase) x PesoKgReal x 1000
    (verificado fila por fila contra el archivo: con esta formula el
    margen de RIOSMA da ~14% de forma consistente, sea CantBase=1 o mayor;
    otros proveedores tambien completan PesoKg/PesoKgReal para logistica
    (ej. ILOLAY con leche por litro) pero ahi PrecioCosto SI es el costo
    total ya calculado, por eso el ajuste NO se generaliza a todos)
  - descuentos = valorDescuento, solo para filas de tipo "Venta" (no devoluciones)
  - agrupa por proveedor y por proveedor+familia
  - el mes que arma queda determinado por el mes de las fechas del archivo
    (si el archivo mezclara mas de un mes, se separan automaticamente)

Uso:
  powershell -File pull-venta-rentabilidad-csv.ps1 -CsvPath "C:\...\venta 4092026.csv"
#>
param(
    [Parameter(Mandatory=$true)][string]$CsvPath,
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/rentabilidad"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "venta-rentabilidad-template.html")
)
$ErrorActionPreference = "Stop"
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
function ToNum($s) { if ([string]::IsNullOrWhiteSpace($s)) { return 0.0 }; return [double]($s -replace ',', '.') }
function ToFecha($s) { [datetime]::ParseExact($s, "d/M/yyyy", [System.Globalization.CultureInfo]::InvariantCulture) }

$EXCLUDED = @("43","1176","16","37")

Write-Log "Leyendo $CsvPath ..."
$enc = [System.Text.Encoding]::GetEncoding(1252)
$lines = [System.IO.File]::ReadAllLines($CsvPath, $enc)
$header = $lines[0] -split ';'
$col = @{}
for ($i = 0; $i -lt $header.Count; $i++) { $col[$header[$i]] = $i }
foreach ($req in @("FechaComprobante","FechaEntrega","ImporteNetoItem","ImporteItem","CodVendedor","PrecioCosto","CantBase","Descuento","valorDescuento","Proveedor","Familia","TipoDeVenta","NumeroVenta","PesoKgReal")) {
    if (-not $col.ContainsKey($req)) { throw "El CSV no tiene la columna esperada '$req'. Encabezado: $($header -join ', ')" }
}
Write-Log "Filas de datos: $($lines.Count - 1)"

# --- agregacion por mes (por si el archivo mezcla mas de un mes) ---
$porMes = @{}
$totalFilas = 0
$totalCalifican = 0
for ($i = 1; $i -lt $lines.Count; $i++) {
    $f = $lines[$i] -split ';'
    $totalFilas++
    $codVend = $f[$col['CodVendedor']]
    if ($EXCLUDED -contains $codVend) { continue }
    $precioCosto = ToNum $f[$col['PrecioCosto']]
    if ($precioCosto -eq 1.0) { continue }
    $fc = $f[$col['FechaComprobante']]
    $fe = $f[$col['FechaEntrega']]
    if ($fc -ne $fe) { continue }
    $fecha = ToFecha $fc
    $mesKey = $fecha.ToString("yyyy-MM")
    $totalCalifican++

    if (-not $porMes.ContainsKey($mesKey)) { $porMes[$mesKey] = @{ agg = @{}; inicioMes = (Get-Date -Year $fecha.Year -Month $fecha.Month -Day 1); minFecha = $fecha; maxFecha = $fecha } }
    $m = $porMes[$mesKey]
    if ($fecha -lt $m.minFecha) { $m.minFecha = $fecha }
    if ($fecha -gt $m.maxFecha) { $m.maxFecha = $fecha }

    $prov = $f[$col['Proveedor']]
    $fam = $f[$col['Familia']]
    if (-not $prov) { $prov = "_SIN_PROVEEDOR_"; $fam = "_SIN_FAMILIA_" }
    elseif (-not $fam) { $fam = "_SIN_FAMILIA_" }

    $neto = ToNum $f[$col['ImporteNetoItem']]
    $conImp = ToNum $f[$col['ImporteItem']]
    $pesoKgReal = ToNum $f[$col['PesoKgReal']]
    $cantidad = ToNum $f[$col['CantBase']]
    $esRiosma = $prov -eq "RIOSMA"
    $cmvItem = if ($esRiosma -and $pesoKgReal -ne 0 -and $cantidad -ne 0) { ($precioCosto / $cantidad) * $pesoKgReal * 1000 } else { $precioCosto }
    $esVenta = $f[$col['TipoDeVenta']] -eq "Venta"
    $desc = if ($esVenta) { ToNum $f[$col['valorDescuento']] } else { 0.0 }

    $key = "$prov|$fam"
    if (-not $m.agg.ContainsKey($key)) { $m.agg[$key] = @{ prov=$prov; fam=$fam; ventaNeta=0.0; ventaConImp=0.0; cmv=0.0; descuentos=0.0 } }
    $m.agg[$key].ventaNeta += $neto
    $m.agg[$key].ventaConImp += $conImp
    $m.agg[$key].cmv += $cmvItem
    $m.agg[$key].descuentos += $desc
}
Write-Log "Filas leidas: $totalFilas | filas que califican (FechaComprobante=FechaEntrega, vendedor real, no servicio): $totalCalifican"

# --- cargar historico existente y fusionar cada mes del archivo ---
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

$hoy = (Get-Date).Date
foreach ($mesKey in $porMes.Keys) {
    $m = $porMes[$mesKey]
    $inicioMes = $m.inicioMes
    $esMesActual = ($inicioMes.Year -eq $hoy.Year -and $inicioMes.Month -eq $hoy.Month)

    $porProveedor = @{}
    foreach ($k in $m.agg.Keys) {
        $row = $m.agg[$k]
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
                nombre = if ($row.fam -eq "_SIN_FAMILIA_") { "Sin familia asignada" } else { $row.fam }
                ventaNeta = [math]::Round($row.ventaNeta,2)
                cmv = [math]::Round($row.cmv,2)
            }
        }
    }
    $proveedoresOut = foreach ($cod in $porProveedor.Keys) {
        $p = $porProveedor[$cod]
        $nombre = if ($cod -eq "_SIN_PROVEEDOR_") { "Sin proveedor asignado" } else { $cod }
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
        periodoHasta = $m.maxFecha.ToString("yyyy-MM-dd")
        cerrado = -not $esMesActual
        totales = $totales
        proveedores = @($proveedoresOut | Sort-Object ventaNeta -Descending)
    }
    foreach ($k in @($meses.Keys)) {
        if ($k -ne $mesKey -and -not $meses[$k].cerrado) { $meses[$k] | Add-Member -NotePropertyName cerrado -NotePropertyValue $true -Force }
    }
    $meses[$mesKey] = $mesData
    Write-Log "Mes $mesKey ($($inicioMes.ToString('yyyy-MM-dd')) a $($m.maxFecha.ToString('yyyy-MM-dd'))) -> Venta neta: $($totales.ventaNeta) | CMV: $($totales.cmv) | Descuentos: $($totales.descuentos) | Proveedores: $($proveedoresOut.Count)"
}

$out = [pscustomobject]@{
    generatedAt = (Get-Date).ToString("o")
    mesActual = (Get-Date -Format "yyyy-MM")
    fuente = "csv:$([System.IO.Path]::GetFileName($CsvPath))"
    meses = $meses
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
$jsonText = $out | ConvertTo-Json -Depth 10 -Compress
[System.IO.File]::WriteAllText($OutPath, $jsonText, (New-Object System.Text.UTF8Encoding $false))
Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force
Write-Log "Guardado: $OutPath"
