<#
Tablero ELEBES · El Chirlo Mayorista -- ventas del ERP Dux Software.

Dux no tiene API habilitada (21/10/2026): los datos salen del informe "Detalle de ventas"
(Estadisticas -> Detalle de ventas -> Exportar), que Dux deja bajar de a 60 dias como maximo.
Los .xls exportados van a scripts/elebes/exportes/ (ignorada por git: son datos en claro).
Este script los lee todos, arma docs/elebes/data.json y copia la plantilla.

Reglas:
  - Venta = "Total Con Iva" (se compara contra el objetivo con IVA que tiene cargado Dux).
  - Margen = Ganancia / "Total Sin IVA" (la ganancia de Dux ya es venta sin IVA - costo total).
  - Ticket = comprobante distinto. Las notas de credito restan venta y no cuentan como ticket.
  - Si dos exportes cubren el mismo dia, manda el exporte mas nuevo (por fecha del archivo).
#>
param(
    [string]$ExportesDir = (Join-Path $PSScriptRoot "elebes/exportes"),
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs/elebes"),
    [string]$OutPath = (Join-Path $DocsDir "data.json"),
    [string]$TemplatePath = (Join-Path $PSScriptRoot "elebes-template.html"),
    [double]$ObjetivoMes = 45000000,
    [double]$ObjetivoDia = 2500000
)
$ErrorActionPreference = "Stop"
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
function R2($n) { [math]::Round([double]$n, 2) }
$INV = [Globalization.CultureInfo]::InvariantCulture
function Num([string]$s) {
    if (-not $s) { return 0.0 }
    $t = $s.Trim() -replace '[\$%\s]', ''
    if ($t -eq '') { return 0.0 }
    $t = $t -replace '\.', '' -replace ',', '.'
    return [double]::Parse($t, $INV)
}
function Fecha([string]$s) { [datetime]::ParseExact($s.Trim().Substring(0, 10), "dd/MM/yyyy", $INV) }

# --- 1. Leer exportes (los .xls se pasan a CSV con Excel) ---
$archivos = @(Get-ChildItem $ExportesDir -File | Where-Object { $_.Extension -in ".xls", ".xlsx", ".csv" } | Sort-Object LastWriteTime)
if (-not $archivos.Count) { throw "No hay exportes de Dux en $ExportesDir" }
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("elebes-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmp | Out-Null
$xl = $null
$porArchivo = @()
try {
    foreach ($a in $archivos) {
        $csv = $a.FullName
        if ($a.Extension -ne ".csv") {
            if (-not $xl) { $xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false; $xl.Visible = $false; $xl.AutomationSecurity = 3 }
            $wb = $xl.Workbooks.Open($a.FullName, 0, $true)
            $csv = Join-Path $tmp ($a.BaseName + ".csv")
            $m = [Type]::Missing
            $wb.SaveAs($csv, 62, $m, $m, $m, $m, 1, $m, $m, $m, $m, $true)   # xlCSVUTF8 con separador regional (;)
            $wb.Close($false)
        }
        $lineas = Get-Content $csv -Encoding UTF8
        $ini = 0
        while ($ini -lt $lineas.Count -and -not $lineas[$ini].TrimStart([char]0xFEFF).StartsWith("Empresa;")) { $ini++ }
        if ($ini -ge $lineas.Count) { throw "$($a.Name): no encontre el encabezado del Detalle de ventas" }
        $filas = @($lineas[$ini..($lineas.Count - 1)] | ConvertFrom-Csv -Delimiter ';' | Where-Object { $_.Comprobante })
        $porArchivo += [pscustomobject]@{ Nombre = $a.Name; Filas = $filas }
        Write-Log "$($a.Name): $($filas.Count) renglones"
    }
} finally {
    if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

# Cada dia sale del exporte mas nuevo que lo trae (los archivos ya vienen ordenados de viejo a nuevo).
$duenoDia = @{}
foreach ($p in $porArchivo) { foreach ($f in $p.Filas) { $duenoDia[$f.'Fecha Comp'.Substring(0, 10)] = $p.Nombre } }
$renglones = foreach ($p in $porArchivo) { $p.Filas | Where-Object { $duenoDia[$_.'Fecha Comp'.Substring(0, 10)] -eq $p.Nombre } }
Write-Log "Renglones unicos: $(@($renglones).Count)"

# --- 2. Agregar por mes ---
function Sumar($tabla, [string]$clave, [double]$venta, [double]$sinIva, [double]$gan, [double]$cant, [double]$desc = 0, [double]$cobrado = 0) {
    if (-not $tabla.ContainsKey($clave)) { $tabla[$clave] = @{ venta = 0.0; sinIva = 0.0; gan = 0.0; cant = 0.0; desc = 0.0; cobrado = 0.0; tickets = @{} } }
    $e = $tabla[$clave]; $e.venta += $venta; $e.sinIva += $sinIva; $e.gan += $gan; $e.cant += $cant; $e.desc += $desc; $e.cobrado += $cobrado
    return $e
}
$meses = @{}
foreach ($r in $renglones) {
    $fecha = Fecha $r.'Fecha Comp'
    $mk = $fecha.ToString("yyyy-MM")
    if (-not $meses.ContainsKey($mk)) {
        $meses[$mk] = @{ dias = @{}; horas = @{}; rubros = @{}; sub = @{}; prods = @{}; marcas = @{}; pagos = @{}; vend = @{}; tickets = @{}; clientes = @{}
                         venta = 0.0; sinIva = 0.0; costo = 0.0; gan = 0.0; cant = 0.0; ncVenta = 0.0; desde = $fecha; hasta = $fecha
                         desc = 0.0; bruto = 0.0; ticketsDesc = @{}; renglonesDesc = 0; cobro = @{ efectivo = 0.0; cuenta = 0.0; tarjeta = 0.0; cheque = 0.0 }; sinCobrar = @{} }
    }
    $m = $meses[$mk]
    $esNC = $r.Comprobante -match '^NC'
    $signo = 1.0
    $venta = Num $r.'Total Con Iva'; $sinIva = Num $r.'Total Sin IVA'; $costo = Num $r.'Costo Total'; $gan = Num $r.Ganancia; $cant = Num $r.Cantidad
    if ($esNC -and $venta -gt 0) { $signo = -1.0 }
    $venta *= $signo; $sinIva *= $signo; $costo *= $signo; $gan *= $signo; $cant *= $signo
    # Descuento en $ = precio de lista x cantidad x % desc. (sin IVA; da igual a lista - "Total Sin IVA").
    # Las notas de credito no suman descuento: son correcciones posteriores.
    $desc = 0.0
    if (-not $esNC) { $desc = (Num $r.'Precio Uni') * $cant * (Num $r.'% Desc.') / 100 }
    # Plata que ingreso con la venta: Dux la reparte en efectivo / tarjeta / cheque / cuenta (Mercado Pago entra por "Cuenta").
    $ef = (Num $r.'Cobro Efectivo') * $signo; $cu = (Num $r.'Cobro Cuenta') * $signo; $tj = (Num $r.'Cobro Tarjeta') * $signo; $ch = (Num $r.'Cobro Cheque') * $signo
    $cobrado = $ef + $cu + $tj + $ch
    if ($fecha -lt $m.desde) { $m.desde = $fecha }; if ($fecha -gt $m.hasta) { $m.hasta = $fecha }
    $m.venta += $venta; $m.sinIva += $sinIva; $m.costo += $costo; $m.gan += $gan; $m.cant += $cant
    $m.desc += $desc; $m.bruto += $sinIva + $desc
    $m.cobro.efectivo += $ef; $m.cobro.cuenta += $cu; $m.cobro.tarjeta += $tj; $m.cobro.cheque += $ch
    if ($esNC) { $m.ncVenta += $venta }
    $tk = $r.Comprobante
    $cli = if ($r.Cliente) { $r.Cliente } else { "SIN CLIENTE" }
    # Saldo por ticket (venta - cobrado) para la alerta de cierre: si queda algo sin cobrar al final del dia.
    if (-not $esNC) {
        if (-not $m.sinCobrar.ContainsKey($tk)) {
            $m.sinCobrar[$tk] = @{ venta = 0.0; cobrado = 0.0; cliente = $cli; fecha = $fecha; hora = ""; vendedor = [string]$r.Vendedor }
            if ($r.'Fecha Registro' -match ' (\d{1,2}:\d{2})') { $m.sinCobrar[$tk].hora = $Matches[1] }
        }
        $m.sinCobrar[$tk].venta += $venta; $m.sinCobrar[$tk].cobrado += $cobrado
    }
    $dia = Sumar $m.dias $fecha.ToString("yyyy-MM-dd") $venta $sinIva $gan $cant $desc $cobrado
    $hora = 0; if ($r.'Fecha Registro' -match ' (\d{1,2}):') { $hora = [int]$Matches[1] }
    $h = Sumar $m.horas ([string]$hora) $venta $sinIva $gan $cant $desc $cobrado
    $rub = if ($r.Rubro) { $r.Rubro } else { "SIN RUBRO" }
    $sr = if ($r.'Sub Rubro') { $r.'Sub Rubro' } else { "SIN SUBRUBRO" }
    $null = Sumar $m.rubros $rub $venta $sinIva $gan $cant $desc $cobrado
    $null = Sumar $m.sub ($rub + "|" + $sr) $venta $sinIva $gan $cant $desc $cobrado
    $pk = $r.'Código Producto' + "|" + $r.Producto
    $pr = Sumar $m.prods $pk $venta $sinIva $gan $cant $desc $cobrado
    $pr.rubro = $rub
    $marca = if ($r.Marca) { $r.Marca } else { "SIN MARCA" }
    $null = Sumar $m.marcas $marca $venta $sinIva $gan $cant $desc $cobrado
    $pago = if ($r.'Forma Pago') { $r.'Forma Pago' } elseif ([math]::Abs($cobrado) -lt 0.01) { "SIN COBRAR" } else { "SIN DATO" }
    $pg = Sumar $m.pagos $pago $venta $sinIva $gan $cant $desc $cobrado
    $vd = Sumar $m.vend $(if ($r.Vendedor) { $r.Vendedor } else { "SIN VENDEDOR" }) $venta $sinIva $gan $cant $desc $cobrado
    if (-not $esNC) {
        $m.tickets[$tk] = 1; $dia.tickets[$tk] = 1; $h.tickets[$tk] = 1; $pg.tickets[$tk] = 1; $vd.tickets[$tk] = 1
        if ($desc -gt 0) { $m.ticketsDesc[$tk] = 1; $m.renglonesDesc++ }
        if ($r.Cliente -and $r.Cliente -ne "CONSUMIDOR FINAL") { $m.clientes[$r.Cliente] = 1 }
    }
}

function Lista($tabla, [int]$top = 0, [string]$orden = "venta") {
    $l = foreach ($k in $tabla.Keys) {
        $e = $tabla[$k]
        [ordered]@{ nombre = $k; venta = R2 $e.venta; ventaSinIva = R2 $e.sinIva; ganancia = R2 $e.gan; unidades = R2 $e.cant; tickets = $e.tickets.Count
                    descuento = R2 $e.desc; cobrado = R2 $e.cobrado }
    }
    $l = @($l | Where-Object { $orden -ne "descuento" -or $_.descuento -gt 0 } | Sort-Object { -$_[$orden] })
    if ($top -gt 0 -and $l.Count -gt $top) { $l = $l[0..($top - 1)] }
    return , $l
}

$hoy = (Get-Date).Date
$salidaMeses = [ordered]@{}
foreach ($mk in ($meses.Keys | Sort-Object)) {
    $m = $meses[$mk]
    $ini = [datetime]::ParseExact($mk + "-01", "yyyy-MM-dd", $INV)
    $fin = $ini.AddMonths(1).AddDays(-1)
    $rubros = foreach ($r in (Lista $m.rubros)) {
        $subs = @{}; foreach ($k in $m.sub.Keys) { if ($k.StartsWith($r.nombre + "|")) { $subs[$k.Substring($r.nombre.Length + 1)] = $m.sub[$k] } }
        $r.subrubros = Lista $subs 12
        $r
    }
    $prods = foreach ($p in (Lista $m.prods 30)) {
        $partes = $p.nombre.Split("|", 2)
        $p.codigo = $partes[0]; $p.nombre = $partes[1]; $p.rubro = $m.prods[$partes[0] + "|" + $partes[1]].rubro
        $p
    }
    $diario = foreach ($d in ($m.dias.Keys | Sort-Object)) {
        $e = $m.dias[$d]; [ordered]@{ fecha = $d; venta = R2 $e.venta; ganancia = R2 $e.gan; ventaSinIva = R2 $e.sinIva; tickets = $e.tickets.Count; descuento = R2 $e.desc; cobrado = R2 $e.cobrado }
    }
    $prodsDesc = foreach ($p in (Lista $m.prods 15 "descuento")) {
        $partes = $p.nombre.Split("|", 2)
        $p.codigo = $partes[0]; $p.nombre = $partes[1]; $p.rubro = $m.prods[$partes[0] + "|" + $partes[1]].rubro
        $p
    }
    # Tickets con saldo pendiente (mas de $100 para no contar redondeos ni vueltos).
    $pendTickets = @(foreach ($tk in $m.sinCobrar.Keys) {
        $e = $m.sinCobrar[$tk]; $saldo = $e.venta - $e.cobrado
        if ($saldo -gt 100) { [ordered]@{ fecha = $e.fecha.ToString("yyyy-MM-dd"); hora = $e.hora; comprobante = $tk; cliente = $e.cliente; vendedor = $e.vendedor
                                          venta = R2 $e.venta; cobrado = R2 $e.cobrado; saldo = R2 $saldo } }
    }) | Sort-Object { $_.fecha + " " + $_.hora }
    $porCli = @{}
    foreach ($p in $pendTickets) {
        if (-not $porCli.ContainsKey($p.cliente)) { $porCli[$p.cliente] = @{ venta = 0.0; tickets = 0; ultima = $p.fecha } }
        $c = $porCli[$p.cliente]; $c.venta += $p.saldo; $c.tickets++; if ($p.fecha -gt $c.ultima) { $c.ultima = $p.fecha }
    }
    $sinCobrar = @(foreach ($c in $porCli.Keys) { [ordered]@{ cliente = $c; venta = R2 $porCli[$c].venta; tickets = $porCli[$c].tickets; ultima = $porCli[$c].ultima } }) | Sort-Object { -$_.venta }
    $totalSinCobrar = 0.0; foreach ($p in $pendTickets) { $totalSinCobrar += $p.saldo }
    $pendDia = @{}; foreach ($p in $pendTickets) { $pendDia[$p.fecha] = [double]$pendDia[$p.fecha] + $p.saldo }
    foreach ($d in $diario) { $d.sinCobrar = R2 ([double]$pendDia[$d.fecha]) }
    $horas = foreach ($h in ($m.horas.Keys | Sort-Object { [int]$_ })) {
        $e = $m.horas[$h]; [ordered]@{ hora = [int]$h; venta = R2 $e.venta; tickets = $e.tickets.Count }
    }
    $salidaMeses[$mk] = [ordered]@{
        periodoDesde = $m.desde.ToString("yyyy-MM-dd")
        periodoHasta = $m.hasta.ToString("yyyy-MM-dd")
        cerrado = ($fin -lt $hoy)
        diasDelMes = $fin.Day
        totales = [ordered]@{ venta = R2 $m.venta; ventaSinIva = R2 $m.sinIva; costo = R2 $m.costo; ganancia = R2 $m.gan; unidades = R2 $m.cant
                              tickets = $m.tickets.Count; diasConVenta = $m.dias.Count; clientesIdentificados = $m.clientes.Count; notasCredito = R2 $m.ncVenta
                              descuento = R2 $m.desc; ventaBrutaSinIva = R2 $m.bruto; ticketsConDescuento = $m.ticketsDesc.Count; renglonesConDescuento = $m.renglonesDesc }
        cobros = [ordered]@{ efectivo = R2 $m.cobro.efectivo; cuenta = R2 $m.cobro.cuenta; tarjeta = R2 $m.cobro.tarjeta; cheque = R2 $m.cobro.cheque
                             total = R2 ($m.cobro.efectivo + $m.cobro.cuenta + $m.cobro.tarjeta + $m.cobro.cheque); sinCobrar = R2 $totalSinCobrar }
        sinCobrar = @($sinCobrar | Select-Object -First 15)
        pendientes = @($pendTickets)
        productosDescuento = @($prodsDesc)
        rubrosDescuento = Lista $m.rubros 0 "descuento"
        diario = @($diario)
        horas = @($horas)
        rubros = @($rubros)
        productos = @($prods)
        marcas = Lista $m.marcas 15
        pagos = Lista $m.pagos
        vendedores = Lista $m.vend
    }
    Write-Log ("{0}: venta {1:N0}  tickets {2}  margen {3:P1}" -f $mk, $m.venta, $m.tickets.Count, $(if ($m.sinIva) { $m.gan / $m.sinIva } else { 0 }))
}

$out = [ordered]@{
    generadoEn = (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
    empresa = "ELEBES SA"
    sucursal = "El Chirlo Mayorista"
    fuente = "Dux Software · Detalle de ventas (" + (($archivos | ForEach-Object { $_.Name }) -join ", ") + ")"
    objetivoMes = $ObjetivoMes
    objetivoDia = $ObjetivoDia
    mesActual = (Get-Date).ToString("yyyy-MM")
    meses = $salidaMeses
}
if (-not (Test-Path $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
[System.IO.File]::WriteAllText($OutPath, ($out | ConvertTo-Json -Depth 8 -Compress), (New-Object System.Text.UTF8Encoding $false))
Copy-Item $TemplatePath (Join-Path $DocsDir "index.html") -Force
Write-Log "Guardado: $OutPath"
