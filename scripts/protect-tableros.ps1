<#
Cifra / descifra los data.json de los tableros protegidos con contrasena.

El repositorio y el sitio son publicos, asi que una contrasena en la pagina NO alcanza:
los datos tienen que estar cifrados. Con este script el sitio solo publica
docs/<tablero>/data.enc.json (AES-256-CBC + HMAC-SHA256, clave derivada con PBKDF2-SHA256,
600.000 iteraciones). docs/gate.js lo descifra en el navegador al ingresar la contrasena.

Modos:
  Cifrar     por cada tablero protegido que tenga data.json recien generado: lo cifra a
             data.enc.json, comprueba que se pueda descifrar y recien ahi borra data.json.
  Descifrar  si solo existe data.enc.json, lo descifra a data.json (para que los scripts
             puedan leer el historico). Despues de correr los scripts, volver a Cifrar.
  Estado     muestra que hay en cada tablero. No toca nada.

La contrasena se toma, en este orden, de: -Clave, la variable de entorno DASHBOARD_PASSWORD,
o el campo "claveTableros" de scripts/panel-config.json (archivo ignorado por git; en GitHub
Actions viene del secret PANEL_CONFIG). Si no hay contrasena configurada el script no hace
nada (los tableros siguen como estaban) y termina sin error.
#>
param(
    [ValidateSet("Cifrar","Descifrar","Estado")][string]$Modo = "Estado",
    [string]$DocsDir = (Join-Path $PSScriptRoot "../docs"),
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$Clave = "",
    [string[]]$Tableros = @("rentabilidad","cobranzas","digip-resumen","digip-pedidos","digip-stock","resumen-comercial","compras","proyeccion","digiorno","compras-stock","rastreador")
)
$ErrorActionPreference = "Stop"
function Write-Log($m) { Write-Host "$(Get-Date -Format 'HH:mm:ss')  $m" }
$ITERACIONES = 600000

function Resolver-Clave {
    if ($Clave) { return $Clave }
    if ($env:DASHBOARD_PASSWORD) { return $env:DASHBOARD_PASSWORD }
    if ($ConfigPath -and (Test-Path $ConfigPath)) {
        try {
            $cfg = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($cfg.claveTableros) { return [string]$cfg.claveTableros }
        } catch { }
    }
    return ""
}

function Concatenar([byte[]]$a, [byte[]]$b) {
    $r = New-Object byte[] ($a.Length + $b.Length)
    [Buffer]::BlockCopy($a, 0, $r, 0, $a.Length)
    [Buffer]::BlockCopy($b, 0, $r, $a.Length, $b.Length)
    return ,$r
}
function Trozo([byte[]]$src, [int]$desde, [int]$largo) {
    $r = New-Object byte[] $largo
    [Buffer]::BlockCopy($src, $desde, $r, 0, $largo)
    return ,$r
}
# Derivar la clave (PBKDF2 con 600.000 vueltas) tarda varios segundos en PowerShell: se hace una
# sola vez por sal. Todos los archivos que se cifran en la misma corrida comparten la sal (cada
# archivo igual lleva su propio IV al azar) y la contrasena no cambia dentro de una corrida.
$script:CacheLlaves = @{}
$script:SalCorrida = $null
function Derivar([string]$clave, [byte[]]$sal, [int]$iter) {
    $ck = [Convert]::ToBase64String($sal) + "|" + $iter
    if ($script:CacheLlaves.ContainsKey($ck)) { return ,$script:CacheLlaves[$ck] }
    $kdf = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($clave, $sal, $iter, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
    try { $k = $kdf.GetBytes(64) } finally { $kdf.Dispose() }
    $script:CacheLlaves[$ck] = $k
    return ,$k
}

function Cifrar-Bytes([byte[]]$plano, [string]$clave) {
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    if ($null -eq $script:SalCorrida) { $script:SalCorrida = New-Object byte[] 16; $rng.GetBytes($script:SalCorrida) }
    $sal = $script:SalCorrida
    $iv = New-Object byte[] 16; $rng.GetBytes($iv)
    $rng.Dispose()
    $llave = Derivar $clave $sal $ITERACIONES
    $kEnc = Trozo $llave 0 32; $kMac = Trozo $llave 32 32
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
    $aes.Key = $kEnc; $aes.IV = $iv
    $ct = $aes.CreateEncryptor().TransformFinalBlock($plano, 0, $plano.Length)
    $aes.Dispose()
    $hm = New-Object System.Security.Cryptography.HMACSHA256 (,$kMac)
    $mac = $hm.ComputeHash((Concatenar $iv $ct))
    $hm.Dispose()
    return [ordered]@{
        v = 1; kdf = "PBKDF2-SHA256"; iter = $ITERACIONES
        salt = [Convert]::ToBase64String($sal); iv = [Convert]::ToBase64String($iv)
        ct = [Convert]::ToBase64String($ct); mac = [Convert]::ToBase64String($mac)
    }
}

function Descifrar-Sobre($sobre, [string]$clave) {
    $sal = [Convert]::FromBase64String($sobre.salt); $iv = [Convert]::FromBase64String($sobre.iv)
    $ct = [Convert]::FromBase64String($sobre.ct); $mac = [Convert]::FromBase64String($sobre.mac)
    $llave = Derivar $clave $sal ([int]$sobre.iter)
    $kEnc = Trozo $llave 0 32; $kMac = Trozo $llave 32 32
    $hm = New-Object System.Security.Cryptography.HMACSHA256 (,$kMac)
    $esperado = $hm.ComputeHash((Concatenar $iv $ct)); $hm.Dispose()
    $dif = 0
    for ($i = 0; $i -lt $esperado.Length; $i++) { $dif = $dif -bor ($esperado[$i] -bxor $mac[$i]) }
    if ($dif -ne 0 -or $esperado.Length -ne $mac.Length) { throw "Contrasena incorrecta o archivo alterado." }
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
    $aes.Key = $kEnc; $aes.IV = $iv
    $plano = $aes.CreateDecryptor().TransformFinalBlock($ct, 0, $ct.Length)
    $aes.Dispose()
    return ,$plano
}

$utf8 = New-Object System.Text.UTF8Encoding $false
$clave = Resolver-Clave
if ($Modo -eq "Estado") {
    foreach ($t in $Tableros) {
        $d = Join-Path $DocsDir $t
        $plano = Test-Path (Join-Path $d "data.json"); $enc = Test-Path (Join-Path $d "data.enc.json")
        Write-Log ("{0,-18} data.json: {1,-3} | data.enc.json: {2,-3}" -f $t, $(if ($plano) { "si" } else { "no" }), $(if ($enc) { "si" } else { "no" }))
    }
    Write-Log ("Contrasena configurada: " + $(if ($clave) { "si" } else { "no" }))
    return
}
if (-not $clave) {
    Write-Log "Sin contrasena configurada (claveTableros en panel-config.json o DASHBOARD_PASSWORD): modo $Modo sin efecto, los tableros quedan como estan."
    return
}
if ($clave.Length -lt 12) { throw "La contrasena es muy corta (minimo 12 caracteres): los datos cifrados son publicos y se podria adivinar." }

foreach ($t in $Tableros) {
    $d = Join-Path $DocsDir $t
    $pPlano = Join-Path $d "data.json"; $pEnc = Join-Path $d "data.enc.json"
    if ($Modo -eq "Cifrar") {
        if (-not (Test-Path $pPlano)) { Write-Log "$t : sin data.json nuevo, se deja como esta."; continue }
        $bytes = [System.IO.File]::ReadAllBytes($pPlano)
        $sobre = Cifrar-Bytes $bytes $clave
        $json = ConvertTo-Json -InputObject $sobre -Compress
        # comprobar que se puede descifrar y que da exactamente lo mismo antes de borrar el original
        $vuelta = Descifrar-Sobre ($json | ConvertFrom-Json) $clave
        if ($vuelta.Length -ne $bytes.Length -or [Convert]::ToBase64String($vuelta) -ne [Convert]::ToBase64String($bytes)) { throw "$t : la verificacion del cifrado fallo, no se borra data.json." }
        [System.IO.File]::WriteAllText($pEnc, $json, $utf8)
        Remove-Item $pPlano -Force
        Write-Log "$t : cifrado OK ($($bytes.Length) bytes -> data.enc.json), data.json borrado."
    } else {
        if (Test-Path $pPlano) { Write-Log "$t : ya existe data.json, no se pisa."; continue }
        if (-not (Test-Path $pEnc)) { Write-Log "$t : no hay data.enc.json para descifrar."; continue }
        $sobre = [System.IO.File]::ReadAllText($pEnc, $utf8) | ConvertFrom-Json
        $plano = Descifrar-Sobre $sobre $clave
        [System.IO.File]::WriteAllBytes($pPlano, $plano)
        Write-Log "$t : descifrado a data.json ($($plano.Length) bytes)."
    }
}
