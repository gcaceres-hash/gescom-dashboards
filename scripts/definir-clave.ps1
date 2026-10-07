<#
Define la contrasena de los tableros protegidos SIN que quede en ningun chat ni en el repositorio.

Lo corre Gisela en su propia terminal:   powershell -File scripts\definir-clave.ps1
  1. Pide la contrasena dos veces (no se ve al escribir).
  2. La guarda en scripts\panel-config.json (campo "claveTableros"; el archivo esta ignorado por git).
  3. Copia al portapapeles el contenido COMPLETO de panel-config.json, listo para pegarlo en el
     secret PANEL_CONFIG de GitHub (Settings > Secrets and variables > Actions > PANEL_CONFIG >
     Update), asi las actualizaciones automaticas usan la misma contrasena.

Reglas de la contrasena: minimo 12 caracteres, solo letras sin acento, numeros, espacios y
signos comunes (- _ . , ; : ! ? @ # % & * + = / ( ) [ ] { } ~ ^ |). No se aceptan comillas,
barra invertida, $ ni acentos porque romperian el secret de GitHub. Conviene una frase larga:
los datos cifrados son publicos, asi que una contrasena corta se podria adivinar.
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "panel-config.json"),
    [string]$Clave = "",   # solo para pruebas; en el uso normal dejar vacio y escribirla cuando la pide
    [switch]$NoPortapapeles
)
$ErrorActionPreference = "Stop"
if (-not (Test-Path $ConfigPath)) { throw "No encuentro $ConfigPath" }

function Leer-Oculto([string]$mensaje) {
    $s = Read-Host $mensaje -AsSecureString
    $b = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($s)
    try { return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($b) } finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) }
}

if ($Clave) { $c1 = $Clave; $c2 = $Clave }
else {
    $c1 = Leer-Oculto "Elegi la contrasena de los tableros (minimo 12 caracteres)"
    $c2 = Leer-Oculto "Escribila de nuevo"
}
if ($c1 -ne $c2) { throw "Las dos contrasenas no coinciden. No se guardo nada." }
if ($c1.Length -lt 16) { throw "Muy corta: usa una frase de al menos 16 caracteres (4 palabras sin relacion entre si). No se guardo nada." }
if (@($c1 -split '[ -]' | Where-Object { $_ }).Count -lt 3) { throw "Usa una frase de al menos 3 palabras separadas por espacios o guiones (por ejemplo: tigre ventana cafe nube). No se guardo nada." }
if ($c1 -match '(?i)puelo|elebes|gescom|dashboard|contrasena|password|12345|tigre ventana') { throw "Esa frase tiene el nombre de la empresa o una palabra muy comun: es facil de adivinar. Elegi palabras sin relacion con el negocio. No se guardo nada." }
if ($c1 -notmatch '^[A-Za-z0-9 _.,;:!?@#%&*+=/()\[\]{}~^|-]+$') { throw "Tiene caracteres no permitidos (acentos, comillas, barra invertida o `$). Proba otra. No se guardo nada." }

$utf8 = New-Object System.Text.UTF8Encoding $false
$cfg = [System.IO.File]::ReadAllText($ConfigPath, $utf8) | ConvertFrom-Json
if ($cfg.PSObject.Properties.Name -contains 'claveTableros') { $cfg.claveTableros = $c1 }
else { $cfg | Add-Member -NotePropertyName claveTableros -NotePropertyValue $c1 }
$json = ConvertTo-Json -InputObject $cfg -Depth 5
[System.IO.File]::WriteAllText($ConfigPath, $json, $utf8)
if (-not $NoPortapapeles) { Set-Clipboard -Value $json }

Write-Host ""
Write-Host "Listo: la contrasena quedo guardada en $ConfigPath (no se sube a GitHub)."
Write-Host "Ya esta copiado al portapapeles el contenido completo de ese archivo."
Write-Host "Falta un paso: en GitHub, Settings > Secrets and variables > Actions > PANEL_CONFIG > Update,"
Write-Host "pegar (Ctrl+V) y guardar. Despues avisale a Claude para activar la proteccion."
