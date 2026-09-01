<#
Registra (o actualiza) la tarea programada de Windows que corre
pull-cuentas-corrientes.ps1 una vez al dia, indefinidamente.
#>
$ErrorActionPreference = "Stop"
$TaskName = "CuentasCorrientes-Refresh"
$ScriptPath = Join-Path $PSScriptRoot "pull-cuentas-corrientes.ps1"

$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""

$trigger = New-ScheduledTaskTrigger -Daily -At "07:00"

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 20)

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Tarea existente removida, se vuelve a crear con la configuracion actual."
}

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "Actualiza una vez al dia (7:00) la Torre de Cobranzas (cuenta corriente de clientes por responsable de cobro) y la despliega a Netlify." | Out-Null

Write-Host "Tarea '$TaskName' registrada. Corre una vez al dia a las 07:00 mientras la sesion de $env:USERNAME este iniciada."
Get-ScheduledTask -TaskName $TaskName | Select-Object TaskName, State
Get-ScheduledTaskInfo -TaskName $TaskName | Select-Object NextRunTime, LastRunTime, LastTaskResult
