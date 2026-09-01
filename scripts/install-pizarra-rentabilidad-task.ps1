<#
Registra (o actualiza) la tarea programada de Windows que corre
pull-venta-rentabilidad.ps1 cada 4 horas, indefinidamente.
#>
$ErrorActionPreference = "Stop"
$TaskName = "PizarraRentabilidad-Refresh"
$ScriptPath = Join-Path $PSScriptRoot "pull-venta-rentabilidad.ps1"

$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""

$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 4) -RepetitionDuration (New-TimeSpan -Days 3650)

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 20)

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Tarea existente removida, se vuelve a crear con la configuracion actual."
}

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "Actualiza cada 4 horas la Pizarra de Rentabilidad (venta y margen por proveedor) y la despliega a Netlify." | Out-Null

Write-Host "Tarea '$TaskName' registrada. Corre cada 4 horas mientras la sesion de $env:USERNAME este iniciada."
Get-ScheduledTask -TaskName $TaskName | Select-Object TaskName, State
Get-ScheduledTaskInfo -TaskName $TaskName | Select-Object NextRunTime, LastRunTime, LastTaskResult
