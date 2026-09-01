<#
Registra (o actualiza) la tarea programada de Windows que corre
stock-dashboard-refresh.ps1 cada 1 hora, indefinidamente.
Correr una sola vez para instalar/actualizar la tarea.
#>
$ErrorActionPreference = "Stop"
$TaskName = "GescomStockDashboard-Refresh"
$ScriptPath = Join-Path $PSScriptRoot "stock-dashboard-refresh.ps1"

$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""

$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Days 3650)

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 20)

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Tarea existente removida, se vuelve a crear con la configuracion actual."
}

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "Actualiza cada hora el dashboard de stock (stock-dashboard.html) con datos frescos de Gescom." | Out-Null

Write-Host "Tarea '$TaskName' registrada. Corre cada 1 hora mientras la sesion de $env:USERNAME este iniciada."
Get-ScheduledTask -TaskName $TaskName | Select-Object TaskName, State
Get-ScheduledTaskInfo -TaskName $TaskName | Select-Object NextRunTime, LastRunTime, LastTaskResult
