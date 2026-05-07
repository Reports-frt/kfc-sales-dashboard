# =====================================================================
# KFC Food Cost - Task Scheduler Setup
# =====================================================================
# Creates a daily scheduled task that runs run_daily_food.bat
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File setup_food_task.ps1
# =====================================================================

$taskName = "KFC Food Cost Daily Update"
$batPath = Join-Path $PSScriptRoot "run_daily_food.bat"
$triggerTime = "09:30"

Write-Host ""
Write-Host "============================================"
Write-Host "KFC Food Cost - Task Scheduler Setup"
Write-Host "============================================"
Write-Host ""
Write-Host "Task name: $taskName"
Write-Host "Script:    $batPath"
Write-Host "Schedule:  Daily at $triggerTime"
Write-Host ""

if (-not (Test-Path $batPath)) {
    Write-Host "ERROR: $batPath not found!" -ForegroundColor Red
    Write-Host "Make sure run_daily_food.bat is in the same folder as this script."
    Read-Host "Press Enter to exit"
    exit 1
}

$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "Removing existing task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

$trigger = New-ScheduledTaskTrigger -Daily -At $triggerTime
$action = New-ScheduledTaskAction -Execute $batPath -WorkingDirectory $PSScriptRoot

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive

try {
    Register-ScheduledTask -TaskName $taskName -Trigger $trigger -Action $action -Settings $settings -Principal $principal -Description "Daily auto-update of KFC food cost dashboard"

    Write-Host ""
    Write-Host "[OK] Task created successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "The task will run every day at $triggerTime."
    Write-Host ""
    Write-Host "To test it now:"
    Write-Host "  Start-ScheduledTask -TaskName '$taskName'"
    Write-Host ""
    Write-Host "To remove it later:"
    Write-Host "  Unregister-ScheduledTask -TaskName '$taskName' -Confirm:" -NoNewline
    Write-Host '$false'
}
catch {
    Write-Host ""
    Write-Host "[FAIL] Failed to create task: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "If you get permission errors, try running PowerShell as Administrator."
}

Write-Host ""
Read-Host "Press Enter to close"
