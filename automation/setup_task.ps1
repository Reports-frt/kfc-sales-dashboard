# =====================================================================
# KFC Dashboard - Auto-Create Scheduled Task
# =====================================================================
# Run this ONCE to create the daily auto-update task.
#
# How to run:
#   1. Right-click this file -> "Run with PowerShell"
#      OR
#   2. Open PowerShell, cd to this folder, run: .\setup_task.ps1
#
# What it does:
#   - Creates a Windows Scheduled Task named "KFC Dashboard Daily Update"
#   - Runs daily at the time you choose
#   - Calls run_update.bat which calls the Python script
#   - Configured to NOT use elevated privileges (required for Outlook COM)
#
# Re-run this script anytime to update the task settings.
# =====================================================================

# Force UTF-8 output (Greek text in console)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$ErrorActionPreference = "Stop"

$TASK_NAME = "KFC Dashboard Daily Update"
$REPO_PATH = "C:\Users\IT\Documents\GitHub\kfc-sales-dashboard"
$BAT_PATH  = Join-Path $REPO_PATH "automation\run_update.bat"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  KFC Dashboard - Scheduled Task Setup" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# --- Validate that prerequisites exist ---
if (-not (Test-Path $BAT_PATH)) {
    Write-Host "ERROR: run_update.bat not found at:" -ForegroundColor Red
    Write-Host "  $BAT_PATH" -ForegroundColor Red
    Write-Host ""
    Write-Host "Make sure you have copied all automation files to the correct folder." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "[OK] Found run_update.bat" -ForegroundColor Green

if (-not (Test-Path (Join-Path $REPO_PATH "_work\.github_pat"))) {
    Write-Host "WARNING: _work\.github_pat not found." -ForegroundColor Yellow
    Write-Host "  The task will be created but won't work until you add the PAT file." -ForegroundColor Yellow
    Write-Host ""
}
else {
    Write-Host "[OK] Found .github_pat" -ForegroundColor Green
}

if (-not (Test-Path (Join-Path $REPO_PATH "_work\budget_source.xlsx"))) {
    Write-Host "WARNING: _work\budget_source.xlsx not found." -ForegroundColor Yellow
    Write-Host "  The task will be created but won't work until you add the budget file." -ForegroundColor Yellow
    Write-Host ""
}
else {
    Write-Host "[OK] Found budget_source.xlsx" -ForegroundColor Green
}

Write-Host ""

# --- Ask the user what time the email arrives ---
Write-Host "What time does the Targit email usually arrive each day?" -ForegroundColor White
Write-Host "(Format: HH:MM in 24-hour time, e.g. 09:00 or 14:30)" -ForegroundColor DarkGray
Write-Host ""

$emailTime = $null
while ($null -eq $emailTime) {
    $input = Read-Host "Email arrival time"
    if ($input -match '^([01]?[0-9]|2[0-3]):([0-5][0-9])$') {
        $emailTime = $input
    }
    else {
        Write-Host "  Invalid format. Use HH:MM (e.g. 09:00, 14:30)" -ForegroundColor Red
    }
}

# Add 30 minutes safety buffer
$emailDateTime = [DateTime]::ParseExact($emailTime, "HH:mm", $null)
$runDateTime   = $emailDateTime.AddMinutes(30)
$runTimeStr    = $runDateTime.ToString("HH:mm")

Write-Host ""
Write-Host "Email arrives at: $emailTime" -ForegroundColor White
Write-Host "Task will run at: $runTimeStr (30 min buffer)" -ForegroundColor Green
Write-Host ""

# --- Remove existing task if it exists ---
$existingTask = Get-ScheduledTask -TaskName $TASK_NAME -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Removing existing task '$TASK_NAME'..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TASK_NAME -Confirm:$false
    Write-Host "[OK] Removed old task" -ForegroundColor Green
    Write-Host ""
}

# --- Build the task components ---

# Trigger: Daily at the chosen time
$trigger = New-ScheduledTaskTrigger -Daily -At $runTimeStr

# Action: run the .bat file
$action = New-ScheduledTaskAction `
    -Execute $BAT_PATH `
    -WorkingDirectory $REPO_PATH

# Settings:
#   - DontStopIfGoingOnBatteries / AllowStartIfOnBatteries: works on laptop battery
#   - StartWhenAvailable: if PC was off at scheduled time, run as soon as it wakes up
#   - ExecutionTimeLimit: max 1 hour
#   - RestartCount/RestartInterval: if the task fails, retry
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 15)

# Principal: run as current user, NOT elevated (Outlook COM requires same privilege level)
$principal = New-ScheduledTaskPrincipal `
    -UserId $env:USERNAME `
    -LogonType Interactive `
    -RunLevel Limited

# --- Register the task ---
Write-Host "Creating scheduled task..." -ForegroundColor Cyan
Register-ScheduledTask `
    -TaskName $TASK_NAME `
    -Description "Daily auto-update of KFC sales dashboard from Targit email" `
    -Trigger $trigger `
    -Action $action `
    -Settings $settings `
    -Principal $principal | Out-Null

Write-Host "[OK] Task created successfully" -ForegroundColor Green
Write-Host ""

# --- Show summary ---
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Task Configuration" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Name:       $TASK_NAME"
Write-Host "  Schedule:   Daily at $runTimeStr"
Write-Host "  Action:     $BAT_PATH"
Write-Host "  Privileges: Standard user (NOT elevated)"
Write-Host "  On battery: Will run"
Write-Host "  Missed run: Will run when PC wakes up"
Write-Host "  On failure: Retry up to 2 times, every 15 min"
Write-Host ""

# --- Offer to test it now ---
Write-Host "Do you want to TEST the task now (run it immediately)?" -ForegroundColor White
Write-Host "This is recommended to verify everything works." -ForegroundColor DarkGray
$test = Read-Host "Test now? (y/n)"

if ($test -eq "y" -or $test -eq "Y") {
    Write-Host ""
    Write-Host "Starting task..." -ForegroundColor Cyan
    Start-ScheduledTask -TaskName $TASK_NAME

    Write-Host "Task started. Monitoring..." -ForegroundColor Cyan
    Write-Host "(This may take up to 15 minutes for full deploy verification)" -ForegroundColor DarkGray
    Write-Host ""

    # Wait for task to finish (up to 30 minutes)
    $startTime = Get-Date
    $timeout   = New-TimeSpan -Minutes 30
    $task      = Get-ScheduledTask -TaskName $TASK_NAME

    do {
        Start-Sleep -Seconds 5
        $taskInfo = Get-ScheduledTaskInfo -TaskName $TASK_NAME
        $state    = (Get-ScheduledTask -TaskName $TASK_NAME).State
        $elapsed  = (Get-Date) - $startTime
        $el = $elapsed.ToString('mm\:ss')
        Write-Host "  [$el] State: $state" -ForegroundColor DarkGray
    } while ($state -eq "Running" -and $elapsed -lt $timeout)

    Write-Host ""
    $taskInfo = Get-ScheduledTaskInfo -TaskName $TASK_NAME
    $exitCode = $taskInfo.LastTaskResult

    if ($exitCode -eq 0) {
        Write-Host "[SUCCESS] Task completed successfully" -ForegroundColor Green
    }
    elseif ($exitCode -eq 267009) {
        Write-Host "[STILL RUNNING] Task did not finish within timeout. Check log file." -ForegroundColor Yellow
    }
    else {
        Write-Host "[FAILED] Task exited with code: $exitCode" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "Check the log for details:" -ForegroundColor White
    Write-Host "  $REPO_PATH\_work\update.log" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "Tip: open the log in real-time with:" -ForegroundColor White
    Write-Host "  Get-Content '$REPO_PATH\_work\update.log' -Tail 50" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Setup complete!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next runs: every day at $runTimeStr" -ForegroundColor White
Write-Host ""
Write-Host "To run manually anytime:" -ForegroundColor White
Write-Host "  Start-ScheduledTask -TaskName '$TASK_NAME'" -ForegroundColor DarkGray
Write-Host ""
Write-Host "To remove the task:" -ForegroundColor White
Write-Host "  Unregister-ScheduledTask -TaskName '$TASK_NAME' -Confirm:`$false" -ForegroundColor DarkGray
Write-Host ""
Write-Host "To re-run this setup with a different time:" -ForegroundColor White
Write-Host "  .\setup_task.ps1" -ForegroundColor DarkGray
Write-Host ""

Read-Host "Press Enter to close"
