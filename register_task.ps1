# ── Held Time Report — register as startup task ──────────────────────────────
# Run this script once from PowerShell (as your normal user, no admin needed).
# After running, the dashboard will start automatically every time you log in.

$taskName   = "HeldTimeReport"
$vbsPath    = "C:\Users\Aileen Holland\OneDrive - Clearspace Offices\Aileen - Design Group\held-time-report\start_hidden.vbs"
$action     = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$vbsPath`""
$trigger    = New-ScheduledTaskTrigger -AtLogOn
$settings   = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit 0
$principal  = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

# Remove existing task if present
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

# Register
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal

Write-Host ""
Write-Host "✅  Task '$taskName' registered successfully." -ForegroundColor Green
Write-Host "    The Held Time Report dashboard will start automatically at login." -ForegroundColor Gray
Write-Host "    It will be available at: http://localhost:5000" -ForegroundColor Cyan
Write-Host ""
Write-Host "To start it now without rebooting, run:" -ForegroundColor Yellow
Write-Host "    Start-ScheduledTask -TaskName '$taskName'" -ForegroundColor White
Write-Host ""
