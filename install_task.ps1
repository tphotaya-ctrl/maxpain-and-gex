# Re-register "MaxPainGEX Daily Update" so a PC that isn't always on still gets its run:
#  - weekdays 09:00 / 11:00 / 14:00, and if the PC was off then, as soon as it's on (StartWhenAvailable)
#  - at logon too (2 min delay), which also catches up after a weekend or a few days off
# run_daily.bat / run_state.py keep it to one real run per day (more only while the data is stale). Interactive logon on
# purpose: the scraper drives a visible Chrome, which needs the logged-in desktop.
#   powershell -ExecutionPolicy Bypass -File install_task.ps1
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$bat = Join-Path $here 'run_daily.bat'
$name = 'MaxPainGEX Daily Update'
$user = "$env:USERDOMAIN\$env:USERNAME"

$action = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument "/c `"$bat`"" -WorkingDirectory $here
$days = 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'
# 09:00 first; 11:00 and 14:00 pick up a previous US session CME hadn't published by 09:00
$times = '09:00', '11:00', '14:00' | ForEach-Object { New-ScheduledTaskTrigger -Weekly -DaysOfWeek $days -At $_ }
$logon = New-ScheduledTaskTrigger -AtLogOn -User $user
$logon.Delay = 'PT2M'
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 2)
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $name -Action $action -Trigger @($times + $logon) `
    -Settings $settings -Principal $principal -Force | Out-Null

$t = Get-ScheduledTask -TaskName $name
"Registered '$name'"
"  triggers:           " + (($t.Triggers | ForEach-Object { $_.CimClass.CimClassName -replace 'MSFT_Task', '' }) -join ', ')
"  StartWhenAvailable: " + $t.Settings.StartWhenAvailable
"  logon type:         " + $t.Principal.LogonType
