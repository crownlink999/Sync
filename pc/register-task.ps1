# Registers the Windows scheduled task 'SVODKI upload' (current user, no admin rights needed)
$d = Split-Path -Parent $MyInvocation.MyCommand.Path
$runner = Join-Path $d 'svodki-run.ps1'
$arg = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $runner + '"'
$a = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
# Run twice a day; an unchanged file is not sent again.
$times = '08:00','15:00'
$t = $times | ForEach-Object { New-ScheduledTaskTrigger -Daily -At $_ }
$s = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName 'SVODKI upload' -Action $a -Trigger $t -Settings $s -Force | Out-Null
'OK registered at ' + ($times -join ', ')
