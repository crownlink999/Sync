# Регистрирует задачу Windows «SVODKI upload» (запуск от текущего пользователя, без прав администратора)
$d = Split-Path -Parent $MyInvocation.MyCommand.Path
$runner = Join-Path $d 'svodki-run.ps1'
$arg = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $runner + '"'
$a = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
# Несколько попыток в день: файл на сервере может появиться позже; неизменённый файл повторно не шлётся.
$times = '08:00','15:00'
$t = $times | ForEach-Object { New-ScheduledTaskTrigger -Daily -At $_ }
$s = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName 'SVODKI upload' -Action $a -Trigger $t -Settings $s -Force | Out-Null
'OK registered at ' + ($times -join ', ')
