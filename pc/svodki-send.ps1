# SVODKI: отправляет файлы из files\srcN на приёмник Google Apps Script (-> GitHub, ветка inbox).
# send.txt: 1-я строка = адрес веб-приложения, 2-я = KEY. Не изменившиеся файлы повторно не шлются (sent.json).
# Файл больше 8 МБ режется на части; в одном запросе уходит одна часть, поэтому лимиты Google не мешают.
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
try { [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultCredentials } catch {}
$Dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Log = Join-Path $Dir 'log.txt'
$Src = Join-Path $Dir 'files'
$State = Join-Path $Dir 'sent.json'
$CHUNK = 8MB
function Log($m) { $l = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m; Add-Content $Log $l -Encoding UTF8; Write-Host $l }

function Post($url, $key, $item) {
    $json = @{ key = $key; files = @($item) } | ConvertTo-Json -Depth 4 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    for ($try = 1; $try -le 3; $try++) {
        try {
            $r = Invoke-RestMethod -Uri $url -Method Post -Body $bytes -ContentType 'application/json; charset=utf-8' -TimeoutSec 600
            if ($r.ok) { return }
            throw ('server: ' + $r.error)
        }
        catch { if ($try -eq 3) { throw }; Start-Sleep -Seconds (5 * $try) }
    }
}

try {
    $cfg = Get-Content -LiteralPath (Join-Path $Dir 'send.txt') -Encoding UTF8 | Where-Object { $_.Trim() }
    $url = $cfg[0].Trim(); $key = $cfg[1].Trim()
    if ($url -notmatch '^https://script\.google\.com/') { throw 'send.txt: первая строка должна быть адресом веб-приложения' }
    $sent = @{}
    if (Test-Path $State) { (Get-Content $State -Raw -Encoding UTF8 | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $sent[$_.Name] = $_.Value } }
    $names = @()
    if (Test-Path -LiteralPath $Src) {
        foreach ($d in (Get-ChildItem -LiteralPath $Src -Directory | Sort-Object Name)) {
            foreach ($f in (Get-ChildItem -LiteralPath $d.FullName -File | Where-Object { $_.Extension -match '^\.(xlsx|xlsm|xls)$' })) {
                $k = $d.Name + '/' + $f.Name
                $s = '{0}|{1}' -f $f.Length, $f.LastWriteTimeUtc.Ticks
                if ($sent[$k] -eq $s) { continue }
                $bytes = [IO.File]::ReadAllBytes($f.FullName)
                $total = [int][Math]::Ceiling($bytes.Length / [double]$CHUNK)
                if ($total -lt 1) { $total = 1 }
                for ($i = 0; $i -lt $total; $i++) {
                    $off = $i * $CHUNK; $len = [Math]::Min($CHUNK, $bytes.Length - $off)
                    $b64 = [Convert]::ToBase64String($bytes, $off, $len)
                    Post $url $key @{ slot = $d.Name; name = $f.Name; part = ($i + 1); parts = $total; b64 = $b64 }
                    Log ('send: ' + $k + ' часть ' + ($i + 1) + '/' + $total + ' OK')
                }
                $sent[$k] = $s
                $sent | ConvertTo-Json | Set-Content $State -Encoding UTF8
                $names += $f.Name
            }
        }
    }
    if (-not $names.Count) { Log 'send: nothing new'; exit 0 }
    Log ('send: OK, files=' + $names.Count + ' (' + ($names -join '; ') + ')')
}
catch { Log ('SEND ERROR: ' + $_.Exception.Message); exit 1 }
