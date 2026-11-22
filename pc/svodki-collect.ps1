# SVODKI: берёт самый свежий (по дате изменения) подходящий файл из каждой папки sources.txt и копирует в files\srcN рядом со скриптом.
# По умолчанию ТОЛЬКО сама папка (без подпапок). Ничего никуда не отправляет.
# sources.txt: ПУТЬ | МАСКА | СЛОВА-ИСКЛЮЧЕНИЯ | дата | подпапки
#   МАСКА: * = любые символы; если начинается с re: - регулярное выражение по имени файла. Регистр не важен.
#   «дата» в 4-м поле = только файлы с датой в имени; «подпапки» в 5-м поле = искать и в подпапках.
#   {ГОД} в пути = текущий год.
$ErrorActionPreference = 'Stop'
$Dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Log = Join-Path $Dir 'log.txt'
$Out = Join-Path $Dir 'files'
$Cfg = Join-Path $Dir 'sources.txt'
function Log($m) { $l = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m; Add-Content -Path $Log -Value $l -Encoding UTF8; Write-Host $l }

function Find-Best($m) {
    $p = $m.Folder -replace '(?i)\{(ГОД|YEAR)\}', ('' + (Get-Date).Year)
    if ($p -match '[*?]') {
        $res = @(Resolve-Path -Path $p -ErrorAction SilentlyContinue | Sort-Object Path -Descending)
        if (-not $res.Count) { Log ('folder not found: ' + $p); return $null }
        $p = $res[0].ProviderPath
    }
    if (-not (Test-Path -LiteralPath $p)) { Log ('folder not reachable: ' + $p); return $null }
    $isRe = $m.Mask -match '^(?i)re:'
    $re = if ($isRe) { $m.Mask.Substring(3) } else { $null }
    $flt = if ($isRe) { '*' } else { $m.Mask }
    $args2 = @{ LiteralPath = $p; File = $true; Filter = $flt; ErrorAction = 'SilentlyContinue' }
    if ($m.Sub) { $args2.Recurse = $true }
    $excl = $m.Excl; $needDate = $m.NeedDate
    Get-ChildItem @args2 |
        Where-Object { $_.Extension -match '^\.(xlsx|xlsm|xls)$' -and $_.Name -notlike '~$*' } |
        Where-Object { (-not $isRe) -or ($_.Name -match ('(?i)' + $re)) } |
        Where-Object { $fn = $_.Name; -not ($excl | Where-Object { $fn.IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -ge 0 }) } |
        Where-Object { (-not $needDate) -or ($_.BaseName -match '\d{1,2}[._ -]\d{1,2}[._ -]\d{2,4}') } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

try {
    if ((Test-Path $Log) -and (Get-Item $Log).Length -gt 500KB) { Remove-Item $Log -Force }
    if (-not (Test-Path -LiteralPath $Cfg)) { throw 'sources.txt not found next to the script' }
    $lines = Get-Content -LiteralPath $Cfg -Encoding UTF8 | Where-Object { $_.Trim() -and -not $_.Trim().StartsWith('#') }
    if (-not $lines) { throw 'sources.txt is empty' }
    if (-not (Test-Path -LiteralPath $Out)) { New-Item -ItemType Directory -Path $Out | Out-Null }

    $groups = New-Object System.Collections.ArrayList
    foreach ($ln in $lines) {
        $t = $ln.Trim(); $isExtra = $t.StartsWith('+')
        if ($isExtra) { $t = $t.Substring(1).Trim() }
        $parts = $t.Split('|')
        $m = @{ Folder = $parts[0].Trim(); Mask = '*'; Excl = @(); NeedDate = $false; Sub = $false }
        if ($parts.Count -ge 2 -and $parts[1].Trim()) { $m.Mask = $parts[1].Trim() }
        if ($parts.Count -ge 3 -and $parts[2].Trim()) { $m.Excl = @($parts[2].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
        if ($parts.Count -ge 4 -and $parts[3].Trim() -match '(?i)^(дата|date)$') { $m.NeedDate = $true }
        if ($parts.Count -ge 5 -and $parts[4].Trim() -match '(?i)^(подпапки|sub)') { $m.Sub = $true }
        if ($isExtra -and $groups.Count -gt 0) { [void]$groups[$groups.Count - 1].Add($m) }
        else { $g = New-Object System.Collections.ArrayList; [void]$g.Add($m); [void]$groups.Add($g) }
    }

    $n = 0; $idx = 0
    foreach ($g in $groups) {
        $idx++
        $sub = Join-Path $Out ('src' + $idx)
        $where = (($g | ForEach-Object { $_.Folder + ' [' + $_.Mask + ']' }) -join ' | ')
        try {
            $best = $null
            foreach ($m in $g) {
                $b = Find-Best $m
                if ($b -and ((-not $best) -or ($b.LastWriteTime -gt $best.LastWriteTime))) { $best = $b }
            }
            if (-not $best) { Log ('no suitable Excel file in: ' + $where); continue }
            if (-not (Test-Path -LiteralPath $sub)) { New-Item -ItemType Directory -Path $sub | Out-Null }
            Get-ChildItem -LiteralPath $sub -Filter '*.part' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            $name = $best.Name
            $dst = Join-Path $sub $name
            $part = $dst + '.' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.part'
            try {
                $s = [IO.File]::Open($best.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
                try { $f = [IO.File]::Create($part); try { $s.CopyTo($f) } finally { $f.Dispose() } } finally { $s.Dispose() }
                Move-Item -LiteralPath $part -Destination $dst -Force
                (Get-Item -LiteralPath $dst).LastWriteTime = $best.LastWriteTime
            }
            catch {
                Log ('cannot copy ' + $best.FullName + ': ' + $_.Exception.Message)
                if (Test-Path -LiteralPath $part) { Remove-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue }
                continue
            }
            Get-ChildItem -LiteralPath $sub -File | Where-Object { $_.Name -ne $name } | Remove-Item -Force
            Log ('saved: ' + $name + ' (' + $best.LastWriteTime.ToString('yyyy-MM-dd HH:mm') + ') -> src' + $idx)
            $n++
        }
        catch { Log ('ERROR [' + $where + ']: ' + $_.Exception.Message) }
    }
    Log ('done: ' + $n + ' file(s)')
}
catch { Log ('ERROR: ' + $_.Exception.Message); exit 1 }
