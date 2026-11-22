$d = Split-Path -Parent $MyInvocation.MyCommand.Path
try { & (Join-Path $d "svodki-collect.ps1") } catch {}
try { & (Join-Path $d "svodki-send.ps1") } catch {}
