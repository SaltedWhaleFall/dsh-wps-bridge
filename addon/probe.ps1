$out = Join-Path $env:TEMP 'dsh-addon-probe.txt'
$lines = @()
$lines += "time=" + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
$lines += "user=" + $env:USERNAME + " computer=" + $env:COMPUTERNAME
$lines += "psScriptRoot=" + $PSScriptRoot
$lines += "pid=" + $PID
$lines += "conclusion=ShellExecute works, addon JS is running"
$lines += "addinDir=" + (Split-Path $PSScriptRoot -Parent)
Set-Content -LiteralPath $out -Value $lines -Encoding UTF8
