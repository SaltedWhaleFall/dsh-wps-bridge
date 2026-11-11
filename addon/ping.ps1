$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName Microsoft.VisualBasic | Out-Null
$cfgPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'wps-bridge.json'
$port = 8765
if (Test-Path -LiteralPath $cfgPath) {
    try { $o = Get-Content -LiteralPath $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json; if ($o.port) { $port = [int]$o.port } } catch {}
}
try {
    $r = Invoke-WebRequest -Uri ("http://127.0.0.1:" + $port + "/ping") -UseBasicParsing -TimeoutSec 3
    [Microsoft.VisualBasic.Interaction]::MsgBox($r.Content, 'OKOnly', "在线 · 端口 " + $port) | Out-Null
} catch {
    [Microsoft.VisualBasic.Interaction]::MsgBox("连不上 127.0.0.1:" + $port + [Environment]::NewLine + [Environment]::NewLine + $_.Exception.Message, 'OKOnly', 'DSH 桥 · 未运行') | Out-Null
}
