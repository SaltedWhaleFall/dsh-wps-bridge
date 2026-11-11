$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName Microsoft.VisualBasic | Out-Null
$cfgPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'wps-bridge.json'
$port = 8765
if (Test-Path -LiteralPath $cfgPath) {
    try { $o = Get-Content -LiteralPath $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json; if ($o.port) { $port = [int]$o.port } } catch {}
}
$ips = @([System.Net.Dns]::GetHostAddresses([System.Net.Dns]::GetHostName()) |
        Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
        ForEach-Object { $_.IPAddressToString } |
        Where-Object { $_ -notlike '127.*' -and $_ -notlike '169.254.*' })
$lines = @()
foreach ($ip in $ips) { $lines += ("bash wps.sh set " + $ip + ":" + $port) }
$msg = ($lines -join [Environment]::NewLine)
if (-not $msg) { $msg = "没拿到局域网 IP，检查网络连接。" }
try { Set-Clipboard -Value $msg } catch {}
[Microsoft.VisualBasic.Interaction]::MsgBox(
    $msg + [Environment]::NewLine + [Environment]::NewLine + "（已复制到剪贴板）",
    'OKOnly', 'DSH 桥 · 手机侧命令') | Out-Null
