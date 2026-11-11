$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName Microsoft.VisualBasic | Out-Null

# 配置放"加载项目录的父级"(jsaddons\)，这样重装加载项不会把它删掉
$cfgPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'wps-bridge.json'
$cfg = [pscustomobject]@{ port = 8765; allowScript = $true; scriptPath = 'C:\Users\admin\wps-skills\scripts\wps-exec.ps1' }
if (Test-Path -LiteralPath $cfgPath) {
    try {
        $o = Get-Content -LiteralPath $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($o.port) { $cfg.port = [int]$o.port }
        if ($null -ne $o.allowScript) { $cfg.allowScript = [bool]$o.allowScript }
        if ($o.scriptPath) { $cfg.scriptPath = [string]$o.scriptPath }
    } catch {}
}

$p = [Microsoft.VisualBasic.Interaction]::InputBox(
    "服务端口（1-65535）" + [Environment]::NewLine + [Environment]::NewLine + "手机侧要用同一个端口",
    "DSH 桥 · 端口", [string]$cfg.port)
if ($p -and $p -match '^\d+$' -and [int]$p -ge 1 -and [int]$p -le 65535) { $cfg.port = [int]$p }

$s = [Microsoft.VisualBasic.Interaction]::InputBox(
    "wps-exec.ps1 的完整路径",
    "DSH 桥 · 脚本路径", [string]$cfg.scriptPath)
if ($s -and (Test-Path -LiteralPath $s)) { $cfg.scriptPath = $s }

$a = [Microsoft.VisualBasic.Interaction]::MsgBox(
    "允许远程执行任意 PowerShell（/script 通道）？" + [Environment]::NewLine + [Environment]::NewLine +
    "选「是」= 我能下发脚本调试；选「否」= 只开放读写动作。",
    [Microsoft.VisualBasic.MsgBoxStyle]::YesNo, "DSH 桥 · 脚本通道")
$cfg.allowScript = ($a -eq [Microsoft.VisualBasic.MsgBoxResult]::Yes)

$cfg | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $cfgPath -Encoding UTF8

[Microsoft.VisualBasic.Interaction]::MsgBox(
    "已保存" + [Environment]::NewLine + [Environment]::NewLine +
    "端口: " + $cfg.port + [Environment]::NewLine +
    "脚本: " + $cfg.scriptPath + [Environment]::NewLine +
    "脚本通道: " + $cfg.allowScript + [Environment]::NewLine + [Environment]::NewLine +
    "文件: " + $cfgPath,
    [Microsoft.VisualBasic.MsgBoxStyle]::OKOnly, "DSH 桥") | Out-Null
