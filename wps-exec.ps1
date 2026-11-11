# wps-exec.ps1  —  WPS 局域网执行端
# 作用：在本机开一个 HTTP 端口，接收 JSON 指令，转调官方 wps-com.ps1 通过 COM 操作 WPS。
# 用法：右键「使用 PowerShell 运行」，或  powershell -ExecutionPolicy Bypass -File wps-exec.ps1
# 停止：在本窗口按 Ctrl+C
#
# 接口：
#   GET  /ping                      探活
#   GET  /                           浏览器里的说明页
#   POST /wps   {"action":..,"params":{}}      转调 wps-com.ps1
#   POST /script {"code":"..."}      直接执行 PowerShell（需 -AllowScript 启动）

param(
    [int]$Port = 8765,
    [string]$Token = "",
    [string]$ComPath = "",
    [switch]$AllowScript,
    [switch]$NoElevate
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

$LogFile = Join-Path $PSScriptRoot 'wps-exec.log'

function Write-Log([string]$msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg
    Write-Host $line -ForegroundColor DarkGray
    try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 } catch {}
}

# ---------- 自动提权（监听局域网端口需要管理员） ----------
$isAdmin = $false
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch {}

if (-not $isAdmin -and -not $NoElevate) {
    Write-Host ""
    Write-Host "  需要管理员权限才能在局域网里被访问，正在弹 UAC，请点「是」。" -ForegroundColor Yellow
    Write-Host "  （不弹窗就右键这个文件 -> 使用 PowerShell 运行）" -ForegroundColor DarkGray
    $argLine = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Port {1}' -f $PSCommandPath, $Port
    if ($Token)     { $argLine += ' -Token "{0}"' -f $Token }
    if ($ComPath)   { $argLine += ' -ComPath "{0}"' -f $ComPath }
    if ($AllowScript) { $argLine += ' -AllowScript' }
    try {
        Start-Process powershell -Verb RunAs -ArgumentList $argLine
    } catch {
        Write-Host "  提权失败：$($_.Exception.Message)" -ForegroundColor Red
    }
    exit
}

# ---------- 找 wps-com.ps1 ----------
function Find-ComPath {
    $cands = @()
    if ($ComPath) { $cands += $ComPath }
    $cands += (Join-Path $PSScriptRoot 'wps-com.ps1')
    $cands += (Join-Path $PSScriptRoot 'wps-office-mcp\scripts\wps-com.ps1')
    $cands += (Join-Path (Split-Path $PSScriptRoot -Parent) 'wps-office-mcp\scripts\wps-com.ps1')
    $cands += (Join-Path $PSScriptRoot '..\wps-com.ps1')
    foreach ($c in $cands) {
        if ($c -and (Test-Path -LiteralPath $c)) { return (Resolve-Path -LiteralPath $c).Path }
    }
    return $null
}
if (-not $ComPath) { $ComPath = Find-ComPath }
if (-not $ComPath -or -not (Test-Path -LiteralPath $ComPath)) {
    Write-Host ""
    Write-Host "  没找到 wps-com.ps1。" -ForegroundColor Red
    Write-Host "  把它放到本脚本同一个文件夹，或者加参数：-ComPath `"D:\...\wps-com.ps1`"" -ForegroundColor Yellow
    Read-Host "按回车退出"
    exit 1
}

# ---------- 防火墙放行（幂等） ----------
try {
    $ruleName = "WPS Exec $Port"
    $existing = netsh advfirewall firewall show rule name="$ruleName" 2>$null
    if (-not ($existing -match 'Rule Name')) {
        netsh advfirewall firewall add rule name="$ruleName" dir=in action=allow protocol=TCP localport=$Port | Out-Null
        Write-Log "已添加防火墙规则：$ruleName"
    }
} catch { Write-Log "防火墙规则添加失败（不影响本机，可能影响局域网）：$($_.Exception.Message)" }


# ---------- 带超时的子进程执行器（防挂死：一个卡住的脚本不能拖垮整个桥） ----------
function Invoke-ChildScript([string]$code, [int]$TimeoutMs = 60000) {
    $tmp = [IO.Path]::Combine([IO.Path]::GetTempPath(), ('wpsrun_' + [Guid]::NewGuid().ToString('N') + '.ps1'))
    [IO.File]::WriteAllText($tmp, $code, (New-Object System.Text.UTF8Encoding($true)))
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName  = (Get-Command powershell).Source
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $tmp + '"'
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = [Text.Encoding]::GetEncoding(936)
    $psi.StandardErrorEncoding  = [Text.Encoding]::GetEncoding(936)
    $timedOut = $false
    $out = ''
    try {
        $p = [Diagnostics.Process]::Start($psi)
        $so = $p.StandardOutput.ReadToEndAsync()
        $se = $p.StandardError.ReadToEndAsync()
        if (-not $p.WaitForExit($TimeoutMs)) {
            $timedOut = $true
            try { $p.Kill() } catch {}
        }
        $o1 = ''; $o2 = ''
        try { $o1 = $so.GetAwaiter().GetResult() } catch {}
        try { $o2 = $se.GetAwaiter().GetResult() } catch {}
        $out = ($o1 + $(if ($o2) { "`n" + $o2 } else { '' })).Trim()
    } catch {
        $out = 'ERR ' + $_.Exception.Message
    } finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
    if ($timedOut) { $out = ($out + "`n[TIMEOUT ${TimeoutMs}ms - 子进程已终止]").Trim() }
    return [pscustomobject]@{ out = $out; timedOut = $timedOut }
}


# ---------- 常驻 COM worker（COM 对象只在同一进程内可复用，这是唯一可靠的架构） ----------
$script:WorkerRoot = $null
$script:WorkerReq  = $null
$script:WorkerResp = $null
$script:WorkerPid  = 0

function Start-Worker {
    $root = Join-Path $env:TEMP 'wps-worker'
    New-Item -ItemType Directory -Force -Path (Join-Path $root 'req'), (Join-Path $root 'resp') | Out-Null
    $script:WorkerRoot = $root
    $script:WorkerReq  = Join-Path $root 'req'
    $script:WorkerResp = Join-Path $root 'resp'
    $wp = Join-Path $PSScriptRoot 'wps-worker.ps1'
    if (-not (Test-Path -LiteralPath $wp)) { Write-Log "worker 脚本缺失：$wp"; return $false }
    $wargs = @('-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $wp + '"'), '-Root', ('"' + $root + '"'))
    $werr = Join-Path $root 'worker.err'
    $wout = Join-Path $root 'worker.out'
    # -WindowStyle 与 -RedirectStandard* 属于互斥参数集，同时给会抛“参数集无法解析”，
    # worker 根本起不来。要日志就只能 -NoNewWindow。
    $p = Start-Process powershell -ArgumentList $wargs -PassThru -NoNewWindow -RedirectStandardError $werr -RedirectStandardOutput $wout
    $script:WorkerPid = $p.Id
    Write-Log "worker 启动 pid=$($p.Id)"
    return $true
}

function Invoke-Worker($payload, [int]$TimeoutMs = 120000) {
    if (-not $script:WorkerReq) { return @{ ok = $false; error = 'worker not configured' } }
    if ($script:WorkerPid -and -not (Get-Process -Id $script:WorkerPid -ErrorAction SilentlyContinue)) {
        Write-Log "worker 已死，重启"
        Start-Worker | Out-Null
        Start-Sleep -Seconds 2
    }
    $id = [Guid]::NewGuid().ToString('N')
    $reqPath  = Join-Path $script:WorkerReq  ($id + '.json')
    $respPath = Join-Path $script:WorkerResp ($id + '.json')
    $json = $payload | ConvertTo-Json -Depth 15 -Compress
    [IO.File]::WriteAllText($reqPath, $json, (New-Object Text.UTF8Encoding($false)))
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $TimeoutMs) {
        if (Test-Path -LiteralPath $respPath) {
            Start-Sleep -Milliseconds 20
            $txt = ''
            try { $txt = [IO.File]::ReadAllText($respPath, [Text.Encoding]::UTF8) } catch { Start-Sleep -Milliseconds 60; continue }
            Remove-Item -LiteralPath $respPath -Force -ErrorAction SilentlyContinue
            try { return ($txt | ConvertFrom-Json) } catch { return @{ ok = $false; error = 'bad worker response'; raw = $txt } }
        }
        if ($script:WorkerPid -and -not (Get-Process -Id $script:WorkerPid -ErrorAction SilentlyContinue)) {
            return @{ ok = $false; error = 'worker died during call' }
        }
        Start-Sleep -Milliseconds 40
    }
    Remove-Item -LiteralPath $reqPath -Force -ErrorAction SilentlyContinue
    return @{ ok = $false; error = ("worker timeout " + $TimeoutMs + "ms") }
}

# ---------- 调 COM 桥（子进程，隔离 exit 与编码） ----------
function Invoke-ComBridge([string]$action, $params, [int]$TimeoutMs = 45000) {
    $json = '{}'
    if ($null -ne $params) {
        if ($params -is [string]) { $json = $params }
        else { $json = $params | ConvertTo-Json -Depth 12 -Compress }
    }
    $escJson   = $json      -replace "'", "''"
    $escPath   = $ComPath   -replace "'", "''"
    $escAction = $action    -replace "'", "''"
    $tmp = [IO.Path]::Combine([IO.Path]::GetTempPath(), ("wpscall_" + [Guid]::NewGuid().ToString('N') + ".ps1"))
    $body = @"
`$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; `$OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
& '$escPath' -Action '$escAction' -Params '$escJson'
"@
    [IO.File]::WriteAllText($tmp, $body, (New-Object System.Text.UTF8Encoding($true)))
    $res = Invoke-ChildScript $body $TimeoutMs
    $out = $res.out
    $parsed = $null
    try {
        $start = $out.IndexOf('{')
        if ($start -ge 0) { $parsed = $out.Substring($start) | ConvertFrom-Json }
    } catch {}
    return [pscustomobject]@{ ok = $true; action = $action; parsed = $parsed; raw = $out }
}

# ---------- HTTP 工具 ----------
function Send-Bytes($ctx, [byte[]]$bytes, [string]$type, [int]$code) {
    $ctx.Response.StatusCode = $code
    $ctx.Response.ContentType = $type
    $ctx.Response.ContentLength64 = $bytes.Length
    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $ctx.Response.OutputStream.Close()
}
function Send-Json($ctx, $obj, [int]$code = 200) {
    $json = $obj | ConvertTo-Json -Depth 12 -Compress
    Send-Bytes $ctx ([Text.Encoding]::UTF8.GetBytes($json)) 'application/json; charset=utf-8' $code
}
function Read-Body($ctx) {
    # 一律按 UTF-8 读。HttpListener 在 Content-Type 没写 charset 时会退回系统 ANSI(GBK)，
    # 中文会被解成乱码并污染下游 JSON（2026-11-04 实测：写进 Excel 的"项目"变成"椤圭洰"）。
    $reader = New-Object IO.StreamReader($ctx.Request.InputStream, [Text.Encoding]::UTF8, $true)
    try { return $reader.ReadToEnd() } finally { $reader.Close() }
}
function Get-LocalIPs {
    $ips = @()
    try {
        $ips = [System.Net.Dns]::GetHostAddresses([System.Net.Dns]::GetHostName()) |
               Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
               ForEach-Object { $_.IPAddressToString }
    } catch {}
    if (-not $ips) {
        try {
            $ips = Get-NetIPAddress -AddressFamily IPv4 |
                   Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' } |
                   ForEach-Object { $_.IPAddress }
        } catch {}
    }
    return @($ips | Sort-Object -Unique)
}

# ---------- 启动 worker ----------
Start-Worker | Out-Null
Start-Sleep -Milliseconds 800

# ---------- 启动 ----------
$prefix = "http://+:$Port/"
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
try {
    $listener.Start()
} catch {
    Write-Host ""
    Write-Host "  端口 $Port 启动失败：$($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  多半是没提权，或端口被占用。换端口：-Port 8766" -ForegroundColor Yellow
    Read-Host "按回车退出"
    exit 1
}

$ips = Get-LocalIPs
Write-Host ""
Write-Host "  ============================================" -ForegroundColor Green
Write-Host "   WPS 执行端已启动" -ForegroundColor Green
Write-Host "  ============================================" -ForegroundColor Green
Write-Host "   端口      : $Port"
Write-Host "   COM 桥    : $ComPath"
Write-Host "  允许脚本  : $AllowScript"
if ($Token) { Write-Host "  口令      : 需要 X-Token 头" }
Write-Host ""
Write-Host "   在手机/浏览器里打开下面任意一个地址验证：" -ForegroundColor Cyan
foreach ($ip in $ips) { Write-Host "     http://${ip}:$Port/ping" -ForegroundColor Cyan }
Write-Host "     http://127.0.0.1:$Port/ping   （本机）"
Write-Host ""
Write-Host "   这个窗口不要关。按 Ctrl+C 停止。" -ForegroundColor DarkGray
Write-Host ""

# 把当前地址写到桌面，IP 变动时重开脚本即可刷新
$addrFile = Join-Path ([Environment]::GetFolderPath('Desktop')) 'wps-地址.txt'
try {
    $L = @()
    $L += 'WPS 执行端 —— 手机侧要用的地址'
    $L += ''
    $L += ('端口: ' + $Port)
    $L += ''
    $L += '局域网地址（在手机上执行：bash wps.sh set 地址）:'
    foreach ($ip in $ips) { $L += ('  ' + $ip + ':' + $Port) }
    $L += ''
    $L += ('更新时间: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    $L += ''
    $L += 'IP 变了就重开这个脚本，本文件会自动更新。'
    $L += '换端口启动: powershell -ExecutionPolicy Bypass -File wps-exec.ps1 -Port 9000 -AllowScript'
    $L -join "`r`n" | Set-Content -LiteralPath $addrFile -Encoding UTF8
    Write-Host ("   地址已写入桌面: " + $addrFile) -ForegroundColor DarkGray
    Write-Host ""
} catch {}

Write-Log "启动 port=$Port com=$ComPath allowScript=$AllowScript ips=$($ips -join ',')"

# 写 pid 文件，供 WPS 加载项「停止服务」按钮使用
$script:PidFile = Join-Path $env:TEMP "wps-exec-$Port.pid"
try { Set-Content -LiteralPath $script:PidFile -Value $PID -Encoding ASCII } catch {}

try {
    while ($listener.IsListening) {
        $ctx = $listener.GetContext()
        $req = $ctx.Request
        $path = $req.Url.AbsolutePath.TrimEnd('/')
        if ($path -eq '') { $path = '/' }
        $method = $req.HttpMethod
        Write-Log "$method $path from $($req.RemoteEndPoint)"

        try {
            # 口令校验（设了才校验；/ping 永远放行，便于探活）
            if ($Token -and $path -ne '/ping' -and $path -ne '/') {
                $got = $req.Headers['X-Token']
                if ($got -ne $Token) { Send-Json $ctx @{ ok = $false; error = 'bad token' } 403; continue }
            }

            switch ($path) {
                '/ping' {
                    Send-Json $ctx ([ordered]@{
                        ok = $true
                        service = 'wps-exec'
                        port = $Port
                        host = $env:COMPUTERNAME
                        user = $env:USERNAME
                        ips = $ips
                        comPath = $ComPath
                        allowScript = [bool]$AllowScript
                        time = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    })
                }
                '/' {
                    $html = @"
<!doctype html><meta charset="utf-8"><title>WPS 执行端</title>
<style>body{font:15px/1.7 -apple-system,'Microsoft YaHei',sans-serif;max-width:640px;margin:60px auto;padding:0 20px;color:#222}code{background:#f3f3f3;padding:2px 6px;border-radius:4px}</style>
<h2>WPS 执行端已启动</h2>
<p>这台机器：<code>$env:COMPUTERNAME</code></p>
<p>接口：</p>
<ul>
<li><code>GET /ping</code> — 探活</li>
<li><code>POST /wps</code> — <code>{"action":"getAppInfo","params":{}}</code></li>
<li><code>POST /script</code> — 直接执行 PowerShell$(if ($AllowScript) { '' } else { '（未开启）' })</li>
</ul>
<p>看到这一页就说明手机能连到这台电脑了。</p>
"@
                    Send-Bytes $ctx ([Text.Encoding]::UTF8.GetBytes($html)) 'text/html; charset=utf-8' 200
                }
                '/diag' {
                    # 只读诊断：在子进程里（与 /wps 同条件）探查 COM，把真实错误带回来
                    $code = @'
$ErrorActionPreference = 'Continue'
$o = [ordered]@{}
$o.apartment = [System.Threading.Thread]::CurrentThread.GetApartmentState().ToString()
$o.psver = $PSVersionTable.PSVersion.ToString()
$o.session = (Get-Process -Id $PID).SessionId
$o.user = $env:USERNAME
$o.bits = [Environment]::Is64BitProcess
foreach ($id in @('Kwps.Application','KWPS.Application','Word.Application','Ket.Application','KET.Application','Kwpp.Application','Excel.Application')) {
    try {
        $t = [Type]::GetTypeFromProgID($id)
        if (-not $t) { $o[$id] = 'progid-not-found'; continue }
        $x = [Activator]::CreateInstance($t)
        $o[$id] = 'OK ' + $x.GetType().FullName
        try { $x.Quit() } catch {}
    } catch {
        $o[$id] = 'FAIL ' + $_.Exception.Message
    }
}
$o.hkcu_Kwps = [bool](Test-Path 'HKCU:\Software\Classes\Kwps.Application')
$o.hklm_Kwps = [bool](Test-Path 'HKLM:\Software\Classes\Kwps.Application')
$o.wpsProcs = @(Get-Process | Where-Object { $_.ProcessName -match 'wps|et|wpp|ksolaunch' } | ForEach-Object { $_.ProcessName + ':' + $_.Id })
$o | ConvertTo-Json -Compress
'@
                    $tmp = [IO.Path]::Combine([IO.Path]::GetTempPath(), ('wpsdiag_' + [Guid]::NewGuid().ToString('N') + '.ps1'))
                    [IO.File]::WriteAllText($tmp, $code, (New-Object System.Text.UTF8Encoding($true)))
                    $out = ''
                    try { $out = (& powershell -NoProfile -ExecutionPolicy Bypass -File $tmp 2>&1 | Out-String).Trim() }
                    catch { $out = 'ERR ' + $_.Exception.Message }
                    finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
                    Send-Json $ctx @{ ok = $true; diag = $out }
                }
                '/wps' {
                    if ($method -ne 'POST') { Send-Json $ctx @{ ok = $false; error = 'POST only' } 405; continue }
                    $raw = Read-Body $ctx
                    $obj = $null
                    try { $obj = $raw | ConvertFrom-Json } catch { Send-Json $ctx @{ ok = $false; error = 'bad json'; raw = $raw } 400; continue }
                    if (-not $obj.action) { Send-Json $ctx @{ ok = $false; error = 'missing action' } 400; continue }
                    $payload = @{ action = [string]$obj.action; params = $obj.params }
                    $r = Invoke-Worker $payload 180000
                    Send-Json $ctx $r
                }
                '/script' {
                    if (-not $AllowScript) { Send-Json $ctx @{ ok = $false; error = 'script disabled; start with -AllowScript' } 403; continue }
                    if ($method -ne 'POST') { Send-Json $ctx @{ ok = $false; error = 'POST only' } 405; continue }
                    $raw = Read-Body $ctx
                    $obj = $null
                    try { $obj = $raw | ConvertFrom-Json } catch { Send-Json $ctx @{ ok = $false; error = 'bad json' } 400; continue }
                    if (-not $obj.code) { Send-Json $ctx @{ ok = $false; error = 'missing code' } 400; continue }
                    $to = 120000
                    try { if ($obj.timeoutMs) { $to = [int]$obj.timeoutMs } } catch {}
                    $r = Invoke-Worker @{ code = [string]$obj.code } $to
                    Send-Json $ctx $r
                }
                '/ps' {
                    # 应急通道：在一次性子进程里跑任意 PowerShell，不经过常驻 worker。
                    # worker 崩掉时用它做现场取证（读日志、验语法、单步复现）。
                    if (-not $AllowScript) { Send-Json $ctx @{ ok = $false; error = 'script disabled; start with -AllowScript' } 403; continue }
                    if ($method -ne 'POST') { Send-Json $ctx @{ ok = $false; error = 'POST only' } 405; continue }
                    $raw = Read-Body $ctx
                    $obj = $null
                    try { $obj = $raw | ConvertFrom-Json } catch { Send-Json $ctx @{ ok = $false; error = 'bad json'; raw = $raw } 400; continue }
                    $code = [string]$obj.code
                    if (-not $code) { Send-Json $ctx @{ ok = $false; error = 'missing code' } 400; continue }
                    $to = 60000
                    try { if ($obj.timeoutMs) { $to = [int]$obj.timeoutMs } } catch {}
                    $rr = Invoke-ChildScript $code $to
                    Send-Json $ctx @{ ok = $true; output = ([string]$rr.out).Trim(); timedOut = [bool]$rr.timedOut }
                }
                '/workerlog' {
                    # 看常驻 worker 的 stdout/stderr 尾巴
                    $dir = $script:WorkerRoot
                    $res = [ordered]@{ root = $dir; alive = $false; pid = $script:WorkerPid }
                    if ($script:WorkerPid) {
                        $res.alive = [bool](Get-Process -Id $script:WorkerPid -ErrorAction SilentlyContinue)
                    }
                    foreach ($n in @('worker.err', 'worker.out')) {
                        $fp = Join-Path $dir $n
                        if (Test-Path -LiteralPath $fp) {
                            $t = ''
                            try { $t = [IO.File]::ReadAllText($fp, [Text.Encoding]::Default) } catch {}
                            if ($t.Length -gt 8000) { $t = $t.Substring($t.Length - 8000) }
                            $res[$n] = $t.Trim()
                        } else {
                            $res[$n] = $null
                        }
                    }
                    Send-Json $ctx $res
                }
                default {
                    Send-Json $ctx @{ ok = $false; error = 'not found'; path = $path } 404
                }
            }
        } catch {
            try { Send-Json $ctx @{ ok = $false; error = $_.Exception.Message } 500 } catch {}
            Write-Log "处理出错：$($_.Exception.Message)"
        }
    }
} finally {
    try { $listener.Stop(); $listener.Close() } catch {}
    if ($script:PidFile) { Remove-Item -LiteralPath $script:PidFile -Force -ErrorAction SilentlyContinue }
    Write-Log "已停止"
}
