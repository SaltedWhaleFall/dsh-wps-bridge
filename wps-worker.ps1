param([string]$Root = '', [int]$PollMs = 120)
$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

if (-not $Root) { $Root = Join-Path $env:TEMP 'wps-worker' }
$ReqDir  = Join-Path $Root 'req'
$RespDir = Join-Path $Root 'resp'
New-Item -ItemType Directory -Force -Path $ReqDir, $RespDir | Out-Null
Get-ChildItem $ReqDir  -Filter *.json -EA 0 | Remove-Item -Force -EA 0
Get-ChildItem $RespDir -Filter *.json -EA 0 | Remove-Item -Force -EA 0

$global:WPS = @{}

function Get-WpsExe {
    foreach ($k in @(
        'HKCU:\Software\Classes\WOW6432Node\CLSID\{45540001-5750-5300-4B49-4E47534F4655}\LocalServer32',
        'HKCU:\Software\Classes\CLSID\{45540001-5750-5300-4B49-4E47534F4655}\LocalServer32')) {
        if (Test-Path $k) {
            $v = (Get-ItemProperty $k).'(default)'
            if ($v -match '"([^"]+)"') { return $Matches[1] }
            return (($v -split ' ')[0] -replace '"','')
        }
    }
    return (Join-Path $env:LOCALAPPDATA 'Kingsoft\WPS Office\12.1.0.29690\office6\wps.exe')
}

$WPS_EXE = Get-WpsExe
$CFG = @{
    excel = @{ Clsid = '{45540001-5750-5300-4B49-4E47534F4655}'; Arg = '/et';  ProgId = @('KET.Application', 'Excel.Application'); Proc = 'et' }
    word  = @{ Clsid = '{000209FF-0000-4B30-A977-D214852036FF}'; Arg = '/wps'; ProgId = @('KWPS.Application', 'Word.Application'); Proc = 'wps' }
    ppt   = @{ Clsid = '{44720441-94BF-4940-926D-4F38FECF2A48}'; Arg = '/wpp'; ProgId = @('KWPP.Application', 'PowerPoint.Application'); Proc = 'wpp' }
}

$OleSig = '[DllImport("oleaut32.dll", PreserveSig=false)] public static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out object ppunk);'

function Get-Com-ByClsid([string]$clsid) {
    if (-not ('Win32.Ole2' -as [type])) { Add-Type -Namespace 'Win32' -Name 'Ole2' -MemberDefinition $OleSig }
    $g = [Guid]$clsid
    $obj = $null
    [Win32.Ole2]::GetActiveObject([ref]$g, [IntPtr]::Zero, [ref]$obj)
    return $obj
}

function Get-App([string]$kind) {
    $a = $null
    if ($global:WPS.ContainsKey($kind)) { $a = $global:WPS[$kind] }
    if ($a) { try { $null = $a.Name; return $a } catch { $global:WPS[$kind] = $null } }
    $c = $CFG[$kind]
    if (-not $c) { throw "unknown app: $kind" }
    $obj = $null

    # 1) ROT 里已有实例就复用。这条走 oleaut32 的 GetActiveObject(CLSID)：
    #    本机 WPS 只注册了 ProgID、CLSID 键缺失，Marshal.GetActiveObject(ProgID) 一律失效。
    try { $obj = Get-Com-ByClsid $c.Clsid } catch { $obj = $null }
    if ($obj) { $global:WPS[$kind] = $obj; return $obj }

    # 2) CLSID 已注册、暂无实例：直接建
    for ($k = 0; $k -lt 2; $k++) {
        try { $obj = [Activator]::CreateInstance([Type]::GetTypeFromCLSID($c.Clsid)); break }
        catch { Start-Sleep -Milliseconds 400 }
    }
    if ($obj) { $global:WPS[$kind] = $obj; return $obj }

    # 3) CLSID 未注册（没有任何 /Automation 实例在跑）：起一个，等它自己注册上来
    Start-Process $WPS_EXE -ArgumentList '/prometheus', $c.Arg, '/Automation'
    for ($k = 0; $k -lt 12; $k++) {
        Start-Sleep -Milliseconds 800
        try { $obj = [Activator]::CreateInstance([Type]::GetTypeFromCLSID($c.Clsid)) } catch { $obj = $null }
        if (-not $obj) { try { $obj = Get-Com-ByClsid $c.Clsid } catch { $obj = $null } }
        if ($obj) { break }
    }
    if (-not $obj) { throw "cannot create COM app: $kind" }
    $global:WPS[$kind] = $obj
    return $obj
}

function Get-Sheet($xl, $name) {
    $wb = $xl.ActiveWorkbook
    if (-not $wb) { throw 'no active workbook (先 newWorkbook 或 open)' }
    if ([string]::IsNullOrEmpty([string]$name)) { return $wb.Sheets.Item(1) }
    return $wb.Sheets.Item([string]$name)
}

function Get-Matrix($rng) {
    $rows = [int]$rng.Rows.Count; $cols = [int]$rng.Columns.Count
    $out = New-Object System.Collections.ArrayList
    for ($i = 1; $i -le $rows; $i++) {
        $line = New-Object System.Collections.ArrayList
        for ($j = 1; $j -le $cols; $j++) { [void]$line.Add($rng.Cells.Item($i, $j).Value2) }
        [void]$out.Add($line.ToArray())
    }
    return ,$out.ToArray()
}

function New-Workbook($App) {
    # WPS 的 Workbooks.Add 返回值不可靠（既能返回 $null 又真的建出簿），
    # 所以判定一律看 $App.Workbooks.Count 有没有涨，不看返回值。
    try { $null = $App.Visible } catch {}
    $before = 0
    try { $before = [int]$App.Workbooks.Count } catch { $before = 0 }
    for ($attempt = 0; $attempt -lt 4; $attempt++) {
        try {
            if ($attempt % 2 -eq 0) { $null = $App.Workbooks.Add([Type]::Missing) }
            else { $null = $App.Workbooks.Add() }
        } catch {
            try { $null = $App.Workbooks.Add() } catch {}
        }
        for ($i = 0; $i -lt 16; $i++) {
            Start-Sleep -Milliseconds 250
            $cnt = 0
            try { $cnt = [int]$App.Workbooks.Count } catch { $cnt = 0 }
            if ($cnt -le $before) { continue }
            $wb = $null
            try { $wb = $App.Workbooks.Item($cnt) } catch { $wb = $null }
            if ($wb) {
                $sc = $null
                try { $sc = $wb.Sheets } catch { $sc = $null }
                if ($sc) {
                    $ok = $false
                    try { $ok = ([int]$sc.Count -ge 1) } catch { $ok = $false }
                    if ($ok) { return $wb }
                }
            }
        }
    }
    throw ('cannot create workbook (Workbooks.Count stayed ' + $before + ' after 4 attempts)')
}

function Do-Action([string]$action, $p) {
    if ($null -eq $p) { $p = @{} }
    switch ($action) {
        'appInfo' {
            $o = [ordered]@{}
            foreach ($k in @('excel','word','ppt')) {
                $v = $null
                if ($global:WPS.ContainsKey($k) -and $global:WPS[$k]) {
                    try { $v = $global:WPS[$k].Name } catch { $v = $null }
                }
                $o[$k] = $v
            }
            $o['exe'] = $WPS_EXE
            $o['pid'] = $PID
            return $o
        }
        'ensure' {
            $k = if ($p.kind) { [string]$p.kind } else { 'excel' }
            $a = Get-App $k
            return @{ kind = $k; name = $a.Name }
        }
        'sheetList' {
            $xl = Get-App 'excel'; $wb = $xl.ActiveWorkbook
            if (-not $wb) { throw 'no active workbook' }
            $arr = @()
            for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $arr += $wb.Sheets.Item($i).Name }
            return @{ workbook = $wb.Name; sheets = $arr; count = $wb.Sheets.Count; active = $xl.ActiveSheet.Name }
        }
        'newWorkbook' {
            $xl = Get-App 'excel'; $wb = New-Workbook $xl
            return @{ name = $wb.Name; sheets = $wb.Sheets.Count }
        }
        'open' {
            $kind = if ($p.kind) { [string]$p.kind } else { 'excel' }
            $a = Get-App $kind
            $path = [string]$p.path
            if (-not (Test-Path -LiteralPath $path)) { throw "file not found: $path" }
            $wb = $a.Workbooks.Open($path)
            return @{ name = $wb.Name; path = $path }
        }
        'getCell' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $c = $sh.Cells.Item([int]$p.row, [int]$p.col)
            return @{ row = [int]$p.row; col = [int]$p.col; value = $c.Value2; text = $c.Text }
        }
        'setCell' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $v = $p.value
            $cell = $sh.Cells.Item([int]$p.row, [int]$p.col)
            if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
                $cell.Formula = [string]$v
            } else {
                $cell.Value2 = $v
            }
            return @{ ok = $true; at = ('R{0}C{1}' -f [int]$p.row, [int]$p.col); wrote = $v }
        }
        'getRange' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $r = $sh.Range([string]$p.range)
            $m = Get-Matrix $r
            return @{ range = [string]$p.range; rows = $r.Rows.Count; cols = $r.Columns.Count; data = $m }
        }
        'setRange' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $m = $p.data
            $rows = @($m).Count
            $cols = if ($rows -gt 0) { @($m[0]).Count } else { 0 }
            for ($i = 0; $i -lt $rows; $i++) {
                for ($j = 0; $j -lt $cols; $j++) {
                    $v = @($m[$i])[$j]
                    $cell = $sh.Range([string]$p.range).Cells.Item($i + 1, $j + 1)
                    if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
                        $cell.Formula = [string]$v
                    } else {
                        $cell.Value2 = [string]$v
                    }
                }
            }
            return @{ ok = $true; rows = $rows; cols = $cols }
        }
        'save' {
            $xl = Get-App 'excel'; $wb = $xl.ActiveWorkbook
            if (-not $wb) { throw 'no active workbook' }
            $wb.Save(); return @{ ok = $true; name = $wb.Name; path = $wb.FullName }
        }
        'saveAs' {
            $xl = Get-App 'excel'; $wb = $xl.ActiveWorkbook
            if (-not $wb) { throw 'no active workbook' }
            $path = [string]$p.path
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
            try { $wb.SaveAs($path) }
            catch { $wb.SaveAs($path, 51) }
            return @{ ok = $true; path = $path; exists = (Test-Path -LiteralPath $path); size = (Get-Item -LiteralPath $path).Length }
        }
        'toPdf' {
            $kind = if ($p.kind) { [string]$p.kind } else { 'excel' }
            $a = Get-App $kind
            $out = if ($p.path) { [string]$p.path } elseif ($p.outputPath) { [string]$p.outputPath } else { '' }
            if (-not $out) { throw 'toPdf 需要 path 或 outputPath 参数' }
            if ($kind -eq 'excel') { $a.ActiveWorkbook.ExportAsFixedFormat(0, $out) }
            elseif ($kind -eq 'word') { $a.ActiveDocument.ExportAsFixedFormat($out, 17) }
            else { $a.ActivePresentation.SaveAs($out, 32) }
            return @{ ok = $true; path = $out; exists = (Test-Path -LiteralPath $out) }
        }
        'findReplace' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $sh.Cells.Replace([string]$p.find, [string]$p.replace) | Out-Null
            return @{ ok = $true }
        }
        'wordText' {
            $w = Get-App 'word'
            $doc = $w.ActiveDocument
            if (-not $doc) { throw 'no active document' }
            return @{ name = $doc.Name; text = $doc.Content.Text }
        }
        'wordOpen' {
            $w = Get-App 'word'
            $doc = $w.Documents.Open([string]$p.path)
            return @{ name = $doc.Name; path = [string]$p.path }
        }
        'wordAppend' {
            $w = Get-App 'word'
            $doc = $w.ActiveDocument
            if (-not $doc) { throw 'no active document' }
            $rng = $doc.Content
            $rng.Collapse(0) | Out-Null
            $rng.InsertAfter([string]$p.text)
            return @{ ok = $true }
        }
        'wordSave' {
            $w = Get-App 'word'
            $doc = $w.ActiveDocument
            if (-not $doc) { throw 'no active document' }
            $doc.Save()
            return @{ ok = $true; path = $doc.FullName }
        }
        'pptInfo' {
            $pp = Get-App 'ppt'
            $pres = $pp.ActivePresentation
            if (-not $pres) { throw 'no active presentation' }
            return @{ name = $pres.Name; slides = $pres.Slides.Count }
        }
        # ---- Excel 结构/格式补充（移植自官方 wps-com.ps1）----
        'insertRows' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $r0 = [int]$p.row; if ($r0 -lt 1) { throw 'row required' }
            $n = if ($p.count) { [int]$p.count } else { 1 }
            $sh.Range("${r0}:$($r0 + $n - 1)").Insert()
            return @{ insertedAt = $r0; count = $n }
        }
        'deleteRows' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $r0 = [int]$p.row; if ($r0 -lt 1) { throw 'row required' }
            $n = if ($p.count) { [int]$p.count } else { 1 }
            $sh.Range("${r0}:$($r0 + $n - 1)").Delete()
            return @{ deletedAt = $r0; count = $n }
        }
        'insertColumns' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $c = [string]$p.column
            $n = if ($p.count) { [int]$p.count } else { 1 }
            for ($i = 0; $i -lt $n; $i++) { $sh.Range("${c}:$c").EntireColumn.Insert() }
            return @{ column = $c; count = $n }
        }
        'deleteColumns' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $c = [string]$p.column
            $n = if ($p.count) { [int]$p.count } else { 1 }
            for ($i = 0; $i -lt $n; $i++) { $sh.Range("${c}:$c").EntireColumn.Delete() }
            return @{ column = $c; count = $n }
        }
        'setBorder' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $rg = $sh.Range([string]$p.range)
            $map = @{ thin = 1; medium = 2; thick = 4; double = 6; none = 0 }
            $w = $map[[string]$p.style]; if ($null -eq $w) { $w = 1 }
            $pos = if ($p.position) { [string]$p.position } else { 'all' }
            $idx = @()
            if ($pos -eq 'all' -or $pos -eq 'outside') { $idx += 7, 8, 9, 10 }
            if ($pos -eq 'all' -or $pos -eq 'inside') { $idx += 11, 12 }
            if ($pos -eq 'left') { $idx += 7 }
            if ($pos -eq 'top') { $idx += 8 }
            if ($pos -eq 'bottom') { $idx += 9 }
            if ($pos -eq 'right') { $idx += 10 }
            foreach ($i in $idx) {
                $b = $rg.Borders.Item($i)
                if ($w -eq 0) { $b.LineStyle = -4142 } else { $b.LineStyle = 1; $b.Weight = $w }
            }
            return @{ range = [string]$p.range; edges = $idx.Count; weight = $w }
        }
        'setNumberFormat' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $rg = $sh.Range([string]$p.range)
            $f = if ($p.format) { [string]$p.format } else { [string]$p.numberFormat }
            if (-not $f) { throw 'format required' }
            $rg.NumberFormat = $f
            return @{ range = [string]$p.range; format = $f }
        }
        'setColumnWidth' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $c = [string]$p.column
            $sh.Range("${c}:$c").ColumnWidth = [double]$p.width
            return @{ column = $c; width = [double]$p.width }
        }
        'setRowHeight' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $r0 = [int]$p.row
            $sh.Range("${r0}:$r0").RowHeight = [double]$p.height
            return @{ row = $r0; height = [double]$p.height }
        }
        'autoFitColumn' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            if ($p.range) { $sh.Range([string]$p.range).Columns.AutoFit() }
            elseif ($p.column) { $c = [string]$p.column; $sh.Range("${c}:$c").EntireColumn.AutoFit() }
            else { $sh.UsedRange.Columns.AutoFit() }
            return @{ ok = $true }
        }
        'freezePanes' {
            $app = Get-App 'excel'
            $sh = $app.ActiveSheet
            if ($p.cell) { $sh.Range([string]$p.cell).Select() }
            elseif ($p.row -and $p.column) { $sh.Range("$([string]$p.column)$([int]$p.row)").Select() }
            $app.ActiveWindow.FreezePanes = $true
            return @{ frozen = $true }
        }
        'unfreezePanes' {
            $app = Get-App 'excel'
            $app.ActiveWindow.FreezePanes = $false
            return @{ frozen = $false }
        }
        'sortRange' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $rg = $sh.Range([string]$p.range)
            $key = $sh.Range([string]$p.keyColumn)
            $ord = if ([string]$p.order -eq 'desc') { 2 } else { 1 }
            $rg.Sort($key, $ord)
            return @{ range = [string]$p.range; order = if ($ord -eq 2) { 'desc' } else { 'asc' } }
        }
        'findInSheet' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $rg = if ($p.range) { $sh.Range([string]$p.range) } else { $sh.UsedRange }
            $what = [string]$p.text
            if (-not $what) { $what = [string]$p.searchText }
            if (-not $what) { throw 'text required' }
            $lookAt = if ($p.matchCase) { 1 } else { 2 }
            # WPS 的 Range.Address 不可靠（方法/属性两套签名都会抛）：优先用 Row/Column 自拼
            $colName = {
                param($n)
                $s = ''
                $k = [int]$n
                while ($k -gt 0) {
                    $m = ($k - 1) % 26
                    $s = [string][char](65 + $m) + $s
                    $k = [int](($k - $m) / 26)
                }
                return $s
            }
            $addr = {
                param($r)
                try { $c = [int]$r.Column; $w = [int]$r.Row; if ($c -ge 1 -and $w -ge 1) { return ((& $colName $c) + [string]$w) } } catch {}
                foreach ($t in @(@(), @($false, $false), @($true), @($true, $true))) {
                    try { $a = [string]($r.Address.Invoke($t)); if ($a) { return ($a -replace '\$', '') } } catch {}
                }
                try { $a = [string]$r.Address; if ($a) { return ($a -replace '\$', '') } } catch {}
                return '?'
            }
            $val = {
                param($r)
                try { return $r.Value2 } catch {}
                try { return $r.Text } catch {}
                return ''
            }
            $out = @()
            $hit = $null
            $findErr = @()
            foreach ($sig in @(
                    @($what, [Type]::Missing, -4163, $lookAt),
                    @($what, [Type]::Missing, [Type]::Missing, $lookAt),
                    @($what))) {
                try { $hit = $rg.Find.Invoke($sig); break } catch { $findErr += $_.Exception.Message }
            }
            if ($hit) {
                $first = & $addr $hit
                $guard = 0
                do {
                    $a = & $addr $hit
                    $out += @{ address = ($a -replace '\$', ''); value = (& $val $hit) }
                    try { $hit = $rg.FindNext($hit) } catch { $hit = $null }
                    $guard++
                    if ($null -eq $hit -or $guard -gt 5000) { break }
                } while ((& $addr $hit) -ne $first)
                return @{ count = $out.Count; hits = $out; mode = 'find' }
            }
            # Find 不可用：退到逐格扫描（上限 50 万格，避免卡死 worker）
            $rows = [int]$rg.Rows.Count; $cols = [int]$rg.Columns.Count
            $scanned = 0
            if (($rows * $cols) -le 500000) {
                for ($i = 1; $i -le $rows; $i++) {
                    for ($j = 1; $j -le $cols; $j++) {
                        $cell = $rg.Cells.Item($i, $j)
                        $tv = [string](& $val $cell)
                        $m = if ($p.matchCase) { $tv.Contains($what) } else { $tv.IndexOf($what, [StringComparison]::OrdinalIgnoreCase) -ge 0 }
                        if ($m) { $out += @{ address = ((& $addr $cell) -replace '\$', ''); value = (& $val $cell) } }
                        $scanned++
                    }
                }
            }
            $res = @{ count = $out.Count; hits = $out; mode = 'scan'; scanned = $scanned }
            if ($findErr.Count -gt 0) { $res['findError'] = $findErr[0] }
            return $res
        }
        'copyRange' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $sh.Range([string]$p.range).Copy()
            return @{ copied = [string]$p.range }
        }
        'pasteRange' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $dst = $sh.Range([string]$p.destination)
            switch ([string]$p.pasteType) {
                'values'  { $dst.PasteSpecial(-4163) }
                'formats' { $dst.PasteSpecial(-4122) }
                'all'     { $sh.Paste() }
                default   { $sh.Paste() }
            }
            return @{ pastedTo = [string]$p.destination }
        }
        'copyRowFormat' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $src = [int]$p.fromRow
            $rows = @()
            if ($p.toRows) { $rows = @($p.toRows | ForEach-Object { [int]$_ }) }
            elseif ($p.toRow) { $rows = @([int]$p.toRow) }
            if ($rows.Count -eq 0) { throw 'toRow or toRows required' }
            $sh.Range("${src}:$src").Copy()
            foreach ($t in $rows) { $sh.Range("${t}:$t").PasteSpecial(-4122) }
            $sh.Application.CutCopyMode = 0
            return @{ fromRow = $src; toRows = $rows; copied = $rows.Count }
        }
        'createSheet' {
            $wb = (Get-App 'excel').ActiveWorkbook
            $name = if ($p.name) { [string]$p.name } else { 'Sheet' + ($wb.Sheets.Count + 1) }
            $sh = $wb.Sheets.Add($wb.Sheets.Item($wb.Sheets.Count))
            $sh.Name = $name
            return @{ name = $sh.Name; count = $wb.Sheets.Count }
        }
        'deleteSheet' {
            $wb = (Get-App 'excel').ActiveWorkbook
            $name = [string]$p.name; if (-not $name) { $name = [string]$p.sheet }
            $wb.Sheets.Item($name).Delete()
            return @{ deleted = $name; count = $wb.Sheets.Count }
        }
        'renameSheet' {
            $wb = (Get-App 'excel').ActiveWorkbook
            $old = [string]$p.oldName; if (-not $old) { $old = [string]$p.sheet }
            $wb.Sheets.Item($old).Name = [string]$p.newName
            return @{ from = $old; to = [string]$p.newName }
        }
        'switchSheet' {
            $wb = (Get-App 'excel').ActiveWorkbook
            $name = [string]$p.sheet; if (-not $name) { $name = [string]$p.name }
            $wb.Sheets.Item($name).Activate()
            return @{ active = $name }
        }
        'getFormula' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $rg = $sh.Range([string]$p.range)
            $o = @()
            foreach ($c in $rg.Cells) { $o += @{ address = ($c.Address() -replace '\$', ''); formula = $c.Formula } }
            return @{ cells = $o }
        }
        'setFormula' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $sh.Range([string]$p.range).Formula = [string]$p.formula
            return @{ range = [string]$p.range; done = $true }
        }
        'getSelection' {
            $app = Get-App 'excel'
            $sel = $app.Selection
            return @{ address = ($sel.Address() -replace '\$', ''); value = $sel.Value2; rows = [int]$sel.Rows.Count; cols = [int]$sel.Columns.Count }
        }
        'setPrintArea' {
            $xl = Get-App 'excel'; $sh = Get-Sheet $xl $p.sheet
            $sh.PageSetup.PrintArea = [string]$p.range
            return @{ printArea = [string]$p.range }
        }
        'getOpenWorkbooks' {
            $app = Get-App 'excel'
            $o = @()
            foreach ($wb in $app.Workbooks) { $o += @{ name = $wb.Name; path = $wb.FullName; sheets = $wb.Sheets.Count } }
            return @{ count = $o.Count; workbooks = $o }
        }
        'switchWorkbook' {
            $app = Get-App 'excel'
            $app.Workbooks.Item([string]$p.name).Activate()
            return @{ active = [string]$p.name }
        }
        'getDocumentStats' {
            $app = Get-App 'excel'
            $wb = $app.ActiveWorkbook
            return @{ workbook = $wb.Name; sheets = $wb.Sheets.Count; used = $wb.ActiveSheet.UsedRange.Address() }
        }
        'selfTest' {
            $r = New-Object System.Collections.ArrayList
            $okN = 0; $failN = 0
            $tmp = Join-Path ([IO.Path]::GetTempPath()) ("selftest-{0}.xlsx" -f (Get-Date -Format 'HHmmss'))
            $cw = $null
            $appE = Get-App 'excel'
            $cw = New-Workbook $appE
            $sh = $cw.Sheets.Item(1)
            $sh.Name = 'ST'
            $cases = @(
                @('setCell-str',   { Do-Action 'setCell' @{ row=1; col=1; value='测试中文'; sheet='ST' } }),
                @('setCell-num',   { Do-Action 'setCell' @{ row=1; col=2; value=2026; sheet='ST' } }),
                @('setCell-dec',   { Do-Action 'setCell' @{ row=1; col=3; value=31.5; sheet='ST' } }),
                @('getCell',       { Do-Action 'getCell' @{ row=1; col=1; sheet='ST' } }),
                @('getRange',      { Do-Action 'getRange' @{ range='A1:C1'; sheet='ST' } }),
                @('setRange',      { Do-Action 'setRange' @{ range='A3:C3'; data=@(@('甲','2026-11-04',1),@('乙','2026-11-05',2)); sheet='ST' } }),
                @('insertRows',    { Do-Action 'insertRows' @{ row=2; count=1; sheet='ST' } }),
                @('deleteRows',    { Do-Action 'deleteRows' @{ row=2; count=1; sheet='ST' } }),
                @('insertColumns', { Do-Action 'insertColumns' @{ column='D'; count=1; sheet='ST' } }),
                @('deleteColumns', { Do-Action 'deleteColumns' @{ column='D'; count=1; sheet='ST' } }),
                @('setBorder',     { Do-Action 'setBorder' @{ range='A1:C3'; style='thin'; position='all'; sheet='ST' } }),
                @('setNumberFormat', { Do-Action 'setNumberFormat' @{ range='A1:C1'; format='General'; sheet='ST' } }),
                @('setColumnWidth', { Do-Action 'setColumnWidth' @{ column='A'; width=18; sheet='ST' } }),
                @('setRowHeight',  { Do-Action 'setRowHeight' @{ row=1; height=20; sheet='ST' } }),
                @('autoFitColumn', { Do-Action 'autoFitColumn' @{ column='C'; sheet='ST' } }),
                @('freezePanes',   { Do-Action 'freezePanes' @{ cell='A2'; sheet='ST' } }),
                @('unfreezePanes', { Do-Action 'unfreezePanes' @{ sheet='ST' } }),
                @('findInSheet',   { Do-Action 'findInSheet' @{ text='测试'; sheet='ST' } }),
                @('getFormula',    { Do-Action 'getFormula' @{ range='C1'; sheet='ST' } }),
                @('setFormula',    { Do-Action 'setFormula' @{ range='C2'; formula='=B1*2'; sheet='ST' } }),
                @('copyRowFormat', { Do-Action 'copyRowFormat' @{ fromRow=1; toRows=@(2,3); sheet='ST' } }),
                @('sortRange',     { Do-Action 'sortRange' @{ range='A1:C3'; keyColumn='A1'; order='asc'; sheet='ST' } }),
                @('getSelection',  { Do-Action 'getSelection' @{} }),
                @('getDocumentStats', { Do-Action 'getDocumentStats' @{} }),
                @('getOpenWorkbooks', { Do-Action 'getOpenWorkbooks' @{} }),
                @('setPrintArea',  { Do-Action 'setPrintArea' @{ range='A1:C3'; sheet='ST' } }),
                @('createSheet',   { Do-Action 'createSheet' @{ name='ST2' } }),
                @('renameSheet',   { Do-Action 'renameSheet' @{ oldName='ST2'; newName='ST2b' } }),
                @('switchSheet',   { Do-Action 'switchSheet' @{ sheet='ST2b' } }),
                @('deleteSheet',   { Do-Action 'deleteSheet' @{ name='ST2b' } }),
                @('saveAs',        { Do-Action 'saveAs' @{ path=$tmp; sheet='ST' } })
            )
            foreach ($c in $cases) {
                try { $null = & $c[1]; $okN++; [void]$r.Add(@{ t = $c[0]; ok = $true }) }
                catch { $failN++; [void]$r.Add(@{ t = $c[0]; ok = $false; err = $_.Exception.Message }) }
            }
            try { $cw.Close($false) } catch {}
            try { if (Test-Path $tmp) { Remove-Item -LiteralPath $tmp -Force } } catch {}
            try { if ($cw) { $cw.Close($false) } } catch {}
            try { if (Test-Path $tmp) { Remove-Item $tmp -Force -EA 0 } } catch {}
            return @{ total = $okN + $failN; pass = $okN; fail = $failN; detail = $r }
        }
        'eval' { return (Invoke-Expression ([string]$p.code)) }
        default { throw "unknown action: $action" }
    }
}

$host.UI.RawUI.WindowTitle = 'WPS Worker (COM STA)'
"WPS worker up. pid=$PID exe=$WPS_EXE root=$Root"

$lastBeat = Join-Path $Root 'worker.alive'
while ($true) {
    try { [IO.File]::WriteAllText($lastBeat, (Get-Date).ToString('s')) } catch {}
    $files = @(Get-ChildItem $ReqDir -Filter *.json -EA 0 | Sort-Object Name)
    foreach ($f in $files) {
        $id = $f.BaseName
        $resp = $null
        $body = $null
        try { $body = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json } catch { $body = $null }
        if ($null -eq $body) {
            $resp = @{ ok = $false; error = 'bad json'; id = $id }
        } else {
            try {
                if ($body.action) {
                    $r = Do-Action ([string]$body.action) $body.params
                    $resp = @{ ok = $true; result = $r; id = $id }
                } elseif ($body.code) {
                    $o = Invoke-Expression ([string]$body.code) 2>&1 | Out-String
                    $resp = @{ ok = $true; output = $o.Trim(); id = $id }
                } else {
                    $resp = @{ ok = $false; error = 'need action or code'; id = $id }
                }
            } catch {
                $pos = ''
                try { $pos = $_.InvocationInfo.PositionMessage } catch {}
                $stk = ''
                try { $stk = $_.ScriptStackTrace } catch {}
                $resp = @{ ok = $false; error = $_.Exception.Message; at = $pos; stack = $stk; id = $id }
            }
        }
        try {
            $json = $resp | ConvertTo-Json -Depth 12 -Compress
            [IO.File]::WriteAllText((Join-Path $RespDir ($id + '.json')), $json, (New-Object Text.UTF8Encoding($false)))
        } catch {}
        Remove-Item -LiteralPath $f.FullName -Force -EA 0
    }
    Start-Sleep -Milliseconds $PollMs
}
