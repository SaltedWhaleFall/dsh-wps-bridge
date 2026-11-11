---
name: wps-doc-ops
version: 1.0.0
category: tools
description: 通过局域网 HTTP 桥远程读写一台 Windows 上正在运行的 WPS（表格 / 文字 / 演示），按 action 精确操作单元格、公式、格式、行高列宽、查找替换与 PDF 导出，并在写之前备份、写之后读回核对。
whenToUse: 当用户要求"把这个录进表里""看看第几行""改一下这个 Excel""批量替换""导出成 PDF""把这段写进 Word"，或者需要在那台跑着 WPS 的 Windows 上执行 PowerShell 时使用。不适用于没有 wps-exec.ps1 在跑的环境；也不适用于需要生成图表、透视表、PPT 版式与动画的请求——那几类动作桥不支持，见「能力边界」。
tags:
  - wps
  - excel
  - word
  - ppt
  - office
  - automation
  - powershell
  - com
disableModelInvocation: false
userInvocable: true
---

# WPS 远程文档操作

隔着局域网操作一台 Windows 上的 WPS。所有动作最终落到那台机器上的一个常驻 PowerShell worker，由它持有 COM 对象。

**为什么必须有 worker**：COM 对象只能在创建它的进程里使用。从新进程 `GetActiveObject` 拿不到同一份实例，改出来的是另一个隐藏的 WPS，用户看不见，保存也没人知道。所以桥一开始就把 COM 关进 worker，外面只发 HTTP。

## 前置：先装电脑端那个插件

这个 skill 只管「怎么用」，不含桥本身。桥是一个单独的 DSH 插件，装在你那台跑 WPS 的 Windows 上（不是手机上）：

- 仓库：`https://github.com/SaltedWhaleFall/dsh-wps-bridge`
- 市场页：`https://deepseek.stream/plugins/dsh-wps-bridge`
- 一键装（在那台 Windows 的 DSH 里）：`cd plugins/_workshop && node link.mjs dsh-wps-bridge`，然后重启宿主
- 桥的进程本身要单独起：双击包里的 `启动WPS桥.bat`（默认 `-Port 8765`，不开 `-AllowScript`）

装完在这边配地址：

```bash
bash wps.sh set http://<那台机器的IP>:8765
bash wps.sh ping          # 期望 {"ok":true,...}
```

地址会跳（校园网 / DHCP），连不上时先 `bash wps.sh find 192.168.1` 扫一遍同网段，别急着重装。

**没装桥、或 bat 没起，这个 skill 里所有命令都会失败**——先确认 `ping` 通，再谈动作。

## 适用场景

- 把一批单据、会议、名单追加进一张已有的表，并按日期/序号重排
- 读某张表的现状（表头、行数、日期分布）再决定怎么改
- 批量替换、统一格式、设列宽行高、冻结窗格、加边框
- 把一个 sheet 或整个工作簿导成 PDF
- 往一个已打开的 Word 里追加段落、读全文、另存

## 能力边界（先看这一段，别承诺做不到的事）

**能做到的 48 个 action**：

| 类别 | action |
|---|---|
| 工作簿 | `appInfo` `getOpenWorkbooks` `open` `newWorkbook` `ensure` `save` `saveAs` `toPdf` |
| 工作表 | `sheetList` `createSheet` `deleteSheet` `renameSheet` `switchSheet` `switchWorkbook` `getDocumentStats` |
| 读取 | `getCell` `getRange` `getFormula` `getSelection` `getOpenWorkbooks` |
| 写入 | `setCell` `setRange` `setFormula` `insertRows` `deleteRows` `insertColumns` `deleteColumns` |
| 格式 | `setBorder` `setNumberFormat` `setColumnWidth` `setRowHeight` `autoFitColumn` `copyRowFormat` `copyRange` `pasteRange` |
| 检索排序 | `findInSheet` `findReplace` `sortRange` |
| 视图/打印 | `freezePanes` `unfreezePanes` `setPrintArea` |
| 文字 | `wordOpen` `wordText` `wordAppend` `wordSave` |
| 演示 | `pptInfo`（**只读**：页数、版式名，改不了） |
| 自检 | `selfTest` `eval` `values` `formats` |

**做不到的**（原始 WPS 助手类技能里常被承诺，这个桥没有）：

- 生成图表（chart）、数据透视表（pivot table）——COM 里没封装
- 编辑 PPT：母版、版式、动画、配色方案——`pptInfo` 只能读
- 公式的"智能生成"：桥只负责把公式写进去，公式本身要你会算
- 任何需要用户在 WPS 界面里点确认的交互

遇到这四类请求，说清楚做不到，给出替代路径（例如导出 CSV 后用 Python 画图、把 PPT 操作步骤写成给人看的清单），不要假装调通了。

## 执行工作流与检查清单

### 第 0 步：探活

```bash
bash wps.sh ping            # 或
bash wps.sh call '{"action":"appInfo","params":{}}'
```

返回里应有 `exe`（wps.exe 路径）与 `pid`。连不上时按这个顺序查：

1. `curl: (52) Empty reply from server` —— 端口有人监听但进程没回：worker 崩了或 WPS 被关了。让用户重启桥（双击 `启动WPS桥.bat`）。
2. `Connection refused` —— 根本没起。同上。
3. IP 变了 —— 校园网/家用 DHCP 会跳。`bash wps.sh find 192.168.1` 扫网段重找。

**不要**在这种情况下改成"我给你一段 VBA 你自己跑"，先确认是桥的问题还是网络的问题。

### 第 1 步：摸清现状，再动手

落笔之前必须先读：

```bash
bash wps.sh call '{"action":"getOpenWorkbooks","params":{}}'        # 有哪些工作簿开着
bash wps.sh call '{"action":"sheetList","params":{}}'               # 有哪些 sheet
bash wps.sh call '{"action":"getDocumentStats","params":{"sheet":"data"}}'   # used 区域，如 "$A$1:$K$539"
bash wps.sh call '{"action":"getRange","params":{"sheet":"data","range":"A1:K3"}}'
```

要改整表（重排序号、插行、批量替换）时，**先把 used 区域整个读出来**，分段读（每段几百行），在本地拼成一份完整快照，后续所有比对都基于这份快照。

### 第 2 步：写之前备份

```bash
# 中文路径必须 base64 传，PowerShell 侧解码后 Copy-Item
B64=$(printf '%s' 'C:\Users\admin\Desktop\wps\目标文件.xlsx' | base64 -w0)
bash wps.sh script "Copy-Item -LiteralPath ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$B64'))) -Destination ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$B64')).Replace('.xlsx','-bak-$(date +%Y%m%d-%H%M).xlsx')) -Force"
```

备完记下原文件字节数，写完之后再比一次——字节数没变说明根本没写进去。

### 第 3 步：改

**单点小改**用 `setCell`（仅限 ASCII；中文会静默丢弃，见下）。**结构化改动一律走 `setRange`**：

```bash
cat > /tmp/payload.json <<'EOF'
{"action":"setRange","params":{"sheet":"data","range":"A1:K539","data":[["日期","序号"],["46082",1]]}}
EOF
curl -s -H 'Content-Type: application/json' --data-binary @/tmp/payload.json \
  http://192.168.1.50:8765/wps
```

`setRange` 是整块覆盖：给的 `data` 有几行就写几行，剩下的原样不动。所以要改一部分，就构造"完整的新区域"，而不是只发改动的行。

改完在本地做一次断言：把非目标行与旧快照逐行比，必须全部相等。相等才发。

### 第 4 步：保存 + 读回核对

```bash
bash wps.sh call '{"action":"save","params":{}}'
bash wps.sh call '{"action":"getDocumentStats","params":{"sheet":"data"}}'
bash wps.sh call '{"action":"getRange","params":{"sheet":"data","range":"A1:K2"}}'
```

`save` 有副作用延迟：返回 `ok:true` 之后文件 mtime 才变。核对三样：`used` 区域对不对、改的行内容对不对、**没改的行还在不在**。

## 输出格式规范与边界条件

### 四个会静默坏数据的坑

**坑 1：`setCell` 传中文会静默丢弃。** 返回 `{"ok":true,"result":{"ok":true,"wrote":"曾燚","at":"R507C9"}}`，读回还是旧值。判断方法：同一个请求换成 ASCII 试一次，能写进去就是编码问题。解法：改用 `setRange` + `--data-binary @file`（UTF-8 JSON 文件），中文正常。

**坑 2：`save` 偶发 `Call was rejected by callee. (Exception from HRESULT: 0x80010001 (RPC_E_CALL_REJECTED))`。** 这是 COM 忙（WPS 正在重算或弹了对话框）。等 3 秒重试一次，通常就过；连续三次失败说明 WPS 里有个模态框等人点，让用户看一眼。

**坑 3：`New-Object -ComObject` 报 `0x80040154 Class not registered`。** 说明拿错了 progid 或跑在了没有 Office/WPS 注册的会话里。不要在这个方向上试第二个 progid，直接用桥已有的 action——worker 已经持有实例。

**坑 4：日期和时间不是文本。** 表格里的日期是 OADate 整数（`46082` = 2026-03-01），时间是 0–1 的小数（`0.8333` = 20:00）。写日期要写整数，写时间要写小数，写字符串进去 WPS 会按本地格式再解释一遍，结果偏移一天。

### 改表的三条硬规矩

1. **序号列整组重编号，不要局部改。** 插一行之后，同组所有序号都可能变，只改插入那一行会留下两行同号。
2. **不要相信"看起来一样"。** 空字符串、全角空格、`\ufeff` BOM 在 `getRange` 里长得几乎一样，写回时会把 BOM 带进去。补回原文前先看它的字节。
3. **同一份数据只维护一份解析结果。** 如果要同时写表格和别处（待办、文档），解析一次、两边共用，不要各自读一遍单子——两边不一致时你分不清是解析错了还是写错了。

### 桥的安全开关

- `-Token`：设了之后每个请求要带 `X-Token` 头
- `-AllowScript`：默认**关闭**。开着才能用 `POST /script` 跑任意 PowerShell。只在需要传文件、备份、读注册表这类动作时临时开，用完关掉
- 默认只监听内网；不要在公网暴露这个端口
