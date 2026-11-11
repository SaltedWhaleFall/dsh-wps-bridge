# dsh-wps-bridge

让 DSH（跑在 Linux / 手机 proot 里）直接操作另一台 Windows 上的 WPS：读写单元格、改文档、导出 PDF、跑 PowerShell。

不是一个 MCP server，而是一条**双端桥**：Windows 端点起一个 HTTP 服务，Linux 端用 shell 调它。

```
DSH (Linux / proot)
    │  bash wps.sh call setCell '{"row":1,"col":1,"value":"hi"}'
    ▼  HTTP + JSON
Windows: wps-exec.ps1  ────►  wps-com.ps1  ────►  WPS COM (Excel / Word / PPT)
         (监听 8765)           (COM 封装)
```

## 为什么不是 MCP

WPS 的 COM 接口只在 Windows 上存在，而 DSH 常跑在手机 proot 里。套一层 MCP 需要 Node 运行时 + stdio 管道，跨机器时反而更麻烦。这里用最笨的办法：**一个 HTTP 端口 + 一个 curl 客户端**，任何能发 HTTP 的地方都能驱动它。

## 文件

| 文件 | 跑在哪 | 作用 |
|---|---|---|
| `wps-exec.ps1` | Windows | HTTP 服务端，收 JSON 转调 COM；需管理员权限（监听局域网） |
| `wps-com.ps1` | Windows | COM 封装，Excel/Word/PPT 具体动作都在这里 |
| `wps-worker.ps1` | Windows | 常驻 worker 模式（`wps-exec.ps1` 的子进程方案，避免每次冷启动 COM） |
| `wps.sh` | Linux / DSH | 客户端：`ping` / `call` / `script` / `find` / `log` |
| `启动WPS桥.bat` | Windows | 双击启动，带 UAC 提升 |
| `addon/` | Windows（可选） | WPS 功能区加载项：一眼看端口、探活、抄手机命令；不装不影响桥 |

## 安装

### 需要装什么、跑什么（清单）

| 在哪 | 要装/要跑 | 必需 | 说明 |
|---|---|---|---|
| Windows | WPS Office（或 Excel / Word / PowerPoint） | 是 | 桥操作的就是它，必须已经有一个文档开着 |
| Windows | PowerShell 5.1（系统自带） | 是 | 端口、COM 都靠它 |
| Windows | `wps-exec.ps1` `wps-com.ps1` `wps-worker.ps1` `启动WPS桥.bat` | 是 | 放在同一个目录里 |
| Windows | **`addon/` 目录 → WPS 菜单栏选项卡**（可选） | 否 | 装了能在 WPS 里直接点按钮探活/看手机命令；不装也能用 |
| Windows | Node.js ≥ 22 + DSH | 否 | 只有把桥当 DSH 插件（`index.js`）装时才需要 |
| Linux / DSH | `wps.sh` | 是 | 客户端，`ping` / `call` / `script` / `find` |

### Windows 侧：起桥

1. 建一个可写目录（别放 `C:\Program Files`，写不了日志），比如 `D:\wps`
2. 把 `wps-exec.ps1`、`wps-com.ps1`、`wps-worker.ps1`、`启动WPS桥.bat` 放进去
3. 先打开一个 WPS 文档，让 WPS 处于运行状态（桥要接管这个正在跑的实例，冷启动的隐藏实例你保存了也看不见）
4. 双击 `启动WPS桥.bat` → 允许 UAC（监听局域网需要管理员）
5. 首次运行 Windows 防火墙弹窗 → 勾「专用网络」（**只勾专用**，别勾公用）
6. 窗口里会打印 `http://192.168.x.x:8765/ping`，浏览器打开看到 `{"ok":true,...}` 就成了

参数都在 `启动WPS桥.bat` 里，可以改：

```
-Port 8765          监听端口
-Token <字符串>     启用后每个请求要带 X-Token 头
-AllowScript        开启 POST /script（能跑任意 PowerShell，默认关）
-NoElevate          不申请管理员（只监听 127.0.0.1 时可省掉 UAC）
```

### Windows 侧：装 WPS 菜单栏加载项（可选）

不装也能用——桥是独立进程，手机端照样能调。装了的好处是：在 WPS 里点一下就能看当前端口、探活、抄走手机要用的命令。

1. 把仓库里的 **`addon/` 整个目录**拷到：

   ```
   %APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon_\
   ```

   也就是 `C:\Users\<你>\AppData\Roaming\kingsoft\wps\jsaddons\wps-claude-addon_\`（**带结尾下划线**，WPS 要求目录名与 `manifest.xml` 里的插件名对应）

2. 目录里应该有：`manifest.xml`、`ribbon.xml`、`js/main.js`、`index.html`、`config.ps1`、`ping.ps1`、`probe.ps1`、`show.ps1`、`help.ps1`
3. **完全退出 WPS，再打开**（不是关文档，是退出进程；任务栏右键 → 退出）
4. 打开任意表格 → 功能区最右边出现 **「DSH 桥」** 选项卡
5. 不想要了：退出 WPS，把 `wps-claude-addon_` 这个目录删掉即可，不改注册表

**这一版加载项里没有「启动/停止服务」按钮**——WPS 的 JS 加载项沙箱不让它起外部进程（点了会触发 `cmd.exe` 安全告警，还起不来）。启停一律走 `启动WPS桥.bat`。

### Linux / DSH 侧：配地址

```bash
bash wps.sh set 192.168.1.50:8765      # 记下地址（换成你那台的 IP:端口）
bash wps.sh ping                     # 探活，期望 {"ok":true,...}
bash wps.sh find 192.168.1             # 不知道 IP 时扫网段（DHCP/校园网会跳）
bash wps.sh show                     # 看当前配置
```

## 配套 Skill（可选）

仓库里带了一份给 agent 用的技能：`skills/wps-doc-ops/SKILL.md`。装了它，agent 就知道这 48 个 action 怎么用、哪些坑会让数据静默坏掉、写表前要先备份。

- 本地装：拷到 `~/.dsh/skills/wps-doc-ops/SKILL.md`（或 `.dsh/skills/` 对应目录）
- 市场装：https://deepseek.stream/skills/wps-doc-ops

## 接口

```
GET  /ping                              探活
GET  /                                  浏览器说明页
POST /wps    {"action":..,"params":{}}  转调 COM
POST /script {"code":"..."}             直接执行 PowerShell（必须 -AllowScript 启动）
```

客户端封装：

```bash
bash wps.sh call getDocumentStats '{"sheet":"data"}'
bash wps.sh call setRange '{"sheet":"data","range":"A1:C2","data":[["a","b","c"]]}'
bash wps.sh call save
bash wps.sh script 'Get-Process wps'
bash wps.sh ps 'Get-Date'        # 应急通道：一次性子进程，worker 挂了也能用
bash wps.sh log                  # worker 的 stdout/stderr 尾巴
```

## 安全

默认**不开** `/script`（能执行任意 PowerShell）。要用必须显式加 `-AllowScript`：

```powershell
powershell -ExecutionPolicy Bypass -File wps-exec.ps1 -Token <随机串> -AllowScript
```

- `-Token` 设了之后，所有请求必须带 `X-WPS-Token` 头
- 服务默认只监听局域网，别把 8765 映射到公网
- `/script` 等价于远程 shell，只在完全信任的网段用

## 踩过的坑

- **中文写入静默丢失**：`setCell` 传中文字符串会返回 `{"ok":true,"wrote":"曾燚"}`，但读回还是旧值。改用 `setRange` + `--data-binary @file`（UTF-8 JSON）才写得进去。
- **`save` 偶发 `RPC_E_CALL_REJECTED`**（`0x80010001`）：WPS 正在忙，等 3 秒重试即可，文件不会坏。
- **COM 类未注册**：在 `wps.sh script` 里 `New-Object -ComObject ket.Application` 会报 `0x80040154`。COM 对象只存在于 worker 进程里，不要另起实例。
- **编辑长文档后必须先 `save` 再读回**，否则读到的是内存里未提交的版本。

## 作为一个 DSH 插件装

仓库根目录就是一个 DSH 插件（`index.js` + `cordis.patch.yml`），脚本作为附带文件一起发。

```bash
# ① 挂载（四步一起做，幂等）
node <你的插件工作区>/link.mjs dsh-wps-bridge
# ② 重启 dsh
```

或者手改 `$DSH_HOME/profiles/web/package.json`：

```json
"dependencies": { "dsh-wps-bridge": "link:<插件目录绝对路径>" },
"dsh": { "profile": { "bundles": ["dsh-wps-bridge"] } }
```

挂上之后多一个工具 `wps`：

| action | 干什么 |
|---|---|
| `ping` | 探活，看那台 Windows 的桥在不在 |
| `call` | 原样转发一个动作：`op` + `params`（`open` / `getRange` / `setRange` / `appInfo` / `exportPdf` …） |
| `read` | `getRange` 简写，给 `range` |
| `write` | `setRange` 简写，给 `range` + `rows` 二维数组 |
| `save` | 保存当前文档 |
| `script` | 跑一段 PowerShell（要 Windows 侧开 `-AllowScript`） |

地址两种给法，环境变量优先：

```bash
export WPS_BRIDGE_URL=http://192.168.1.50:8765
export WPS_BRIDGE_TOKEN=...        # 与 Windows 侧 -Token 一致，没开就留空
```

或插件配置 `{ "url": "...", "token": "..." }`。

> 插件不启动任何服务、不监听端口 —— 它只是客户端。真正干活的是 Windows 那台。

## 灵感来源

这个桥的形态不是原创的。

**「把 WPS 包成一个本地桥，让模型用一句话读写表格」这件事，`lc2panda/wps-skills` 已经做过** —— 那套仓库里是一个 TypeScript 写的 MCP server（`wps-office-mcp`）加一组 Skills，走「把 WPS 暴露成 MCP 工具」的路子。本项目借用的是这个**形态判断**，不是它的代码：它是 Node/TypeScript 的 MCP server，这里是 Windows 侧常驻 worker 直连 COM 的另一条实现，两边没有共享源码。

差别在实现层：

- 服务端是常驻 PowerShell worker（`wps-worker.ps1`），请求排队串行执行，避免 COM 重入
- 不走 MCP 协议，直接 HTTP JSON：`POST /wps {"action":..,"params":{}}`，Linux 侧 `curl` 一行就能调
- 加了 `-Token` 与 `-AllowScript` 两道开关（后者默认关闭任意脚本执行）
- `wps.sh` 带 `find` 自动发现与 `~/.wps-exec.conf` 持久化

致谢：[lc2panda/wps-skills](https://github.com/lc2panda/wps-skills)（MIT）。

`wps-com.ps1` 的 COM 封装按 WPS 官方 COM 接口整理，动作命名尽量贴近常见叫法。

## License

MIT
