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

## 安装

**Windows 侧**

1. 建一个可写目录（别放 `C:\Program Files`，写不了日志），比如 `D:\wps`
2. 把 `wps-exec.ps1`、`wps-com.ps1`、`wps-worker.ps1` 放进去
3. 先打开一个 WPS 文档，让 WPS 处于运行状态
4. 右键 `启动WPS桥.bat` → 允许 UAC
5. 首次运行 Windows 防火墙弹窗 → 勾「专用网络」
6. 窗口里会打印 `http://192.168.x.x:8765/ping`，浏览器打开看到 `{"ok":true,...}` 就成了

**Linux / DSH 侧**

```bash
bash wps.sh set 172.16.1.85:8765     # 记下地址
bash wps.sh ping                     # 探活
bash wps.sh find 172.16.1            # 不知道 IP 时扫网段
bash wps.sh show                     # 看当前配置
```

地址存在 `~/.wps-exec.conf`，也可以用环境变量临时覆盖：`WPS_HOST=1.2.3.4:8765 bash wps.sh ping`。

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

## 灵感来源

这个桥不是原创的形态。**「用本地 HTTP 端口把『模型能发什么』和『Office 能做什么』解耦，两侧各写薄薄一层」这个思路，来自社区里已有人写过的 WPS 桥接插件**——本仓库是在那套思路上的重写与整理，不是从零发明。

具体差别在实现层：

- 服务端换成了常驻 worker（`wps-worker.ps1`），请求排队串行执行，避免 COM 重入
- `wps-com.ps1` 的 COM 封装按 WPS 官方 COM 接口重新整理，动作表与参数命名重写
- 加了 `-Token` 与 `-AllowScript` 两道开关（默认关闭任意脚本执行）
- Linux 侧客户端 `wps.sh` 带 `find` 自动发现与 `~/.wps-exec.conf` 持久化

原作者的仓库/昵称我这边没有准确记录，所以不在这里编一个。若你认出这套形态的出处，欢迎提 issue 补上署名。

`wps-com.ps1` 的 COM 封装基于 WPS 官方 COM 接口整理。

## License

MIT
