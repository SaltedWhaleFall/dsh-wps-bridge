WPS 执行端 · 使用说明
=====================

这个文件夹里有两个文件，必须放在同一个文件夹里：
  wps-exec.ps1   ← 你要运行的
  wps-com.ps1    ← 官方 COM 桥，wps-exec 会调用它

一、放到电脑上
  在 D 盘建一个文件夹，比如 D:\wps，把这两个文件都放进去。
  （不要放在 C:\Program Files 下面，那里写不了日志）

二、先打开 WPS
  随便打开一个 Excel 或 Word 文档，让 WPS 处于运行状态。
  没打开的话，脚本会自己新建一个隐藏实例，你看不到窗口，容易以为没反应。

三、运行
  右键 wps-exec.ps1 →「使用 PowerShell 运行」
  会弹 UAC，点「是」。（监听局域网端口必须要管理员权限）
  如果右键没有这一项：开始菜单搜 PowerShell → 右键「以管理员身份运行」→ 输入：
      powershell -ExecutionPolicy Bypass -File "D:\wps\wps-exec.ps1"

四、看窗口里打印的地址
  会出现类似：
      http://192.168.1.23:8765/ping
      http://127.0.0.1:8765/ping   （本机）
  第一行那个就是手机要访问的地址。

五、第一次运行会弹 Windows 防火墙
  勾选「专用网络」，点「允许访问」。
  如果没弹或者点了取消，手机就连不上——去
  「Windows 安全中心 → 防火墙和网络保护 → 允许应用通过防火墙」
  里找 PowerShell，把「专用」勾上。

六、验证
  手机浏览器打开  http://你电脑的IP:8765/ping
  看到 {"ok":true,...} 就成了。

常见问题
--------
1. 手机打不开，电脑本机 127.0.0.1 能开
   → 防火墙没放行，或者手机和电脑不在同一个 WiFi。
     校园网/公司网常常开了「客户端隔离」，同一 WiFi 也互相看不见。
     这种情况改用手机热点，让电脑连手机热点。

2. 提示端口被占用
   → 换端口：powershell -ExecutionPolicy Bypass -File wps-exec.ps1 -Port 8766

3. 提示「没找到 wps-com.ps1」
   → 两个文件没放在一起，或者用 -ComPath "完整路径" 指定。

4. 想让我能直接执行任意 PowerShell（排障用，风险自负）
   → 加 -AllowScript 启动。默认关闭。

安全提示
--------
默认没有口令，同一局域网里任何人都能调用。家里 WiFi 无所谓，
公共 WiFi 下建议加 -Token 你的口令，我这边会带 X-Token 头。

停止：在那个窗口按 Ctrl+C，或直接关窗口。

七、自检（装完先跑这个）
  curl "http://127.0.0.1:8765/call?action=selfTest"
  或者手机浏览器打开  http://你电脑的IP:8765/ping  先确认能通，
  再让我跑 selfTest。全绿是 pass 31 / fail 0。
  第一次跑可能报 cannot create workbook —— 那是 WPS 刚起来还没注册完，
  再跑一次就好（现在这一版已经修掉，正常一次就过）。

八、worker 是常驻的，别反复重启
  wps-worker.ps1 是常驻 COM worker，wps-exec 只在它死了才拉新的。
  反复杀它会让 WPS 里堆一堆没关的工作簿；新版本已经会自动复用
  已经在跑的实例、跑完自己把簿关掉，但没事还是别手动杀。


九、为什么 WPS 每次都弹"检测到风险加载项"（2026-11-04 21:00 处理）

三个叠加的原因，按嫌疑从大到小：

1. **加载即执行**。`addon/js/main.js` 的 `OnAddinLoad` 原来会立刻 `ShellExecute`
   一个 PowerShell（跑 probe.ps1）。加载项一装上、WPS 一启动就起子进程 ——
   这是安全机制最敏感的行为。**已停用**，现在只在手动点「探针」按钮时才跑。
2. **jsaddons 目录里堆了旧备份**。`install-addon.ps1` 以前把旧版备份成
   `wps-claude-addon_.bak-<时间戳>` 留在同一目录里，WPS 每次启动扫到这个目录会看到
   三四个同名插件。**已把备份挪到** `C:\Users\admin\wps-skills\jsaddons-backup\`，
   installer 也已改成往那里备份。
3. `publish.xml` 里用的是 `enable="enable_dev"`（开发者模式）。这个先留着 ——
   前两条清掉后如果还弹，再考虑改成正式启用或换成 `authaddin.json` 里已有的授权条目。

验证方法：**完全退出 WPS**（任务管理器里 wps.exe / et.exe / wpp.exe 全结束，不是关窗口），
再重开，看还弹不弹。弹的话先看 `%APPDATA%\kingsoft\wps\jsaddons\` 是不是又多了东西。
