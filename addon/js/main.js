/**
 * DSH bridge addin - ribbon dispatcher only.
 *
 * Design rule: this file NEVER does real work inside the WPS process.
 * Every button spawns a detached PowerShell process that shows its own
 * MsgBox. If that process blocks or crashes, WPS is unaffected.
 *
 * 2026-11-11: 去掉了「启动服务 / 停止服务」两个按钮 —— WPS 的安全策略
 * 会拦 ShellExecute 起服务进程，那个开关从来没成功过，只会在点击时
 * 弹出 cmd.exe 的风险警告。启动桥改回双击「启动WPS桥.bat」。
 * 同时去掉了 cmd.exe /c start 那一层：它除了多弹一个黑框、多触发一次
 * 告警之外没别的作用，ShellExecute 本来就能起独立进程。
 *
 * All user-facing text lives in the .ps1 files (MsgBox), so this file
 * stays pure ASCII and cannot be mangled by the JS engine's codepage.
 */

var ADDIN_DIR = '%APPDATA%\\kingsoft\\wps\\jsaddons\\wps-claude-addon_';

function runScript(name) {
    var script = ADDIN_DIR + '\\' + name + '.ps1';
    var ps = 'powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + script + '"';

    try {
        if (wps.OAAssist && wps.OAAssist.ShellExecute) {
            wps.OAAssist.ShellExecute(ps);
            return;
        }
    } catch (e) {
        /* fall through */
    }

    try {
        wps.WpsApplication().Run(ps);
    } catch (e2) {
        /* nothing left to try */
    }
}

function OnAddinLoad(ribbonUI) {
    // 加载即执行是安全策略最敏感的行为 —— 这里永远不 spawn 进程。
    return true;
}

function OnShowCommand(control) { runScript('show'); }
function OnConfig(control) { runScript('config'); }
function OnPing(control) { runScript('ping'); }
function OnHelp(control) { runScript('help'); }

/* ribbon "getEnabled"/"getVisible" callbacks may be requested by the host */
function getEnabled(control) { return true; }
function getVisible(control) { return true; }
