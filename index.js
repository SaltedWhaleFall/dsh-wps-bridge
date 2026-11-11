/**
 * dsh-wps-bridge — 宿主侧。
 *
 * 它只干一件事：把「这台机器上的 agent」和「那台 Windows 上的 WPS」之间架一条最短的线。
 *
 *   Windows 侧：wps-exec.ps1 起一个 HTTP 端口，wps-worker.ps1 常驻排队跑 COM
 *   这里：一个工具 `wps`，把 action + params 原样转过去，把 JSON 结果原样拿回来
 *
 * 不经过 MCP，不做协议转换 —— 桥两边都是 JSON in / JSON out，多一层翻译就多一处失败。
 *
 * ── 配置 ──────────────────────────────────────────────────────────────
 *   环境变量   WPS_BRIDGE_URL    例如 http://192.168.1.50:8765
 *              WPS_BRIDGE_TOKEN  与 Windows 侧 -Token 一致（没开就不用)
 *   插件配置   { url, token }
 */

export const name = 'dsh-wps-bridge'

/** defineTool 拿不到时退化成恒等（测试/极简宿主里发生）。 */
let defineTool = (d) => d
try {
  const mod = await import('@deepseek-ai/dsh-tools')
  if (mod && typeof mod.defineTool === 'function') defineTool = mod.defineTool
} catch { /* 宿主以外跑 —— 恒等即可 */ }

let cfg = {}
let log = () => {}

function baseUrl() {
  const u = String(cfg.url || process.env.WPS_BRIDGE_URL || '').trim()
  return u.replace(/\/+$/, '')
}

function authHeaders() {
  const t = String(cfg.token || process.env.WPS_BRIDGE_TOKEN || '').trim()
  return t ? { token: t, 'x-wps-token': t } : {}
}

/** 统一出口：超时、错误、非 JSON 全折成 {ok:false, message}。 */
async function call(path, body, timeoutMs = 120000) {
  const base = baseUrl()
  if (!base) {
    return {
      ok: false,
      message: '还没配地址。设 WPS_BRIDGE_URL=http://<那台机器的IP>:8765，或者在插件配置里写 { "url": "..." }。',
    }
  }
  const ctl = new AbortController()
  const timer = setTimeout(() => ctl.abort(), timeoutMs)
  try {
    const res = await fetch(base + path, {
      method: body === undefined ? 'GET' : 'POST',
      headers: Object.assign({ 'Content-Type': 'application/json' }, authHeaders()),
      body: body === undefined ? undefined : JSON.stringify(body),
      signal: ctl.signal,
    })
    const text = await res.text()
    let data
    try { data = JSON.parse(text) } catch { data = { ok: res.ok, raw: text.slice(0, 2000) } }
    if (!res.ok && data && data.ok === undefined) data.ok = false
    return data
  } catch (e) {
    const msg = e && e.name === 'AbortError'
      ? `等了 ${Math.round(timeoutMs / 1000)} 秒没回。那台机器可能没开、端口不通、或者 WPS 卡在弹窗上。`
      : '连不上：' + ((e && e.message) || e)
    return { ok: false, message: msg }
  } finally {
    clearTimeout(timer)
  }
}

/** 优先用桥自己给的 message，没有再退化成 JSON。 */
function fmt(r, n = 600) {
  if (r && typeof r.message === 'string' && r.message) return r.message
  return short(r, n)
}

function short(v, n = 600) {
  const s = typeof v === 'string' ? v : JSON.stringify(v)
  if (!s) return ''
  return s.length > n ? s.slice(0, n) + ` …（还有 ${s.length - n} 字）` : s
}

export function apply(rootCtx, rawConfig) {
  try {
    if (rawConfig && typeof rawConfig === 'object') cfg = Object.assign({}, cfg, rawConfig)
    if (rootCtx && typeof rootCtx.logger === 'function') log = rootCtx.logger('dsh-wps-bridge')
  } catch { /* 配置坏了也照挂，工具里会提示 */ }

  const mount = (ctx) => {
    try {
      const tools = ctx && ctx.tools
      if (!tools || typeof tools.register !== 'function') {
        log('宿主没给 tools 服务，跳过工具注册')
        return
      }

      tools.register(defineTool({
        name: 'wps',
        description:
          '操作另一台 Windows 上的 WPS（Excel/Word/PPT）。'
          + 'ping 看桥活没活；call 调一个动作（open/used/getRange/setRange/setCell/save/appInfo/exportPdf 等）；'
          + 'script 跑一段 PowerShell（需要 Windows 侧开了 -AllowScript）；'
          + 'read 是 getRange 的简写；write 是 setRange 的简写（中文务必走这个，不要用 setCell）。',
        parameters: {
          action: {
            type: 'string',
            enum: ['ping', 'call', 'script', 'read', 'write', 'save'],
            description: 'ping=探活；call=原样转发一个动作；script=跑 PowerShell；read/write=范围读写简写；save=保存当前文档',
          },
          op: { type: 'string', description: 'call 时的 WPS 动作名，例如 open / getRange / setRange / appInfo / exportPdf' },
          params: { type: 'object', description: 'call 时的参数对象，原样传给 WPS 桥' },
          range: { type: 'string', description: 'read/write 的 A1 范围，例如 data!A1:K539' },
          rows: { type: 'array', description: 'write 的二维数组（每行一个数组）' },
          code: { type: 'string', description: 'script 时的 PowerShell 代码' },
          timeoutMs: { type: 'number', description: '超时毫秒，默认 120000' },
        },
        output: {
          schema: { type: 'json' },
          render: (_args, value) => [{ type: 'text', text: value.message }],
        },
        async execute(args) {
          const a = args || {}
          const act = String(a.action || 'ping')
          const t = Number(a.timeoutMs) > 0 ? Number(a.timeoutMs) : 120000

          if (act === 'ping') {
            const r = await call('/ping', undefined, Math.min(t, 15000))
            return { ok: !!(r && r.ok !== false), message: '桥：' + fmt(r, 400) }
          }

          if (act === 'script') {
            const code = String(a.code || '').trim()
            if (!code) return { ok: false, message: 'script 要给 code。' }
            const r = await call('/script', { code }, t)
            return { ok: !!(r && r.ok !== false), message: fmt(r, 4000) }
          }

          if (act === 'read') {
            const range = String(a.range || '').trim()
            if (!range) return { ok: false, message: 'read 要给 range，例如 data!A1:K20。' }
            const r = await call('/wps', { action: 'getRange', params: { range } }, t)
            return { ok: !!(r && r.ok !== false), message: fmt(r, 6000) }
          }

          if (act === 'write') {
            const range = String(a.range || '').trim()
            const rows = Array.isArray(a.rows) ? a.rows : null
            if (!range || !rows || !rows.length) return { ok: false, message: 'write 要给 range 和 rows（二维数组）。' }
            const r = await call('/wps', { action: 'setRange', params: { range, data: rows } }, t)
            return { ok: !!(r && r.ok !== false), message: fmt(r, 2000) }
          }

          if (act === 'save') {
            const r = await call('/wps', { action: 'save', params: a.params || {} }, t)
            return { ok: !!(r && r.ok !== false), message: fmt(r, 1000) }
          }

          // call
          const op = String(a.op || '').trim()
          if (!op) return { ok: false, message: 'call 要给 op（动作名）。' }
          const r = await call('/wps', { action: op, params: a.params || {} }, t)
          return { ok: !!(r && r.ok !== false), message: fmt(r, 6000) }
        },
      }))

      log('工具 wps 已挂载；地址 = ' + (baseUrl() || '(未配置)'))
    } catch (e) {
      log('挂载失败：' + ((e && e.message) || e))
    }
  }

  if (rootCtx && typeof rootCtx.inject === 'function') {
    rootCtx.inject(['tools'], mount)
  } else {
    mount(rootCtx)
  }
}
