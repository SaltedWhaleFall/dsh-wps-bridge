#!/usr/bin/env bash
# wps.sh —— 从 DSH 这边驱动 Windows 上的 wps-exec.ps1
#
#   地址保存在 ~/.wps-exec.conf（一行 WPS_HOST=IP:PORT），首次用 set 写。
#   bash wps.sh set 172.16.1.85:8765   改地址/端口
#   bash wps.sh show                    看当前地址
#   bash wps.sh find                    自动扫网段找 8765
#   bash wps.sh find 172.16.1           指定扫 172.16.1.0/24
#   bash wps.sh ping
#   bash wps.sh call getAppInfo
#   bash wps.sh call setCell '{"row":1,"col":1,"value":"hi"}'
#   bash wps.sh script 'Get-Process wps'
#   bash wps.sh ps '代码'        应急通道：一次性子进程（worker 挂了也能用）
#   bash wps.sh log             看 worker 的 stdout/stderr 尾巴
#   WPS_HOST=1.2.3.4:8765 bash wps.sh ping   临时覆盖（优先级最高）
set -uo pipefail

CONF="${WPS_CONF:-$HOME/.wps-exec.conf}"
PORT_DEFAULT=8765

conf_get() { [ -f "$CONF" ] && sed -n 's/^WPS_HOST=//p' "$CONF" | tail -1; }
conf_set() { printf 'WPS_HOST=%s\n' "$1" > "$CONF"; }

HOST="${WPS_HOST:-$(conf_get || true)}"

usage() { sed -n '2,16p' "$0" | sed 's/^# \{0,1\}//'; exit 2; }

pretty() {
  if command -v python3 >/dev/null 2>&1; then
    python3 -c 'import sys,json
raw=sys.stdin.read()
try:
    print(json.dumps(json.loads(raw),ensure_ascii=False,indent=2))
except Exception:
    print(raw)' 2>/dev/null || cat
  else
    cat
  fi
}

do_find() {
  local prefix="${1:-}"
  if [ -z "$prefix" ]; then
    prefix=$(ip -4 addr show 2>/dev/null | awk '/inet /{print $2}' | cut -d/ -f1 | grep -E '^(172\.16|192\.168|10\.)' | head -1 | cut -d. -f1-3)
  fi
  [ -z "$prefix" ] && { echo "拿不到本机网段，请手动指定：bash wps.sh find 172.16.1"; exit 2; }
  local port="${2:-$PORT_DEFAULT}"
  echo "扫描 ${prefix}.1-254:$port ..."
  seq 1 254 | xargs -P 4 -I{} bash -c '
    ip="'"$prefix"'.{}"
    if timeout 0.6 bash -c "echo > /dev/tcp/$ip/'"$port"'" 2>/dev/null; then echo "$ip:'"$port"'"; fi
  ' 2>/dev/null | sort -t. -k4 -n | while read -r hit; do
    local r
    r=$(curl -sS -m 3 "http://$hit/ping" 2>/dev/null || true)
    case "$r" in
      *'"service":"wps-exec"'*) echo "  ✔ $hit   $r" ;;
      *) echo "  ? $hit   (端口开着但不是 wps-exec)" ;;
    esac
  done
}

cmd="${1:-ping}"
case "$cmd" in
  set)
    v="${2:?需要 IP:PORT，例如 172.16.1.85:8765}"
    case "$v" in *:*) ;; *) v="$v:$PORT_DEFAULT" ;; esac
    conf_set "$v"; echo "已写入 $CONF → WPS_HOST=$v"
    ;;
  show)
    echo "配置文件 : $CONF"
    echo "当前地址 : ${HOST:-<未设置>}"
    echo "端口     : ${HOST##*:}"
    [ -n "$HOST" ] && curl -sS -m 5 "http://$HOST/ping" | pretty || true
    ;;
  find)
    do_find "${2:-}" "${3:-}"
    ;;
  ping|call|script|ps|log)
    [ -z "$HOST" ] && { echo "还没设地址。先 bash wps.sh set IP:PORT 或 bash wps.sh find"; exit 2; }
    case "$HOST" in http*) URL="$HOST" ;; *) URL="http://$HOST" ;; esac
    TOKEN="${WPS_TOKEN:-}"; AUTH=()
    [ -n "$TOKEN" ] && AUTH=(-H "X-Token: $TOKEN")
    case "$cmd" in
      ping) curl -sS -m 8 "${AUTH[@]}" "$URL/ping" | pretty ;;
      call)
        action="${2:?需要 action}"; params="${3:-{\}}"
        curl -sS -m 180 "${AUTH[@]}" -H 'Content-Type: application/json' \
          -d "{\"action\":\"$action\",\"params\":$params}" "$URL/wps" | pretty ;;
      script)
        code="${2:?需要 PowerShell 代码}"
        curl -sS -m 180 "${AUTH[@]}" -H 'Content-Type: application/json' \
          --data-binary "$(python3 -c 'import json,sys;print(json.dumps({"code":sys.argv[1]}))' "$code")" \
          "$URL/script" | pretty ;;
      ps)
        # 应急通道：一次性子进程，不依赖常驻 worker
        code="${2:?需要 PowerShell 代码}"
        curl -sS -m 180 "${AUTH[@]}" -H 'Content-Type: application/json' \
          --data-binary "$(python3 -c 'import json,sys;print(json.dumps({"code":sys.argv[1],"timeoutMs":int(sys.argv[2]) if len(sys.argv)>2 else 60000}))' "$code" "${3:-60000}")" \
          "$URL/ps" | pretty ;;
      log)
        curl -sS -m 10 "${AUTH[@]}" "$URL/workerlog" | pretty ;;
    esac
    ;;
  *) usage ;;
esac
