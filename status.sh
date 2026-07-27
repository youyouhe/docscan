#!/bin/bash
# ============================================================
# DocScan 状态报告（调试用）
# 用法: ./status.sh
#       sg docker -c ./status.sh     # 想看 docker 容器细节但还没重登录时
# 不需要 sudo；docker 信息在无权限时自动降级并提示。
# ============================================================
set -uo pipefail   # 故意不开 -e：调试脚本要尽量多输出，单点失败不中断

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

if [ -t 1 ]; then
    C_BLUE=$'\033[1;34m'; C_GREEN=$'\033[1;32m'; C_YELLOW=$'\033[1;33m'
    C_RED=$'\033[1;31m';  C_DIM=$'\033[2m';      C_CYAN=$'\033[1;36m'; C_RESET=$'\033[0m'
else
    C_BLUE=''; C_GREEN=''; C_YELLOW=''; C_RED=''; C_DIM=''; C_CYAN=''; C_RESET=''
fi
step(){ printf '\n%s══ %s %s\n' "$C_BLUE"  "$*" "$C_RESET"; }
ok(){   printf '  %s✅ %s%s\n'  "$C_GREEN"  "$*" "$C_RESET"; }
warn(){ printf '  %s⚠️  %s%s\n' "$C_YELLOW" "$*" "$C_RESET"; }
err(){  printf '  %s❌ %s%s\n'  "$C_RED"    "$*" "$C_RESET"; }
info(){ printf '  %s%s%s\n'    "$C_CYAN"   "$*" "$C_RESET"; }
dim(){  printf '  %s%s%s\n'    "$C_DIM"    "$*" "$C_RESET"; }

docker_ok(){ docker info >/dev/null 2>&1; }

printf '%s╔════════════════════════════════════╗%s\n' "$C_BLUE" "$C_RESET"
printf '%s║      DocScan 状态报告（调试）      ║%s\n' "$C_BLUE" "$C_RESET"
printf '%s╚════════════════════════════════════╝%s\n' "$C_BLUE" "$C_RESET"

# ---------- 1. 进程 ----------
step "服务进程"
PIDS=(); PORTS=()
for pf in "$SCRIPT_DIR"/.docscan-*.pid; do
    [ -f "$pf" ] || continue
    p=$(cat "$pf" 2>/dev/null || true)
    if [ -n "$p" ] && kill -0 "$p" 2>/dev/null; then
        pt=$(basename "$pf" | sed -E 's/^\.docscan-([0-9]+)\.pid$/\1/')
        PIDS+=("$p"); PORTS+=("${pt:-8800}")
    fi
done
if [ "${#PIDS[@]}" -eq 0 ]; then
    for p in $(pgrep -f "api.py" 2>/dev/null || true); do
        PIDS+=("$p"); PORTS+=("8800")
    done
fi
if [ "${#PIDS[@]}" -eq 0 ]; then
    err "DocScan API 未运行（用 ./start.sh 启动）"
else
    for i in "${!PIDS[@]}"; do
        pid="${PIDS[$i]}"; port="${PORTS[$i]}"
        elapsed=$(ps -o etime= -p "$pid" 2>/dev/null | tr -d ' ' || echo '?')
        ok "API 运行中  PID=$pid  端口=$port  已运行 $elapsed"
        dim "命令: $(ps -o args= -p "$pid" 2>/dev/null)"
    done
fi

PID="${PIDS[0]:-}"; PORT="${PORTS[0]:-8800}"; BASE="http://localhost:$PORT"

# ---------- 2. 健康检查 / 认证 ----------
step "健康检查 / 认证（端口 $PORT）"
if h=$(curl -s -m 3 "$BASE/api/health" 2>/dev/null); then
    ok "GET /api/health → $h"
else
    err "无法连接 $BASE/api/health"
fi
KEY=""
[ -f "$SCRIPT_DIR/.docscan-api-key" ] && KEY=$(cat "$SCRIPT_DIR/.docscan-api-key" 2>/dev/null || true)
if [ -n "$KEY" ]; then
    dim "API Key (.docscan-api-key): $KEY"
    code=$(curl -s -o /dev/null -m 5 -w '%{http_code}' -H "X-API-Key: $KEY" "$BASE/api/conversions" 2>/dev/null || echo 000)
    if [ "$code" = "200" ]; then ok "带 key 访问 /api/conversions → 200（key 与进程一致）"
    else warn "带 key 访问 /api/conversions → $code（key 可能与运行中进程不一致）"; fi
else
    warn "无 .docscan-api-key（若用 DOCSCAN_API_KEY 环境变量启动，以进程环境为准）"
fi

# ---------- API 端点（动态取自 /openapi.json，需 key） ----------
step "API 端点"
AUTH_HDR=()
[ -n "$KEY" ] && AUTH_HDR=(-H "X-API-Key: $KEY")
if spec=$(curl -s -m 3 "${AUTH_HDR[@]}" "$BASE/openapi.json" 2>/dev/null) && [ -n "$spec" ]; then
    dim "base: $BASE   受保护端点需带 -H \"X-API-Key: \$KEY\""
    printf '%s' "$spec" | python3 -c 'import sys, json
base = sys.argv[1]
d = json.load(sys.stdin)
public = {"/api/health", "/api/docs"}
rows = []
for p in sorted(d.get("paths", {})):
    ms = [m.upper() for m in d["paths"][p] if m.lower() in ("get", "post", "put", "delete", "patch")]
    if not ms:
        continue
    tag = "  (免key)" if p in public else ("  (页面/代理)" if not p.startswith("/api/") else "")
    rows.append((",".join(ms), p, tag))
w = max((len(m) for m, _, _ in rows), default=0)
for m, p, t in rows:
    print(f"  {m:<{w}}  {base}{p}{t}")' "$BASE"
else
    warn "无法获取 $BASE/openapi.json（服务未起或端口不对）"
fi

# ---------- 3. 运行中 API 的 DOCSCAN_* 配置 ----------
step "运行中 API 的 DOCSCAN_* 配置"
# 配置项及代码默认值（与 api.py 保持一致）；ORDER 决定显示顺序
declare -A DEFAULTS=(
    [DOCSCAN_API_KEY]="<自动生成>"
    [DOCSCAN_CORS_ORIGINS]="*"
    [DOCSCAN_MAX_UPLOAD_MB]="100"
    [DOCSCAN_RETENTION_HOURS]="0 (关)"
    [DOCSCAN_CONVERT_CONCURRENCY]="4"
)
ORDER=(DOCSCAN_API_KEY DOCSCAN_CORS_ORIGINS DOCSCAN_MAX_UPLOAD_MB DOCSCAN_RETENTION_HOURS DOCSCAN_CONVERT_CONCURRENCY)

# 读取实际生效值：优先 runtime.env，降级 /proc/environ
declare -A ACTUAL=()
RUNTIME_ENV="$SCRIPT_DIR/.docscan-$PORT.runtime.env"
SRC=""
if [ -z "$PID" ]; then
    warn "无运行中的 API 进程（下方为代码默认值）"
elif [ -f "$RUNTIME_ENV" ]; then
    SRC="$RUNTIME_ENV（start.sh 启动时落盘）"
    while IFS='=' read -r k v; do
        case "$k" in DOCSCAN_*) ACTUAL["$k"]="$v";; esac
    done < "$RUNTIME_ENV"
else
    env_raw=$(cat "/proc/$PID/environ" 2>/dev/null || true)
    if [ -n "$env_raw" ]; then
        SRC="/proc/$PID/environ"
        while IFS= read -r line; do
            k=${line%%=*}; v=${line#*=}
            case "$k" in DOCSCAN_*) ACTUAL["$k"]="$v";; esac
        done <<< "$(printf '%s' "$env_raw" | tr '\0' '\n')"
    else
        warn "无 runtime.env 且 /proc 不可读（ptrace 限制），下方为默认值"
    fi
fi
[ -n "$SRC" ] && dim "来源: $SRC"

# 统一表：每项 = 实际值，标注「已设置(覆盖默认)」或「默认」
for k in "${ORDER[@]}"; do
    if [ -n "${ACTUAL[$k]+x}" ]; then
        printf '  %s%-28s = %s%s  %s(已设置)%s\n' "$C_CYAN" "$k" "${ACTUAL[$k]}" "$C_RESET" "$C_YELLOW" "$C_RESET"
    else
        printf '  %s%-28s = %s  (默认)%s\n' "$C_DIM" "$k" "${DEFAULTS[$k]}" "$C_RESET"
    fi
done

# ---------- 4. ONLYOFFICE ----------
step "ONLYOFFICE"
if curl -s -m 3 -o /dev/null -w '%{http_code}' http://localhost:8079/healthcheck 2>/dev/null | grep -q 200; then
    ok "ONLYOFFICE :8079 /healthcheck → 200"
else
    err "ONLYOFFICE :8079 健康检查失败（容器未起？docker compose up -d）"
fi
if docker_ok; then
    st=$(docker ps -a --filter name=onlyoffice --format '{{.Status}} ({{.Image}})' 2>/dev/null | head -1 || true)
    [ -n "$st" ] && info "容器: $st" || warn "未找到 onlyoffice 容器"
else
    warn "无 docker 权限 → 跳过容器详情（用 sg docker -c ./status.sh 查看）"
fi

# ---------- 5. Docker / 权限 ----------
step "Docker / 权限"
if id -nG 2>/dev/null | grep -qw docker; then ok "当前用户在 docker 组"
else warn "当前用户不在 docker 组（重登录或用 sg docker -c ...）"; fi
if docker_ok; then
    info "docker: $(docker --version 2>&1)"
    info "compose: v$(docker compose version --short 2>/dev/null || echo '?')"
else
    warn "docker 命令不可用（无权限或 daemon 未运行）"
fi

# ---------- 6. 数据目录 ----------
step "数据目录"
for d in docs pdfs mds docx_store; do
    if [ -d "$SCRIPT_DIR/$d" ]; then
        n=$(find "$SCRIPT_DIR/$d" -maxdepth 1 -type f 2>/dev/null | wc -l | tr -d ' ')
        sz=$(du -sh "$SCRIPT_DIR/$d" 2>/dev/null | cut -f1)
        info "$d/: ${n} 个文件, ${sz}"
    else
        dim "$d/: 不存在"
    fi
done

# ---------- 7. 依赖版本 ----------
step "依赖版本"
info "python3: $(python3 --version 2>&1)"
if command -v pandoc >/dev/null 2>&1; then info "pandoc: $(pandoc --version 2>&1 | head -1)"; else warn "pandoc 未安装"; fi
if command -v curl    >/dev/null 2>&1; then info "curl: $(curl --version 2>&1 | head -1 | cut -d' ' -f1-2)"; else warn "curl 未安装"; fi

# ---------- 8. 代码版本 ----------
step "代码版本"
if command -v git >/dev/null 2>&1 && [ -d "$SCRIPT_DIR/.git" ]; then
    info "commit: $(git -C "$SCRIPT_DIR" rev-parse --short HEAD 2>/dev/null) - $(git -C "$SCRIPT_DIR" log -1 --format='%s' 2>/dev/null)"
    if git -C "$SCRIPT_DIR" diff --quiet 2>/dev/null; then ok "工作区干净"
    else warn "工作区有未提交改动"; fi
else
    dim "非 git 仓库"
fi

# ---------- 9. 日志尾部 ----------
step "最近日志（docscan-$PORT.log 末尾 15 行）"
LOG="$SCRIPT_DIR/docscan-$PORT.log"
if [ -f "$LOG" ]; then
    tail -n 15 "$LOG" 2>/dev/null \
      | sed -E "s/(ERROR|Error|Traceback|Exception|FAILED)/${C_RED}\1${C_RESET}/g" \
      | sed 's/^/  /'
else
    dim "无日志文件 docscan-$PORT.log"
fi

echo ""
dim "提示: ./start.sh 启动 · ./stop.sh 停止 · ./restart.sh 重启 · ./setup.sh 重装 · ./status.sh 本报告"
