#!/bin/bash
# ============================================================
# DocScan API — 启动脚本
# 用法: ./start.sh [port]
# 默认端口: 8800
# ============================================================
set -e

PORT="${1:-8800}"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PID_FILE="$SCRIPT_DIR/.docscan-$PORT.pid"
LOG_FILE="$SCRIPT_DIR/docscan-$PORT.log"

cd "$SCRIPT_DIR"

# ---------- 0. 检查端口占用 ----------
if ss -tln 2>/dev/null | grep -q ":$PORT\b"; then
    echo "❌ 端口 $PORT 已被占用"
    echo "   请使用其他端口: ./start.sh 8800"
    exit 1
fi

# ---------- 1. 确保 ONLYOFFICE 容器就绪 ----------
echo "📦 检查 ONLYOFFICE 容器…"
if ! docker ps --format '{{.Names}}' | grep -q '^onlyoffice$'; then
    if docker ps -a --format '{{.Names}}' | grep -q '^onlyoffice$'; then
        echo "   容器已存在但未运行，启动中…"
        docker start onlyoffice
    else
        echo "   首次使用，拉取 ONLYOFFICE 镜像 (约 3GB，需要几分钟)…"
        echo "   （网络不稳会自动重试，支持断点续传；持续失败见末尾加速器提示）"
        for i in $(seq 1 5); do
            if docker pull onlyoffice/documentserver:latest; then
                break
            fi
            if [ "$i" -eq 5 ]; then
                echo "   ❌ 镜像拉取连续 5 次失败。"
                echo "      多半是访问 Docker Hub 网络不稳（connection reset）。可配置镜像加速器后重试："
                echo "        sudo tee /etc/docker/daemon.json >/dev/null <<'EOF'"
                echo '        { "registry-mirrors": ["https://<你的加速器地址，如阿里云专属地址>.mirror.aliyuncs.com"] }'
                echo "        EOF"
                echo "        sudo systemctl restart docker && ./start.sh"
                exit 1
            fi
            echo "   ⚠️  拉取中断（connection reset），10 秒后重试 ($i/5)…"
            sleep 10
        done
        docker compose up -d
    fi
    sleep 10
fi

for i in $(seq 1 30); do
    if curl -s -o /dev/null http://localhost:8079/healthcheck 2>/dev/null; then
        echo "   ONLYOFFICE 就绪 ✅"
        break
    fi
    [ "$i" -eq 30 ] && { echo "   ❌ ONLYOFFICE 启动超时"; exit 1; }
    sleep 2
done

# ---------- 2. 禁用 JWT ----------
echo "🔐 禁用 ONLYOFFICE JWT 认证…"
docker exec onlyoffice python3 -c "
import json
with open('/etc/onlyoffice/documentserver/local.json') as f:
    cfg = json.load(f)
cfg['services']['CoAuthoring']['token']['enable']['request']['inbox'] = False
cfg['services']['CoAuthoring']['token']['enable']['request']['outbox'] = False
cfg['services']['CoAuthoring']['token']['enable']['browser'] = False
# 新版 ONLYOFFICE 默认禁止从私有/回环 IP 下载文档(防 SSRF)；DocScan 让容器直接回源到宿主机网关 IP 取文件，会被拦、转换报 -4
_rfa = cfg['services']['CoAuthoring'].setdefault('request-filtering-agent', {})
_rfa['allowPrivateIPAddress'] = True
_rfa['allowMetaIPAddress'] = True
with open('/etc/onlyoffice/documentserver/local.json', 'w') as f:
    json.dump(cfg, f, indent=2)
" 2>/dev/null && echo "   JWT 已禁用 + 私有 IP 已放行 ✅"
docker exec onlyoffice supervisorctl restart ds:docservice ds:converter >/dev/null 2>&1
sleep 3

# ---------- 3. 确保中文字体 ----------
echo "🔤 检查中文字体…"
if docker exec onlyoffice fc-list ":family=宋体" 2>/dev/null | grep -q .; then
    echo "   字体就绪 ✅"
else
    echo "   注册中文字体（首次或容器重建后，约 1-2 分钟）…"
    # 字体由 docker-compose 挂载到 /usr/share/fonts/truetype/custom（只读）；
    # 若未挂载（裸 docker run），则从仓库 fonts/ 拷入容器。
    if ! docker exec onlyoffice test -s /usr/share/fonts/truetype/custom/simsun.ttc 2>/dev/null; then
        docker exec onlyoffice mkdir -p /usr/share/fonts/truetype/custom
        [ -d fonts ] && docker cp fonts/. onlyoffice:/usr/share/fonts/truetype/custom/
    fi
    docker exec onlyoffice fc-cache -fv >/dev/null 2>&1
    docker exec onlyoffice documentserver-generate-allfonts.sh >/dev/null 2>&1
    echo "   字体注册完成 ✅"
fi

# ---------- 4. 确定 API Key ----------
# 优先环境变量；其次复用已保存的 key（跨重启稳定）；都没有则生成并落盘。
KEY_FILE="$SCRIPT_DIR/.docscan-api-key"
if [ -n "$DOCSCAN_API_KEY" ]; then
    API_KEY="$DOCSCAN_API_KEY"
    echo "🔐 使用环境变量 DOCSCAN_API_KEY"
elif [ -s "$KEY_FILE" ]; then
    API_KEY="$(cat "$KEY_FILE")"
    echo "🔐 复用已保存的 API Key"
else
    API_KEY="$(python3 -c 'import secrets;print(secrets.token_urlsafe(24))' 2>/dev/null || head -c 32 /dev/urandom | base64)"
    printf '%s' "$API_KEY" > "$KEY_FILE"
    chmod 600 "$KEY_FILE" 2>/dev/null
    echo "🔐 已生成并保存新的 API Key ($KEY_FILE)"
fi

# 多 key 列表：首行为主 key；追加更多 key（每行一个，# 注释）即可分发/撤销，热生效
KEYS_FILE="$SCRIPT_DIR/.docscan-api-keys"
if [ ! -s "$KEYS_FILE" ]; then
    printf '%s\n' "$API_KEY" > "$KEYS_FILE"
    chmod 600 "$KEYS_FILE" 2>/dev/null
    echo "🔐 初始化 key 列表 $KEYS_FILE（追加 key 每行一个，免重启生效）"
fi

# ---------- 5. 读取本地配置 (docscan.env，可选) ----------
ENV_FILE="$SCRIPT_DIR/docscan.env"
if [ -f "$ENV_FILE" ]; then
    echo "⚙️  读取本地配置 ($ENV_FILE)"
    set -a
    . "$ENV_FILE"
    set +a
else
    echo "⚙️  无 docscan.env，使用代码默认（可创建该文件自定义，见 README）"
fi

# ---------- 6. 启动 FastAPI ----------
echo "🚀 启动 DocScan API (端口 $PORT)…"
# 把生效的 DOCSCAN_* 落盘，供 status.sh 读取（绕过 /proc ptrace 跨会话限制）
RUNTIME_ENV="$SCRIPT_DIR/.docscan-$PORT.runtime.env"
{ env | grep '^DOCSCAN_' 2>/dev/null; echo "DOCSCAN_API_KEY=$API_KEY"; } | sort -u > "$RUNTIME_ENV"
chmod 600 "$RUNTIME_ENV" 2>/dev/null
# nohup 继承当前 DOCSCAN_*（含 docscan.env）+ 显式 API_KEY
env DOCSCAN_API_KEY="$API_KEY" nohup python3 api.py --port "$PORT" > "$LOG_FILE" 2>&1 &
PID=$!
echo "$PID" > "$PID_FILE"

# ---------- 7. 验证（检查 JSON 内容，防止误判 MinIO 等） ----------
for i in $(seq 1 20); do
    RESP=$(curl -s "http://localhost:$PORT/api/health" 2>/dev/null || true)
    if echo "$RESP" | grep -q '"status".*"ok"'; then
        echo ""
        echo "═══════════════════════════════════════════"
        echo "  ✅ DocScan API 已启动"
        echo "  📡 http://localhost:$PORT"
        echo "  🖥️  前端 Demo: http://localhost:$PORT"
        echo "  📖 Swagger: 已关闭（公网部署不暴露）；端点见 ./status.sh"
        echo "  🔑 API Key: $API_KEY"
        echo "  📋 PID: $PID"
        echo "  📝 日志: $LOG_FILE"
        echo "═══════════════════════════════════════════"
        exit 0
    fi
    sleep 1
done

echo "❌ 启动超时或端口被其他服务占用"
echo "   最后响应: ${RESP:0:200}"
echo "   日志: $LOG_FILE"
exit 1
