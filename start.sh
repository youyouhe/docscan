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
        echo "   首次使用，拉取并启动 ONLYOFFICE (约 3GB，需要几分钟)…"
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
# 新版 ONLYOFFICE 默认禁止从私有/回环 IP 下载文档(防 SSRF)；DocScan 用 localhost:9999 取源文件会被拦、转换报 -4
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

# ---------- 4. 启动容器内文件服务器 ----------
echo "📁 启动文件服务器…"
docker exec -d onlyoffice sh -c 'cd /tmp && nohup python3 -m http.server 9999 >/tmp/fs.log 2>&1 &' 2>/dev/null
sleep 1
if docker exec onlyoffice curl -s --connect-timeout 3 -o /dev/null http://127.0.0.1:9999/ 2>/dev/null; then
    echo "   文件服务器就绪 ✅ (端口 9999)"
else
    echo "   ⚠️  文件服务器可能启动失败，继续…"
fi

# ---------- 5. 启动 FastAPI ----------
echo "🚀 启动 DocScan API (端口 $PORT)…"
nohup python3 api.py --port "$PORT" > "$LOG_FILE" 2>&1 &
PID=$!
echo "$PID" > "$PID_FILE"

# ---------- 6. 验证（检查 JSON 内容，防止误判 MinIO 等） ----------
for i in $(seq 1 20); do
    RESP=$(curl -s "http://localhost:$PORT/api/health" 2>/dev/null || true)
    if echo "$RESP" | grep -q '"status".*"ok"'; then
        echo ""
        echo "═══════════════════════════════════════════"
        echo "  ✅ DocScan API 已启动"
        echo "  📡 http://localhost:$PORT"
        echo "  📖 Swagger: http://localhost:$PORT/api/docs"
        echo "  🖥️  前端 Demo: http://localhost:$PORT"
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
