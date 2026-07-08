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
    echo "   容器未运行，正在启动…"
    docker start onlyoffice
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
with open('/etc/onlyoffice/documentserver/local.json', 'w') as f:
    json.dump(cfg, f, indent=2)
" 2>/dev/null && echo "   JWT 已禁用 ✅"
docker exec onlyoffice supervisorctl restart ds:docservice ds:converter >/dev/null 2>&1
sleep 3

# ---------- 3. 确保中文字体 ----------
echo "🔤 检查中文字体…"
FONT_MISSING=$(docker exec onlyoffice fc-list ":family=宋体" 2>/dev/null | wc -l)
if [ "$FONT_MISSING" -eq 0 ] && [ -d "/mnt/c/Windows/Fonts" ]; then
    echo "   安装 Windows 中文字体…"
    docker exec onlyoffice mkdir -p /var/www/onlyoffice/documentserver/core-fonts/cn
    for f in simsun.ttc simhei.ttf simkai.ttf simfang.ttf; do
        [ -f "/mnt/c/Windows/Fonts/$f" ] && docker cp "/mnt/c/Windows/Fonts/$f" "onlyoffice:/var/www/onlyoffice/documentserver/core-fonts/cn/"
    done
    docker exec onlyoffice fc-cache -fv >/dev/null 2>&1
    docker exec onlyoffice supervisorctl restart ds:converter >/dev/null 2>&1
    sleep 3
    echo "   字体安装完成 ✅"
else
    echo "   字体就绪 ✅"
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
