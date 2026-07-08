#!/bin/bash
# ============================================================
# DocScan API — 停止脚本
# 用法: ./stop.sh [port]
# 不指定端口则停止所有正在运行的实例
# ============================================================
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PORT="${1:-}"

stop_one() {
    local pf="$1"
    local pid
    pid=$(cat "$pf" 2>/dev/null)
    if [ -z "$pid" ]; then
        rm -f "$pf"
        return
    fi
    if ! kill -0 "$pid" 2>/dev/null; then
        echo "   PID $pid 已不存在，清理"
        rm -f "$pf"
        return
    fi
    echo "   🛑 停止 PID $pid…"
    kill "$pid" 2>/dev/null
    for i in $(seq 1 10); do
        if ! kill -0 "$pid" 2>/dev/null; then
            echo "   ✅ 已停止"
            rm -f "$pf"
            return
        fi
        sleep 1
    done
    kill -9 "$pid" 2>/dev/null
    rm -f "$pf"
    echo "   ✅ 已强制停止"
}

if [ -n "$PORT" ]; then
    PF="$SCRIPT_DIR/.docscan-$PORT.pid"
    if [ -f "$PF" ]; then
        echo "🛑 停止端口 $PORT 的 DocScan…"
        stop_one "$PF"
    else
        echo "⚠️  未找到端口 $PORT 的 PID 文件"
    fi
else
    PIDS=$(ls "$SCRIPT_DIR"/.docscan-*.pid 2>/dev/null)
    if [ -z "$PIDS" ]; then
        echo "⚠️  未找到 PID 文件（可能未运行）"
        # 尝试通过常见端口清理
        for p in 8800 9000 8080; do
            if lsof -ti ":$p" 2>/dev/null | xargs -r ps -p 2>/dev/null | grep -q api.py; then
                echo "   发现端口 $p 上有残留进程，清理…"
                kill $(lsof -ti ":$p") 2>/dev/null
                echo "   ✅ 已清理端口 $p"
            fi
        done
        exit 0
    fi
    echo "🛑 停止所有 DocScan 实例…"
    for pf in $PIDS; do
        stop_one "$pf"
    done
fi
echo "✅ 完成"
