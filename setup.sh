#!/bin/bash
# ============================================================
# DocScan — 一键安装 + 启动
# 用法: ./setup.sh [port]      (默认端口 8800，透传给 start.sh)
#
# 幂等，可重复运行。依次做：
#   1. 确保 docker daemon 运行
#   2. 安装 docker compose v2（用户级二进制，免 sudo）
#   3. 安装 pandoc（md→docx）
#   4. 将当前用户加入 docker 组（解决 docker 权限）
#   5. pip 安装 Python 依赖
#   6. 启动 start.sh（首次拉取 ONLYOFFICE ~3GB，需要几分钟）
#
# 设计要点：加入 docker 组后组成员资格需重新登录才永久生效；
# 本脚本用 `sg docker -c` 在同一次运行内临时获得权限直接启动，
# 从而真正做到「安装 + 启动」一气呵成。
# ============================================================
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"
PORT="${1:-}"

# ---------- 颜色（非 tty 时自动关闭）----------
if [ -t 1 ]; then
    C_BLUE=$'\033[1;34m'; C_GREEN=$'\033[1;32m'; C_YELLOW=$'\033[1;33m'
    C_RED=$'\033[1;31m';  C_DIM=$'\033[2m';      C_RESET=$'\033[0m'
else
    C_BLUE=''; C_GREEN=''; C_YELLOW=''; C_RED=''; C_DIM=''; C_RESET=''
fi
step(){ printf '\n%s▸ %s%s\n'  "$C_BLUE"   "$*" "$C_RESET"; }
ok(){   printf '  %s✅ %s%s\n' "$C_GREEN"  "$*" "$C_RESET"; }
warn(){ printf '  %s⚠️  %s%s\n' "$C_YELLOW" "$*" "$C_RESET"; }
dim(){  printf '  %s%s%s\n'    "$C_DIM"    "$*" "$C_RESET"; }
die(){  printf '%s❌ %s%s\n'    "$C_RED"    "$*" "$C_RESET" >&2; exit 1; }

# 不能用 root 整体运行（否则 usermod/sg 逻辑错乱）；脚本内部按需 sudo 提权。
[ "$(id -u)" -eq 0 ] && die "请用普通用户运行本脚本（不要 sudo），脚本会在需要时自动调用 sudo。"

# ---------- 0. 确认 sudo 可用 ----------
step "确认 sudo 可用"
sudo -v || die "需要 sudo 权限完成安装，请使用具备 sudo 权限的用户运行。"

# ---------- 1. 前置命令 ----------
step "环境检查"
command -v python3 >/dev/null || die "未找到 python3"
command -v curl    >/dev/null || die "未找到 curl（请先 sudo apt install curl）"
command -v docker  >/dev/null || die "未找到 docker，请先安装 Docker Engine: https://docs.docker.com/engine/install/"
PYTHON="$(command -v python3)"
ok "python3: $($PYTHON --version 2>&1)  ($PYTHON)"

# ---------- 2. docker daemon ----------
step "Docker daemon"
if systemctl is-active --quiet docker 2>/dev/null; then
    ok "已在运行"
elif sudo systemctl start docker; then
    ok "已启动"
else
    die "无法启动 docker daemon"
fi

# ---------- 3. docker compose v2（用户级二进制，免 sudo）----------
step "Docker Compose v2"
if docker compose version >/dev/null 2>&1; then
    ok "已安装（v$(docker compose version --short 2>/dev/null || echo '?')）"
else
    case "$(uname -m)" in
        x86_64)  CA="linux-x86_64" ;;
        aarch64) CA="linux-aarch64" ;;
        *) die "不支持的架构 $(uname -m)，请手动安装 docker compose" ;;
    esac
    mkdir -p "$HOME/.docker/cli-plugins"
    URL="https://github.com/docker/compose/releases/latest/download/docker-compose-$CA"
    dim "下载 → $HOME/.docker/cli-plugins/docker-compose"
    curl -fSL "$URL" -o "$HOME/.docker/cli-plugins/docker-compose" || die "下载失败：$URL"
    chmod +x "$HOME/.docker/cli-plugins/docker-compose"
    ok "已安装（v$(docker compose version --short 2>/dev/null || echo '?')）"
fi

# ---------- 4. pandoc ----------
step "pandoc（md→docx 转换）"
if command -v pandoc >/dev/null 2>&1; then
    ok "已安装（$(pandoc --version 2>&1 | head -1)）"
else
    dim "sudo apt-get install pandoc"
    sudo apt-get update -qq && sudo apt-get install -y pandoc || die "pandoc 安装失败"
    ok "已安装"
fi

# ---------- 5. docker 组（解决 docker 权限）----------
step "Docker 权限"
NEED_SG=0
if id -nG 2>/dev/null | grep -qw docker; then
    ok "当前用户已在 docker 组"
else
    dim "sudo usermod -aG docker $USER"
    sudo usermod -aG docker "$USER" || die "加入 docker 组失败"
    ok "已加入 docker 组"
    warn "组成员资格需重新登录才永久生效；本次用 sg 临时获得权限启动。"
    command -v sg >/dev/null 2>&1 || die "未找到 sg 命令；请注销重新登录后直接运行 ./start.sh"
    NEED_SG=1
fi

# ---------- 6. Python 依赖 ----------
step "Python 依赖"
dim "$PYTHON -m pip install -r requirements.txt"
"$PYTHON" -m pip install -q -r requirements.txt \
    || die "pip 安装失败（conda/系统 Python 受限时可加 --break-system-packages 重试）"
ok "依赖已就绪"

# ---------- 7. 启动 ----------
step "启动 DocScan"
if [ "$NEED_SG" -eq 1 ]; then
    # 当前 shell 尚无 docker 组 → 用 sg 在 docker 组上下文里启动；
    # 显式传入当前 PATH / HOME，确保 start.sh 用的 python3 与上面装依赖的是同一个。
    sg docker -c "cd '$SCRIPT_DIR' && PATH='$PATH' HOME='$HOME' ./start.sh ${PORT:-}" \
        || die "启动失败（见上方输出）"
else
    ./start.sh ${PORT:-} || die "启动失败（见上方输出）"
fi
