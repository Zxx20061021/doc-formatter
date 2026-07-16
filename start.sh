#!/bin/bash
# 公文格式助手 - 一键启动脚本
# 启动 Flask 服务器 + Cloudflare 隧道，自动获取公网地址

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

export PATH="$HOME/bin:$PATH"

echo "═══════════════════════════════════════"
echo "  📄 公文格式助手 - 启动中..."
echo "═══════════════════════════════════════"

# 检查依赖
if ! command -v python3 &> /dev/null; then
    echo "❌ 未找到 python3，请先安装 Python"
    exit 1
fi

if ! command -v cloudflared &> /dev/null; then
    echo "❌ 未找到 cloudflared，请先运行: curl -fsSL https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-darwin-amd64.tgz | tar xz -C ~/bin/"
    exit 1
fi

# 安装 Python 依赖（如需要）
if [ ! -d ".venv" ]; then
    echo "📦 首次运行，安装 Python 依赖..."
    python3 -m venv .venv
    source .venv/bin/activate
    pip install -r requirements.txt -q
else
    source .venv/bin/activate
fi

# 创建必要目录
mkdir -p uploads temp

# 启动 Flask 服务器（后台）
echo "🚀 启动本地服务器..."
python3 app.py &
FLASK_PID=$!

# 等待 Flask 启动
sleep 3

# 检查 Flask 是否启动成功
if ! curl -s -o /dev/null http://localhost:5800/ 2>/dev/null; then
    echo "❌ 本地服务器启动失败"
    kill $FLASK_PID 2>/dev/null
    exit 1
fi

echo "✅ 本地服务器已启动"

# 启动 Cloudflare 隧道
echo "🌐 正在创建公网隧道..."
echo ""
echo "───────────────────────────────────────"

cloudflared tunnel --url http://localhost:5800 2>&1 | while IFS= read -r line; do
    # 捕获公网 URL
    if echo "$line" | grep -q "trycloudflare.com"; then
        URL=$(echo "$line" | grep -oE 'https://[a-z0-9-]+\.trycloudflare\.com')
        if [ -n "$URL" ]; then
            echo ""
            echo "╔═══════════════════════════════════════╗"
            echo "║  ✅ 公网地址已生成！                   ║"
            echo "║                                       ║"
            echo "║  $URL"
            echo "║                                       ║"
            echo "║  微信/手机/电脑均可访问                ║"
            echo "║  关闭此窗口将停止服务                  ║"
            echo "╚═══════════════════════════════════════╝"
            echo ""
            # 将 URL 复制到剪贴板
            echo -n "$URL" | pbcopy 2>/dev/null && echo "📋 已复制到剪贴板"
        fi
    fi
    # 只显示关键日志
    if echo "$line" | grep -qE "ERR|INF \+|INF \||https://"; then
        echo "$line"
    fi
done

# 清理
kill $FLASK_PID 2>/dev/null
echo "服务已停止"
