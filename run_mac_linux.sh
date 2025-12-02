#!/bin/bash
echo "🚀 正在启动 excel_webdatamerger v0.1.0 ..."
echo "---------------------------------------"

# 切换到脚本所在目录
cd "$(dirname "$0")"

# 检查 Python
if ! command -v python3 &> /dev/null; then
  echo "❌ 未检测到 python3，请先安装 Python 3.9+"
  exit 1
fi

# 虚拟环境目录
VENV_DIR="venv"

# 若虚拟环境不存在则创建
if [ ! -d "$VENV_DIR" ]; then
  echo "🧱 正在创建虚拟环境..."
  python3 -m venv "$VENV_DIR"
fi

# 激活虚拟环境
source "$VENV_DIR/bin/activate"

# 安装依赖
echo "📦 检查并安装依赖..."
pip install --upgrade pip
pip install -r requirements.txt

# 运行程序
echo "✅ 启动 GUI 程序..."
python3 main.py
