@echo off
chcp 65001 >nul
echo 🚀 正在启动 excel_datamerger v1.0 ...
echo ---------------------------------------

REM 切换到脚本所在目录
cd /d "%~dp0"

REM 检查 py 启动器
where py >nul 2>&1
if errorlevel 1 (
    echo ❌ 未检测到 Python，请先安装 Python 3.9+
    echo.
    echo 下载地址: https://www.python.org/downloads/
    echo 注意: 安装时勾选 "Add Python to PATH"
    pause
    exit /b 1
)

REM 显示Python版本
echo 检测到 Python:
py --version

REM 虚拟环境目录
set VENV_DIR=venv

REM 若虚拟环境不存在则创建
if not exist "%VENV_DIR%\" (
    echo.
    echo 🧱 正在创建虚拟环境...
    py -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo ❌ 创建虚拟环境失败
        pause
        exit /b 1
    )
)

REM 激活虚拟环境
echo.
echo 📦 激活虚拟环境...
call "%VENV_DIR%\Scripts\activate.bat"
if errorlevel 1 (
    echo ❌ 激活虚拟环境失败
    pause
    exit /b 1
)

REM 升级pip
echo.
echo 📦 升级 pip...
python -m pip install --upgrade pip --quiet

REM 安装依赖
echo.
echo 📦 检查并安装依赖...
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo ❌ 安装依赖失败
    pause
    exit /b 1
)

REM 运行程序
echo.
echo ✅ 启动 GUI 程序...
echo.
python main.py

REM 程序结束后暂停
if errorlevel 1 (
    echo.
    echo ❌ 程序运行出错
    pause
)
