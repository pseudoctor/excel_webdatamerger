@echo off
chcp 65001 >nul
echo 🚀 正在启动 excel_webdatamerger v0.1.0 ...
echo ---------------------------------------

REM 切换到脚本所在目录
cd /d "%~dp0"

REM 检查 Python (优先 python, 其次 py)
set PYTHON_CMD=
where python >nul 2>&1
if not errorlevel 1 (
    set PYTHON_CMD=python
) else (
    where py >nul 2>&1
    if not errorlevel 1 (
        set PYTHON_CMD=py
    )
)

if "%PYTHON_CMD%"=="" (
    echo ❌ 未检测到 Python，请先安装 Python 3.9+
    echo.
    echo 下载地址: https://www.python.org/downloads/
    echo 注意: 安装时勾选 "Add Python to PATH"
    pause
    exit /b 1
)

REM 显示Python版本
echo 检测到 Python:
%PYTHON_CMD% --version

REM 虚拟环境目录
set VENV_DIR=venv

REM 若虚拟环境不存在则创建
if not exist "%VENV_DIR%\" (
    echo.
    echo 🧱 正在创建虚拟环境...
    %PYTHON_CMD% -m venv "%VENV_DIR%"
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
python -m pip install --upgrade pip
if errorlevel 1 (
    echo ⚠️  升级 pip 失败，继续尝试安装依赖...
)

REM 安装依赖
echo.
echo 📦 检查并安装依赖...
echo 正在安装: pandas, openpyxl, xlrd, chardet
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ❌ 安装依赖失败，请检查网络连接或手动执行以下命令：
    echo    venv\Scripts\activate
    echo    pip install -r requirements.txt
    pause
    exit /b 1
)

REM 验证关键模块
echo.
echo 🔍 验证关键模块安装...
python -c "import pandas; import openpyxl; import xlrd; print('✅ 所有依赖已正确安装')"
if errorlevel 1 (
    echo.
    echo ❌ 模块验证失败，尝试重新安装...
    pip install --force-reinstall pandas openpyxl xlrd chardet
    if errorlevel 1 (
        echo ❌ 重新安装失败
        pause
        exit /b 1
    )
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
