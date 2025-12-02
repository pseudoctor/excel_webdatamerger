@echo off
chcp 65001 >nul
echo 🔧 虚拟环境修复工具
echo =======================================
echo.

REM 切换到脚本所在目录
cd /d "%~dp0"

REM 检查 Python
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
    echo ❌ 未检测到 Python
    pause
    exit /b 1
)

echo 检测到 Python:
%PYTHON_CMD% --version
echo.

REM 询问是否删除旧环境
if exist "venv\" (
    echo 发现已存在的虚拟环境 (venv\)
    echo.
    set /p "CONFIRM=是否删除并重建？(y/n): "
    if /i "!CONFIRM!"=="y" (
        echo.
        echo 🗑️  删除旧环境...
        rmdir /s /q venv
        if exist "venv\" (
            echo ❌ 删除失败，请手动删除 venv 文件夹后重试
            pause
            exit /b 1
        )
    ) else (
        echo 取消操作
        pause
        exit /b 0
    )
)

REM 创建新环境
echo.
echo 🧱 创建新的虚拟环境...
%PYTHON_CMD% -m venv venv
if errorlevel 1 (
    echo ❌ 创建失败
    pause
    exit /b 1
)

REM 激活环境
echo.
echo 📦 激活虚拟环境...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo ❌ 激活失败
    pause
    exit /b 1
)

REM 升级pip
echo.
echo 📦 升级 pip...
python -m pip install --upgrade pip

REM 安装依赖
echo.
echo 📦 安装依赖 (详细模式)...
echo.
pip install -r requirements.txt -v
if errorlevel 1 (
    echo.
    echo ❌ 安装失败
    pause
    exit /b 1
)

REM 验证安装
echo.
echo 🔍 验证安装...
echo.
python -c "import sys; print(f'Python: {sys.version}')"
python -c "import pandas; print(f'pandas: {pandas.__version__}')"
python -c "import openpyxl; print(f'openpyxl: {openpyxl.__version__}')"
python -c "import xlrd; print(f'xlrd: {xlrd.__version__}')"
python -c "import chardet; print(f'chardet: {chardet.__version__}')"

if errorlevel 1 (
    echo.
    echo ❌ 验证失败
    pause
    exit /b 1
)

echo.
echo ✅ 虚拟环境修复完成！
echo.
echo 现在可以运行 run_windows.bat 启动程序
echo.
pause
