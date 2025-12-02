@echo off
chcp 65001 >nul
echo 🔍 环境检查工具
echo =======================================
echo.

REM 切换到脚本所在目录
cd /d "%~dp0"

REM 1. 检查 Python
echo [1] 检查 Python 安装
set PYTHON_CMD=
where python >nul 2>&1
if not errorlevel 1 (
    set PYTHON_CMD=python
    python --version
    echo ✅ 找到 python 命令
) else (
    where py >nul 2>&1
    if not errorlevel 1 (
        set PYTHON_CMD=py
        py --version
        echo ✅ 找到 py 命令
    ) else (
        echo ❌ 未找到 Python
    )
)
echo.

REM 2. 检查虚拟环境
echo [2] 检查虚拟环境
if exist "venv\" (
    echo ✅ 虚拟环境存在: venv\
    if exist "venv\Scripts\python.exe" (
        echo ✅ Python 解释器: venv\Scripts\python.exe
    ) else (
        echo ❌ 虚拟环境损坏 (缺少 python.exe)
    )
) else (
    echo ❌ 虚拟环境不存在
)
echo.

REM 3. 检查依赖文件
echo [3] 检查依赖文件
if exist "requirements.txt" (
    echo ✅ requirements.txt 存在
    echo 内容:
    type requirements.txt
) else (
    echo ❌ requirements.txt 不存在
)
echo.

REM 4. 检查主程序
echo [4] 检查主程序
if exist "main.py" (
    echo ✅ main.py 存在
) else (
    echo ❌ main.py 不存在
)
echo.

REM 5. 如果虚拟环境存在，检查已安装的包
if exist "venv\Scripts\activate.bat" (
    echo [5] 检查已安装的包
    call venv\Scripts\activate.bat
    echo.
    pip list | findstr /i "pandas openpyxl xlrd chardet"
    if errorlevel 1 (
        echo ⚠️  关键依赖可能未安装
    ) else (
        echo.
        echo ✅ 找到部分/全部依赖
    )
    echo.

    echo [6] 测试导入模块
    python -c "import pandas; print('✅ pandas')" 2>nul || echo ❌ pandas 导入失败
    python -c "import openpyxl; print('✅ openpyxl')" 2>nul || echo ❌ openpyxl 导入失败
    python -c "import xlrd; print('✅ xlrd')" 2>nul || echo ❌ xlrd 导入失败
    python -c "import chardet; print('✅ chardet')" 2>nul || echo ❌ chardet 导入失败
)
echo.

echo =======================================
echo 检查完成
echo.
echo 如果发现问题，请运行 fix_venv.bat 修复
echo.
pause
