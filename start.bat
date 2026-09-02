@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.10+ 并勾选 Add to PATH
    pause
    exit /b 1
)

python -c "import win32clipboard, docx" 1>nul 2>nul
if errorlevel 1 (
    echo [提示] 正在安装依赖...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [错误] 依赖安装失败
        pause
        exit /b 1
    )
)

python main.py
if errorlevel 1 (
    echo.
    echo [错误] 程序异常退出
    pause
)
