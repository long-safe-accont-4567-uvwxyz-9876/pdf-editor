@echo off
title PDF Editor
cd /d "%~dp0"

echo.
echo   ==========================================
echo      PDF 编辑器 - 正在启动
echo   ==========================================
echo.

set "PY="

if exist "C:\Users\lenovo\.workbuddy\binaries\python\versions\3.13.12\python.exe" set "PY=C:\Users\lenovo\.workbuddy\binaries\python\versions\3.13.12\python.exe"
if not defined PY if exist "%~dp0python\python.exe" set "PY=%~dp0python\python.exe"
if not defined PY (
    where python >nul 2>nul
    if not errorlevel 1 set "PY=python"
)

if not defined PY (
    echo   [错误] 没有找到 Python 运行环境。
    echo   请先安装 Python 3.10 或更高版本。
    echo.
    pause
    exit /b 1
)

"%PY%" -c "import pymupdf" >nul 2>nul
if errorlevel 1 (
    echo   [错误] 缺少 PyMuPDF 组件。
    echo.
    echo   请执行以下命令安装：
    echo     "%PY%" -m pip install pymupdf
    echo.
    pause
    exit /b 1
)

"%PY%" "%~dp0app.py"

echo.
echo   程序已退出。
pause
