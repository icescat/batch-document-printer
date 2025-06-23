@echo off
chcp 65001 >nul
echo ===============================================
echo          办公文档批量打印器 v4.1 构建脚本
echo ===============================================
echo.

echo [1/4] 检查环境...
python --version
if %errorlevel% neq 0 (
    echo ❌ Python 未安装或未添加到PATH
    pause
    exit /b 1
)

echo [2/4] 检查依赖...
python -c "import tkinterdnd2; print('✓ tkinterdnd2 已安装')" 2>nul
if %errorlevel% neq 0 (
    echo ❌ tkinterdnd2 未安装，正在安装...
    pip install tkinterdnd2==0.3.0
)

python -c "import PyInstaller; print('✓ PyInstaller 已安装')" 2>nul
if %errorlevel% neq 0 (
    echo ❌ PyInstaller 未安装，正在安装...
    pip install PyInstaller==6.3.0
)

echo [3/4] 清理旧文件...
if exist "build" rmdir /s /q "build" 2>nul
if exist "dist" rmdir /s /q "dist" 2>nul
echo ✓ 清理完成

echo [4/4] 开始构建...
echo 使用专用spec文件构建，包含tkinterdnd2支持...
pyinstaller --clean "办公文档批量打印器v4.1.spec"

if %errorlevel% neq 0 (
    echo ❌ 构建失败
    pause
    exit /b 1
)

echo.
echo ===============================================
echo ✅ 构建完成！
echo 📁 输出位置: dist\办公文档批量打印器v4.1.exe
if exist "dist\办公文档批量打印器v4.1.exe" (
    for %%A in ("dist\办公文档批量打印器v4.1.exe") do (
        set size=%%~zA
        set /a sizeMB=!size!/1024/1024
        echo 📊 文件大小: !sizeMB! MB
    )
)
echo ===============================================
echo.

echo 是否要测试运行构建的程序？(Y/N)
set /p choice=
if /i "%choice%"=="Y" (
    echo 正在启动...
    start "" "dist\办公文档批量打印器v4.1.exe"
)

echo 构建完成，按任意键退出...
pause >nul 