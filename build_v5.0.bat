@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
echo ===============================================
echo          办公文档批量打印器 v5.0 构建脚本
echo ===============================================
echo.

echo [1/6] 检查环境...
python --version
if %errorlevel% neq 0 (
    echo ❌ Python 未安装或未添加到PATH
    pause
    exit /b 1
)

echo [2/6] 检查项目文件结构...
echo 检查核心文件...
if not exist "main.py" (
    echo ❌ 缺少入口文件: main.py
    pause
    exit /b 1
)
echo ✓ main.py 存在

if not exist "src" (
    echo ❌ 缺少源码目录: src
    pause
    exit /b 1
)
echo ✓ src/ 目录存在

if not exist "external\SumatraPDF\SumatraPDF.exe" (
    echo ❌ 缺少关键组件: external\SumatraPDF\SumatraPDF.exe
    echo   SumatraPDF是PDF打印的核心组件，请确保已正确放置
    pause
    exit /b 1
)
echo ✓ SumatraPDF.exe 存在

if not exist "resources\app_icon.ico" (
    echo ❌ 缺少资源文件: resources\app_icon.ico
    pause
    exit /b 1
)
echo ✓ 应用图标存在

if not exist "办公文档批量打印器v5.0.spec" (
    echo ❌ 缺少PyInstaller配置: 办公文档批量打印器v5.0.spec
    pause
    exit /b 1
)
echo ✓ PyInstaller配置文件存在

echo [3/6] 检查依赖...
echo 检查核心依赖...
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

echo 检查项目特定依赖...
python -c "import win32com.client; print('✓ pywin32 已安装')" 2>nul
if %errorlevel% neq 0 (
    echo ⚠️ pywin32 未安装，这可能影响Office文档处理
)

python -c "import PIL; print('✓ Pillow 已安装')" 2>nul
if %errorlevel% neq 0 (
    echo ⚠️ Pillow 未安装，这可能影响图片处理
)

echo 安装所有依赖...
python -m pip install -r requirements.txt --quiet
echo ✓ 依赖检查完成

echo [4/6] 清理旧文件...
if exist "build" (
    echo 删除旧的build目录...
    rmdir /s /q "build" 2>nul
)
if exist "dist" (
    echo 删除旧的dist目录...
    rmdir /s /q "dist" 2>nul
)
echo ✓ 清理完成

echo [5/6] 开始构建...
echo 使用专用spec文件构建，包含SumatraPDF和资源文件...
echo.
echo 构建配置:
echo - 包含SumatraPDF: external\SumatraPDF\
echo - 包含资源文件: resources\
echo - 支持所有文档格式处理器
echo - 优化的tkinterdnd2支持
echo.

python -m PyInstaller --clean "办公文档批量打印器v5.0.spec"

if %errorlevel% neq 0 (
    echo ❌ 构建失败
    echo 请检查上述错误信息，常见问题:
    echo 1. 缺少必要的Python依赖
    echo 2. SumatraPDF文件路径不正确
    echo 3. 文件被杀毒软件阻止
    pause
    exit /b 1
)

echo [6/6] 构建验证...
if not exist "dist\办公文档批量打印器v5.0.exe" (
    echo ❌ 构建的可执行文件不存在
    pause
    exit /b 1
)

echo 检查构建结果...
for %%A in ("dist\办公文档批量打印器v5.0.exe") do (
    set size=%%~zA
    set /a sizeMB=!size!/1024/1024
    echo ✓ 可执行文件大小: !sizeMB! MB
)

echo 检查内嵌资源...
if exist "dist\_internal\external\SumatraPDF\SumatraPDF.exe" (
    echo ✓ SumatraPDF 已正确内嵌
) else (
    echo ⚠️ SumatraPDF 可能未正确内嵌
)

if exist "dist\_internal\resources\app_icon.ico" (
    echo ✓ 资源文件已正确内嵌
) else (
    echo ⚠️ 资源文件可能未正确内嵌
)

echo.
echo ===============================================
echo ✅ 构建完成！
echo.
echo 📁 输出位置: dist\办公文档批量打印器v5.0.exe
echo 📊 文件大小: !sizeMB! MB
echo 🔧 包含组件:
echo   - SumatraPDF (PDF/图片打印核心)
echo   - 所有文档处理器
echo   - GUI界面组件
echo   - 资源文件
echo ===============================================
echo.

echo 是否要测试运行构建的程序？(Y/N)
set /p choice=
if /i "%choice%"=="Y" (
    echo 正在启动测试...
    start "" "dist\办公文档批量打印器v5.0.exe"
    echo 程序已启动，请验证以下功能：
    echo 1. 主界面正常显示
    echo 2. 可以添加各种格式文档
    echo 3. PDF打印功能正常（依赖SumatraPDF）
    echo 4. 图片打印功能正常
)

echo.
echo 构建完成，按任意键退出...
pause >nul 