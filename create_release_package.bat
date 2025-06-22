@echo off
echo 正在创建发布包...

REM 创建发布目录
if not exist "release" mkdir release
if exist "release\*" del /q "release\*"

REM 复制可执行文件
copy "dist\办公文档批量打印器.exe" "release\"

REM 复制说明文件
copy "RELEASE_NOTES.txt" "release\"

REM 复制README（可选）
copy "README.md" "release\项目说明.md"

echo 发布包已创建在 release 目录中
echo 包含以下文件:
dir "release" /b

echo.
echo 您现在可以：
echo 1. 将 release 目录中的文件打包成 ZIP
echo 2. 在 GitHub Release 中上传这些文件
echo.
pause