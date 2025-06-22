@echo off
echo Building Batch Document Printer v2.0...

REM Clean old build files
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

echo Installing dependencies...
pip install -r requirements.txt

echo Starting packaging...
pyinstaller --onefile --windowed --name="BatchDocPrinter_v2.0" --icon=resources/app_icon.ico --add-data="resources;resources" --hidden-import=comtypes.client --hidden-import=win32api --hidden-import=win32print main.py

if exist "dist\BatchDocPrinter_v2.0.exe" (
    echo.
    echo Build successful!
    echo Executable location: dist\BatchDocPrinter_v2.0.exe
    echo File size:
    dir "dist\BatchDocPrinter_v2.0.exe" | findstr "BatchDocPrinter_v2.0.exe"
) else (
    echo.
    echo Build failed!
    echo Please check error messages.
)

echo.
pause 