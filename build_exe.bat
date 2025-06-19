@echo off
echo Building GST Scrutiny application...
echo.

echo Step 1: Building frontend...
cd ..\frontend\gst-scrutiny-ui
call npm run build

if %ERRORLEVEL% NEQ 0 (
    echo Frontend build failed!
    cd ..\..\backend
    pause
    exit /b %ERRORLEVEL%
)

echo Frontend build completed successfully!
echo.

echo Step 2: Returning to backend and building executable...
cd ..\..\backend

python -m PyInstaller --onefile --noconfirm --icon=logo.ico --add-data "..\frontend\gst-scrutiny-ui\build;frontend/gst-scrutiny-ui/build" main.py

echo.
if %ERRORLEVEL% EQU 0 (
    echo Build completed successfully!
    echo Executable should be in the 'dist' folder.
) else (
    echo Build failed with error code %ERRORLEVEL%
)

pause