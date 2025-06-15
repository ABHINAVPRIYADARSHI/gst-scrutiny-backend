@echo off
REM Navigate to frontend and build React app
echo ===== Building React frontend =====
cd ..\frontend\gst-scrutiny-ui

REM Check if node_modules exists to skip re-installation
IF NOT EXIST "node_modules" (
    echo Installing dependencies...
    npm install
)

REM Build React app
npm run build

IF %ERRORLEVEL% NEQ 0 (
    echo React build failed. Exiting...
    pause
    exit /b %ERRORLEVEL%
)

REM Go back to backend
cd ..\..\backend

echo ===== Creating executable with PyInstaller =====
python -m PyInstaller --onefile --noconfirm --add-data "..\frontend\gst-scrutiny-ui\build;frontend/gst-scrutiny-ui/build" main.py

echo ===== Build Complete =====
pause
