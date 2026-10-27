
@echo off
setlocal
cd /d "%~dp0"

set "PLAYWRIGHT_BROWSERS_PATH=0"

echo [1/3] Installing Chromium if needed...
python -m playwright install chromium
if errorlevel 1 goto error

echo.
echo [2/3] Building executable...
python -m PyInstaller --noconfirm --onedir --console --name MCXDownloader --distpath dist_new --workpath build_temp --collect-all playwright --collect-all pypyodbc run.py
if errorlevel 1 goto error

echo.
echo [3/3] Checking bundled Chromium...
if not exist "dist\MCXDownloader\_internal\playwright\driver\package\.local-browsers\chromium-1248\chrome-win64\chrome.exe" (
    echo Chromium was not found at the expected location.
    echo Check the actual Chromium directory before deploying.
    pause
    exit /b 1
)

echo.
echo BUILD SUCCESSFUL
echo Application: dist\MCXDownloader\MCXDownloader.exe
pause
exit /b 0

:error
echo.
echo BUILD FAILED.
pause
exit /b 1
