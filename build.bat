@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo   Build - GCal Teams Sync
echo   Generates GCalSync.exe via PyInstaller
echo ============================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: Virtual environment not found.
    echo Run SYNC.bat first to complete the installation.
    echo.
    pause
    exit /b 1
)

echo [1/3] Checking PyInstaller...
.venv\Scripts\python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo PyInstaller not found, installing...
    .venv\Scripts\pip install pyinstaller -q
    if errorlevel 1 (
        echo ERROR installing PyInstaller.
        pause
        exit /b 1
    )
) else (
    echo PyInstaller already installed.
)

echo.
echo [2/3] Building GCalSync.exe...
.venv\Scripts\pyinstaller ^
    --onefile ^
    --name GCalSync ^
    --icon assets\icon.ico ^
    --distpath . ^
    --workpath build\work ^
    --specpath build ^
    --collect-all msal ^
    --collect-all google ^
    --collect-all googleapiclient ^
    --collect-all pystray ^
    --collect-all PIL ^
    --hidden-import win32timezone ^
    --hidden-import win32com.client ^
    --hidden-import pywintypes ^
    --hidden-import dedup ^
    --noconfirm ^
    --clean ^
    src\app.py
if errorlevel 1 (
    echo ERROR building GCalSync.exe.
    pause
    exit /b 1
)

echo.
echo [3/3] Cleaning build artifacts...
if exist "build" rmdir /s /q "build"

echo.
echo ============================================================
echo   Build complete!
echo ============================================================
echo.
if exist "GCalSync.exe" echo   GCalSync.exe - unified app (wizard + monitor + sync)
echo.
echo   Double-click GCalSync.exe to open the monitor.
echo   On first run, the setup wizard opens automatically.
echo.
echo   DO NOT include in distributed files:
echo     google_credentials.json  config.json  google_token.json
echo     ms_token_cache.bin       sync_state.db
echo.
pause
