@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

set VENV=%~dp0.venv
set VENV_PY=%VENV%\Scripts\python.exe
set VENV_PYW=%VENV%\Scripts\pythonw.exe
set SYNC=%~dp0src\sync.py
set WIZ=%~dp0src\setup_wizard.py
set APP=%~dp0GCalSync.exe

REM ── Detect first run (no configuration) ──────────────────────────────────────
if not exist "%~dp0google_token.json" goto :wizard
if not exist "%~dp0ms_token_cache.bin" goto :wizard
goto :menu


REM ============================================================================
REM  WIZARD - initial setup
REM ============================================================================
:wizard
title GCal -> Teams  [Initial setup]
cls
if exist "%APP%" (
    "%APP%" setup
    goto :check_after_wizard
)
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found on this machine.
    echo.
    echo Download at: https://www.python.org/downloads/
    echo IMPORTANT: check "Add Python to PATH" during installation.
    echo After installing, close and reopen this file.
    echo.
    pause
    exit /b
)
python "%WIZ%"

:check_after_wizard
if not exist "%~dp0google_token.json" (
    pause
    exit /b
)
if not exist "%~dp0ms_token_cache.bin" (
    pause
    exit /b
)
goto :menu


REM ============================================================================
REM  MENU
REM ============================================================================
:menu
title GCal -> Teams
cls
echo ============================================================
echo   Google Calendar -^> Microsoft Teams Sync
echo ============================================================
echo.
echo   [1] Open monitor  (account status and manual sync)
echo   [2] Test sync now
echo   [3] Run in loop  (window stays open, syncs every 5 min)
echo   [4] Schedule auto-start on Windows login
echo   [5] Remove auto-start
echo   [6] Clean temp files  (log and Python cache)
echo   [7] Re-login  (if token expired)
echo   [8] Remove duplicate events from Outlook
echo   [0] Exit
echo.
choice /c 123456780 /n

if errorlevel 9 goto :sair
if errorlevel 8 goto :dedup
if errorlevel 7 goto :refazer_login_menu
if errorlevel 6 goto :limpar
if errorlevel 5 goto :remover_tarefa
if errorlevel 4 goto :agendar_tarefa
if errorlevel 3 goto :rodar_loop
if errorlevel 2 goto :testar
if errorlevel 1 goto :abrir_monitor


:abrir_monitor
cls
if exist "%APP%" (
    start "" "%APP%"
) else (
    echo GCalSync.exe not found. Run build.bat to generate it.
    echo.
    pause
)
goto :menu


:testar
cls
echo Running a sync...
echo.
call :run_sync once
echo.
pause
goto :menu


:rodar_loop
cls
echo Syncing every 5 minutes. Press Ctrl+C to stop.
echo.
call :run_sync run
pause
goto :menu


:agendar_tarefa
call :fn_agendar
goto :menu


:remover_tarefa
cls
echo Removing auto-start...
schtasks /delete /tn "GCal-Teams-Sync" /f >nul 2>&1
if %errorlevel%==0 (
    echo Removed. The sync will no longer start with Windows.
) else (
    echo Task not found or already removed.
)
echo.
pause
goto :menu


:limpar
cls
echo Cleaning temp files...
if exist "src\__pycache__" rmdir /s /q "src\__pycache__"
if exist "sync.log" del /q "sync.log"
echo Done.
echo.
pause
goto :menu


:refazer_login_menu
:refazer_login
cls
echo Deleting tokens and re-logging in...
echo.
if exist "%~dp0google_token.json" del /q "%~dp0google_token.json"
if exist "%~dp0ms_token_cache.bin" del /q "%~dp0ms_token_cache.bin"
call :run_sync login
echo.
pause
goto :menu


:dedup
cls
echo Scanning Outlook for duplicate [GCal] events...
echo.
if exist "%APP%" (
    "%APP%" dedup
) else (
    "%VENV_PY%" "%~dp0src\dedup.py"
)
echo.
pause
goto :menu


:sair
endlocal
exit /b


REM ============================================================================
REM  FUNCTION: run sync with the given argument
REM  Prefers GCalSync.exe; falls back to Python + script if not found.
REM ============================================================================
:run_sync
if exist "%APP%" (
    "%APP%" %1
) else (
    "%VENV_PY%" "%SYNC%" %1
)
exit /b %errorlevel%


REM ============================================================================
REM  FUNCTION: schedule auto-start
REM ============================================================================
:fn_agendar
cls
echo Registering in Windows Task Scheduler...
echo.
schtasks /create /tn "GCal-Teams-Sync" /tr "wscript.exe \"%~dp0run-oculto.vbs\"" /sc ONLOGON /ru "%USERNAME%" /f
if %errorlevel%==0 (
    echo.
    echo Done. The sync will start automatically on each login.
    echo No window needs to be opened.
    echo.
    echo To verify: Task Scheduler -^> GCal-Teams-Sync
    echo To remove: option [5] from this menu.
) else (
    echo.
    echo Error registering. Run this file as Administrator:
    echo right-click SYNC.bat -^> Run as administrator
)
echo.
pause
goto :eof
