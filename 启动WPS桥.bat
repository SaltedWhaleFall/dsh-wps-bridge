@echo off
title WPS Bridge 8765
cd /d "%~dp0"

set "PS1=C:\Users\admin\wps-skills\scripts\wps-exec.ps1"
if not exist "%PS1%" set "PS1=%~dp0wps-exec.ps1"
if not exist "%PS1%" (
    echo.
    echo   [X] Cannot find wps-exec.ps1
    echo       Expected: C:\Users\admin\wps-skills\scripts\wps-exec.ps1
    echo.
    pause
    exit /b 1
)

echo.
echo   Starting WPS bridge on port 8765 ...
echo   Keep this window open. Closing it stops the bridge.
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Port 8765 -AllowScript
echo.
echo   Bridge stopped.
pause
