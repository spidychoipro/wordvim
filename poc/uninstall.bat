@echo off
echo ========================================
echo   WordVim Uninstaller
echo ========================================
echo.

:: Check admin
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] Please run as administrator.
    echo Right-click this file and select "Run as administrator"
    pause
    exit /b 1
)

set SCRIPT_DIR=%~dp0
set REGASM=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe
set DLL=%SCRIPT_DIR%bin\Debug\net472\PoC.dll

:: Unregister COM server
echo [1/2] Unregistering COM server...
if exist "%DLL%" (
    "%REGASM%" /unregister "%DLL%" 2>nul
)

:: Remove Word add-in registry key
echo [2/2] Removing Word add-in registry key...
reg delete "HKCU\Software\Microsoft\Office\16.0\Word\Addins\WordVimPoC.Connect" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\Word\Addins\WordVimPoC.Connect" /f >nul 2>&1

echo.
echo ========================================
echo   Uninstallation complete!
echo ========================================
echo.
echo Close Word completely, then reopen it.
echo.
pause
