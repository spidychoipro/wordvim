@echo off
echo ========================================
echo   WordVim Installer
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

:: Check DLL exists
if not exist "%DLL%" (
    echo [ERROR] PoC.dll not found.
    echo Please build first: dotnet build
    pause
    exit /b 1
)

:: Register COM server
echo [1/2] Registering COM server...
"%REGASM%" /codebase "%DLL%"
if %errorLevel% neq 0 (
    echo [ERROR] COM registration failed.
    pause
    exit /b 1
)

:: Create Word add-in registry key
echo [2/2] Creating Word add-in registry key...
reg add "HKCU\Software\Microsoft\Office\16.0\Word\Addins\WordVimPoC.Connect" /v FriendlyName /t REG_SZ /d "WordVim" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\Word\Addins\WordVimPoC.Connect" /v Description /t REG_SZ /d "Vim keybindings for Microsoft Word" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\Word\Addins\WordVimPoC.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul

echo.
echo ========================================
echo   Installation complete!
echo ========================================
echo.
echo Close Word completely, then reopen it.
echo You should see "-- NORMAL --" in the title bar.
echo.
pause
