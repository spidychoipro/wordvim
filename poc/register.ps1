#!/usr/bin/env pwsh
$ErrorActionPreference = "Stop"

$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
$dll = Join-Path $PSScriptRoot "bin\Debug\net472\PoC.dll"

if (-not (Test-Path $dll)) {
    Write-Error "DLL not found: $dll. Run 'dotnet build' first."
    exit 1
}

Write-Host "Registering COM server..." -ForegroundColor Cyan
& $regasm /codebase $dll

Write-Host "Creating Word add-in registry key..." -ForegroundColor Cyan
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Word\Addins\WordVimPoC.Connect"
New-Item -Path $regPath -Force | Out-Null
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "WordVim PoC"
Set-ItemProperty -Path $regPath -Name "Description" -Value "Proof of Concept"
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord

Write-Host ""
Write-Host "Done. Close Word completely, then reopen it." -ForegroundColor Green
