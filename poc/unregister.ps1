#!/usr/bin/env pwsh
$ErrorActionPreference = "Stop"

$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
$dll = Join-Path $PSScriptRoot "bin\Debug\net472\PoC.dll"

if (Test-Path $dll) {
    Write-Host "Unregistering COM server..." -ForegroundColor Cyan
    & $regasm /u $dll
}

$regPath = "HKCU:\Software\Microsoft\Office\Word\Addins\WordVimPoC.Connect"
if (Test-Path $regPath) {
    Write-Host "Removing registry key..." -ForegroundColor Cyan
    Remove-Item -Path $regPath
}

Write-Host "Done." -ForegroundColor Green
