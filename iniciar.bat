@echo off
cd /d "%~dp0"
title Automacao Teams via PowerShell
echo Desbloqueando scripts para execucao no ambiente corporativo...
powershell.exe -NoProfile -Command "Get-ChildItem -Path '%~dp0' -Recurse | Unblock-File -ErrorAction SilentlyContinue"
echo Iniciando o aplicativo de Automacao Teams...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0App.ps1"
if %errorlevel% neq 0 (
    echo.
    echo Ocorreu um erro ao executar o aplicativo.
    pause
)
