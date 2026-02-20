@echo off
:: Gera o .exe do Monitor de Impressoes (chama build.ps1)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1"
if errorlevel 1 pause
