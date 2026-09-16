@echo off
cd /d "%~dp0"
echo Installing Excel Image Assistant - Mohamed v1.7...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Addin.ps1"
echo.
pause
