@echo off
cd /d "%~dp0"
echo Building Excel Image Assistant - Mohamed v1.7...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build-Addin.ps1"
echo.
echo If the build says PASS, run: Install Add-in.cmd
pause
