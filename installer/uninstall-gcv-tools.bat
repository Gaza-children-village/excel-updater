@echo off
rem Removes GCV Tools for this Windows user (Excel must be closed).
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall-gcv-tools.ps1"
echo.
pause
