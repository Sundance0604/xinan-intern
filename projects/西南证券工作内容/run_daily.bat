@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scriptsun_daily.ps1" %*
exit /b %errorlevel%
