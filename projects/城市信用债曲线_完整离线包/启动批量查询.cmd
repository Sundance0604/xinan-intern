@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
where node.exe >nul 2>nul
if errorlevel 1 goto bundled
node.exe "%~dp0batch-server.cjs" --open %*
goto done
:bundled
if not exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" goto missing
"%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" "%~dp0batch-server.cjs" --open %*
goto done
:missing
echo Node.js was not found. Please install Node.js 18 or later.
:done
if /i "%~1"=="--check" exit /b
pause
endlocal
