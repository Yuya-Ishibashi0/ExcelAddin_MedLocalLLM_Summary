@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build-iexpress.ps1" %*
if errorlevel 1 (
  echo Build failed.
  pause
  exit /b 1
)
echo Build finished.
pause
