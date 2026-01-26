@echo off
setlocal

cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_exe.ps1"

if errorlevel 1 (
  echo.
  echo Build failed.
  pause
  exit /b 1
)

echo.\necho Build finished. Output: "%~dp0dist\Density2excel\Density2excel.exe"
pause
