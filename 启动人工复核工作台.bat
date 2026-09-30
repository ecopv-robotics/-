@echo off
setlocal EnableExtensions

set "APP_DIR=%~dp0"
pushd "%APP_DIR%" || (
  echo Failed to open the application directory.
  pause
  exit /b 1
)

where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
  python "%APP_DIR%workbench_launcher.py"
  set "EXIT_CODE=%ERRORLEVEL%"
  popd
  exit /b %EXIT_CODE%
)

where py >nul 2>nul
if %ERRORLEVEL% EQU 0 (
  py -3 "%APP_DIR%workbench_launcher.py"
  set "EXIT_CODE=%ERRORLEVEL%"
  popd
  exit /b %EXIT_CODE%
)

echo 未找到 Python 3，请先安装 Python 3.10 或更高版本。
pause
popd
exit /b 1
