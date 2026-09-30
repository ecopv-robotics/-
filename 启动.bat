@echo off
setlocal EnableExtensions

set "APP_DIR=%~dp0"
pushd "%APP_DIR%" || (
  echo Failed to open the application directory.
  pause
  exit /b 1
)

rem Read the user environment variable so pythonw can access the API key.
for /f "tokens=1,2,*" %%A in ('reg query "HKCU\Environment" /v DEEPSEEK_API_KEY 2^>nul ^| findstr /R /C:"DEEPSEEK_API_KEY"') do set "DEEPSEEK_API_KEY=%%C"

if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" "%APP_DIR%main.py"
) else (
  py -3 "%APP_DIR%main.py"
)
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
