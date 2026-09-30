@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set PYTHONUTF8=1
rem Match the existing launcher: inherit the locally configured LLM credential without printing it.
for /f "tokens=1,2,*" %%A in ('reg query "HKCU\Environment" /v DEEPSEEK_API_KEY 2^>nul ^| findstr /R /C:"DEEPSEEK_API_KEY"') do set "DEEPSEEK_API_KEY=%%C"
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" web_app.py
) else if exist "runtime\python.exe" (
  "runtime\python.exe" web_app.py
) else (
  py -3 web_app.py
)
if errorlevel 1 pause
