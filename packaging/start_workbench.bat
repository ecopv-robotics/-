@echo off
setlocal
cd /d "%~dp0"

if not exist "%~dp0config.yaml" (
  if exist "%~dp0config.example.yaml" (
    copy /y "%~dp0config.example.yaml" "%~dp0config.yaml" >nul
  ) else (
    echo Missing config.yaml and config.example.yaml.
    pause
    exit /b 2
  )
)

set "PADDLE_PDX_CACHE_HOME=%~dp0model_cache"
set "PATH=%~dp0tools;%PATH%"
"%~dp0ECOPV_Mail_Audit_Workbench.exe"
if errorlevel 1 pause
