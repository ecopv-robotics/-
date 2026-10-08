@echo off
chcp 65001 >nul
setlocal EnableExtensions
call "%~dp0启动网页版.bat" --open-review %*
exit /b %ERRORLEVEL%
