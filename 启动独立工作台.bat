@echo off
setlocal EnableExtensions

set "APP_DIR=%~dp0"
pushd "%APP_DIR%" || (
  echo 无法打开工作台目录。
  pause
  exit /b 1
)

rem 先使用当前用户安装的 Python，再回退到 PATH 中的 python/py。
rem 不使用同一括号块中的 %ERRORLEVEL%，避免其在解析时提前展开。
set "PYTHON_EXE="
if exist "%LocalAppData%\Programs\Python\Python312\python.exe" set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python312\python.exe"
if not defined PYTHON_EXE if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python311\python.exe"

if not defined PYTHON_EXE (
  for /f "delims=" %%P in ('where python 2^>nul') do if not defined PYTHON_EXE set "PYTHON_EXE=%%P"
)

if defined PYTHON_EXE goto run_python

where py >nul 2>nul
if not errorlevel 1 goto run_py_launcher

echo 未找到 Python 3，请先安装 Python 3.10 或更高版本。
pause
popd
exit /b 1

:run_python
"%PYTHON_EXE%" "%APP_DIR%workbench_launcher.py" %*
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%

:run_py_launcher
py -3 "%APP_DIR%workbench_launcher.py" %*
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
