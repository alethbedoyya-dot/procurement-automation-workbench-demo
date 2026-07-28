@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Procurement Automation Workbench

if not exist "requirements.txt" (
    echo [ERROR] requirements.txt is missing. Keep this launcher in the project folder.
    goto :failed
)
if not exist "run_workbench.py" (
    echo [ERROR] run_workbench.py is missing. Keep all project files together.
    goto :failed
)

call :find_python
if not defined PYTHON_CMD (
    echo [ERROR] Python 3 was not found. Ask the project maintainer for help.
    goto :failed
)

set "VENV_PY=.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo Creating the local Python environment. Please wait...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Could not create the .venv environment.
        goto :failed
    )
    set "NEEDS_DEPENDENCIES=1"
)

if not exist ".venv\.requirements-installed.txt" (
    set "NEEDS_DEPENDENCIES=1"
) else (
    fc /b "requirements.txt" ".venv\.requirements-installed.txt" >nul
    if errorlevel 1 set "NEEDS_DEPENDENCIES=1"
)

if defined NEEDS_DEPENDENCIES (
    echo Installing or updating dependencies. Please keep this window online...
    "%VENV_PY%" -m pip install --upgrade pip
    if errorlevel 1 goto :failed
    "%VENV_PY%" -m pip install -r requirements.txt
    if errorlevel 1 goto :failed
    copy /y "requirements.txt" ".venv\.requirements-installed.txt" >nul
)

if not exist ".venv\.playwright-msedge-ready.txt" (
    echo Preparing the Edge automation component. Please wait...
    "%VENV_PY%" -m playwright install msedge
    if errorlevel 1 goto :failed
    > ".venv\.playwright-msedge-ready.txt" echo Edge automation component is ready.
)

echo Starting the procurement automation workbench...
"%VENV_PY%" "run_workbench.py"
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
    echo.
    echo [ERROR] The workbench stopped unexpectedly. Keep this window open and send a screenshot to the maintainer.
    pause
)
exit /b %EXIT_CODE%

:find_python
py -3 --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py -3"
    exit /b 0
)
python --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    exit /b 0
)
exit /b 1

:failed
echo.
echo Keep this window open and send a screenshot to the maintainer.
pause
exit /b 1
