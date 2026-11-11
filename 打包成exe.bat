@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==== Excel Library : build Windows EXE ====
if exist ".buildenv\Scripts\python.exe" goto :have_env
where py >nul 2>nul
if %errorlevel%==0 (py -3 -m venv .buildenv) else (python -m venv .buildenv)
:have_env
if not exist ".buildenv\Scripts\python.exe" goto :no_python
echo [1/2] Installing build tools (first time takes a few minutes)...
".buildenv\Scripts\python.exe" -m pip install --upgrade pip -q
".buildenv\Scripts\python.exe" -m pip install -r requirements-build.txt
if errorlevel 1 goto :fail
echo [2/2] Building...
set PYTHONUTF8=1
".buildenv\Scripts\python.exe" build.py
if errorlevel 1 goto :fail
explorer "%~dp0dist"
pause
exit /b 0
:no_python
echo ERROR: Python not found. Install Python 3.12 from https://www.python.org/downloads/
echo        and tick "Add python.exe to PATH".
pause
exit /b 1
:fail
echo BUILD FAILED - see messages above.
pause
exit /b 1
