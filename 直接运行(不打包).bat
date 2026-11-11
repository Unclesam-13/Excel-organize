@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist ".venv\Scripts\pythonw.exe" goto :run
echo First run: installing packages...
where py >nul 2>nul
if %errorlevel%==0 (py -3 -m venv .venv) else (python -m venv .venv)
if not exist ".venv\Scripts\python.exe" (echo Python not found & pause & exit /b 1)
".venv\Scripts\python.exe" -m pip install -r requirements.txt pywebview
if errorlevel 1 (echo Install failed & pause & exit /b 1)
:run
start "" ".venv\Scripts\pythonw.exe" desktop.py
