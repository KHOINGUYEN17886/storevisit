@echo off
title StoreVisit Report Automation Launcher
cd /d "%~dp0"
echo Starting StoreVisit App...
echo Activating Python virtual environment...
if exist ".venv\Scripts\python.exe" (
    start "" ".venv\Scripts\python.exe" "run_app.py"
) else (
    echo Error: Python virtual environment not found in .venv!
    echo Please run installation steps first.
    pause
)
