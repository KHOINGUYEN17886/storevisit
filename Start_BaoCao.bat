@echo off
title StoreVisit Pro - Retail Commander
cd /d "%~dp0"

echo =======================================================
echo    STOREVISIT PRO - AN PHUOC RETAIL COMMANDER
echo =======================================================
echo.
echo Dang khoi dong StoreVisit Pro...

if not exist ".venv\Scripts\python.exe" goto NO_PYTHON

if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "run_app.py"
    goto DONE
)

start "" ".venv\Scripts\python.exe" "run_app.py"
goto DONE

:NO_PYTHON
echo [ERROR] Khong tim thay moi truong Python (.venv)!
echo Vui long kiem tra lai thu muc cai dat.
pause

:DONE
