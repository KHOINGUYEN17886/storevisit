@echo off
chcp 65001 > nul
title StoreVisit Pro - Retail Commander Executive Launcher
cd /d "%~dp0"

echo =======================================================
echo    STOREVISIT PRO - AN PHƯỚC RETAIL COMMANDER
echo =======================================================
echo.
echo Đang khởi động hệ thống báo cáo tự động...

if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "run_app.py"
    echo [OK] Đã khởi động StoreVisit Pro ở chế độ giao diện đồ họa GUI.
) else if exist ".venv\Scripts\python.exe" (
    start "" ".venv\Scripts\python.exe" "run_app.py"
    echo [OK] Đã khởi động StoreVisit Pro.
) else (
    echo [LỖI] Không tìm thấy môi trường Python (.venv)!
    echo Vui lòng kiểm tra lại thư mục cài đặt hệ thống.
    pause
)
