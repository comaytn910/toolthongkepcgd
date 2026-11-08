@echo off
chcp 65001 > nul
title THIET LAP BAN DAU HE THONG PCGD 2026 - SO GDDT THAI NGUYEN
cd /d "%~dp0"

echo =========================================================================================
echo    THIẾT LẬP BAN ĐẦU HỆ THỐNG TỔNG HỢP PCGD ^& XMC 2026 - SỞ GDĐT THÁI NGUYÊN
echo =========================================================================================
echo.

set "SYS_PYTHON="

:: Kiem tra cac duong dan python he thong
if exist "C:\Users\ADMIN\AppData\Local\Programs\Python\Python313\python.exe" (
    set "SYS_PYTHON=C:\Users\ADMIN\AppData\Local\Programs\Python\Python313\python.exe"
    goto RUN_SETUP
)

where python >nul 2>nul
if %ERRORLEVEL% equ 0 (
    set "SYS_PYTHON=python"
    goto RUN_SETUP
)

where py >nul 2>nul
if %ERRORLEVEL% equ 0 (
    set "SYS_PYTHON=py -3"
    goto RUN_SETUP
)

echo [LỖI] Máy tính chưa được cài đặt Python (yêu cầu bản 3.10 đến 3.13)!
echo.
echo HƯỚNG DẪN XỬ LÝ:
echo 1. Truy cập https://www.python.org/downloads/ để tải bản Python 3.13.
echo 2. Khi chạy cài đặt, BẮT BUỘC ĐÁNH DẤU CHỌN: "Add python.exe to PATH".
echo 3. Cài đặt xong, chạy lại tệp này (CAI_DAT_BAN_DAU.bat).
echo.
pause
exit /b 1

:RUN_SETUP
"%SYS_PYTHON%" scripts/setup_new_machine.py

if %ERRORLEVEL% neq 0 (
    echo.
    echo [!] Quá trình thiết lập gặp lỗi. Vui lòng xem tệp "setup_error.log" để biết chi tiết.
    echo.
    pause
    exit /b %ERRORLEVEL%
)

echo.
pause
exit /b 0
