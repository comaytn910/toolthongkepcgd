@echo off
chcp 65001 > nul
title TẠO PHÍM TẮT MÀN HÌNH CHÍNH (DESKTOP)
cd /d "%~dp0"

echo =========================================================================================
echo    TẠO LỐI TẮT (SHORTCUT) TRÊN MÀN HÌNH DESKTOP - HỆ THỐNG PCGD 2026
echo =========================================================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\create_shortcut.ps1"

echo.
echo =========================================================================================
echo  Bây giờ Người đẹp có thể ra ngoài màn hình Desktop và click đúp vào biểu tượng:
echo  "HE THONG PCGD 2026 - THAI NGUYEN" để vào ngay ứng dụng!
echo =========================================================================================
echo.
pause
