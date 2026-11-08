@echo off
chcp 65001 > nul
title DUNG HE THONG PCGD 2026
cd /d "%~dp0"

echo =========================================================================================
echo    DUNG HE THONG TONG HOP SO LIEU PHO CAP GIAO DUC VA XOA MU CHU 2026
echo =========================================================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\stop_service.ps1" -Port 8501 -ProjectDir "%~dp0"

echo.
timeout /t 3 /nobreak > nul
exit /b 0
