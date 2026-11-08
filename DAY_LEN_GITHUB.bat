@echo off
chcp 65001 >nul
title DAY DU AN LEN GITHUB - PCGD 2026
cd /d "%~dp0"

echo =========================================================================
echo   HE THONG TONG HOP SO LIEU PCGD - XMC 2026 TINH THAI NGUYEN
echo                   DAY DU AN LEN GITHUB
echo =========================================================================
echo.
echo Kho luu tru: https://github.com/comaytn910/toolthongkepcgd.git
echo Nhanh day:   main
echo.
echo LUU Y:
echo Neu day la lan dau tien day len GitHub tu may nay, mot cua so trinh duyet
echo hoac hop thoai dang nhap GitHub se hien len.
echo Nguoi dep chi can bam "Sign in with your browser" de xac thuc.
echo.
echo Dang tien hanh dong goi va day du lieu (khoang 120MB, mat 30s - 1 phut tuy mang)...
echo.

git push -u origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo =========================================================================
    echo   CHUC MUNG! TOAN BO DU AN DA DUOC DAY LEN GITHUB THANH CONG 100%!
    echo =========================================================================
    echo.
    echo Sang may khac, Nguoi dep chi can:
    echo 1. git clone https://github.com/comaytn910/toolthongkepcgd.git
    echo 2. Chay tep CAI_DAT_BAN_DAU.bat la dung duoc ngay.
    echo.
) else (
    echo.
    echo =========================================================================
    echo   [!] CO LOI TRONG QUA TRINH DAY LEN GITHUB
    echo =========================================================================
    echo Vui long kiem tra:
    echo - Ket noi mang Internet.
    echo - Quyen truy cap cua tai khoan GitHub comaytn910.
    echo.
)

pause
