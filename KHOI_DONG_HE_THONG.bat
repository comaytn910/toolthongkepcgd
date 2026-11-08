@echo off
chcp 65001 > nul
title HE THONG TONG HOP PCGD VA XMC 2026 - SO GDDT THAI NGUYEN
cd /d "%~dp0"

echo =========================================================================================
echo    HE THONG TONG HOP SO LIEU PHO CAP GIAO DUC VA XOA MU CHU 2026 - SO GDDT THAI NGUYEN
echo =========================================================================================
echo.

REM 1. Xac dinh trinh thong dich Python
set "PYTHON_EXE="

if exist "%~dp0.venv\Scripts\python.exe" (
    set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
    goto PYTHON_FOUND
)

if exist "C:\Users\ADMIN\AppData\Local\Programs\Python\Python313\python.exe" (
    set "PYTHON_EXE=C:\Users\ADMIN\AppData\Local\Programs\Python\Python313\python.exe"
    goto PYTHON_FOUND
)

where python >nul 2>nul
if %ERRORLEVEL% equ 0 (
    set "PYTHON_EXE=python"
    goto PYTHON_FOUND
)

echo [LOI] Khong tim thay moi truong Python tren may tinh!
echo Vui long chay tep "CAI_DAT_BAN_DAU.bat" hoac cai dat Python 3.13.
echo.
pause
exit /b 1

:PYTHON_FOUND
echo [*] Dang su dung moi truong: %PYTHON_EXE%

REM 2. Kiem tra xem moi truong da co du thu vien chua (kiem tra tuc thi < 0.05s)
"%PYTHON_EXE%" -c "import importlib.util as u, sys; sys.exit(0 if all(u.find_spec(m) for m in ['streamlit', 'openpyxl', 'pandas']) else 1)" >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo.
    echo [CANH BAO] Moi truong Python chua cai dat du thu vien can thiet!
    echo Vui long chay tep "CAI_DAT_BAN_DAU.bat" de thiet lap tu dong.
    echo.
    pause
    exit /b 1
)

REM 3. Kiem tra chinh xac trang thai cong 8501 qua PowerShell
set "PORT_RES="
for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\check_port_service.ps1" -Port 8501 -ProjectDir "%~dp0"`) do (
    set "PORT_RES=%%i"
)

echo %PORT_RES% | findstr "STATUS=PCGD_RUNNING" > nul
if %ERRORLEVEL% equ 0 (
    echo [*] May chu PCGD cua du an dang chay san sang tren cong 8501.
    goto OPEN_BROWSER
)

echo %PORT_RES% | findstr "STATUS=PORT_CONFLICT" > nul
if %ERRORLEVEL% equ 0 (
    echo.
    echo =========================================================================================
    echo [LOI XUNG DOT CONG] Cong 8501 dang bi chiem boi mot ung dung KHAC!
    echo Chi tiet: %PORT_RES%
    echo.
    echo He thong tu choi chay de de bao ve an toan cho ung dung dang hoat dong.
    echo Vui long kiem tra hoac tat ung dung tren roi khoi dong lai.
    echo =========================================================================================
    echo.
    pause
    exit /b 2
)

REM Truong hop STATUS=NOT_RUNNING -> Khoi dong moi
echo [*] Dang khoi dong may chu Web truc tuyen...
start /b "" "%PYTHON_EXE%" -m streamlit run ui/app.py --server.port 8501 --server.headless true > nul 2>&1
timeout /t 3 /nobreak > nul

:OPEN_BROWSER
echo [*] Dang mo giao dien Web tren trinh duyet: http://localhost:8501
start http://localhost:8501

echo.
echo =========================================================================================
echo  HE THONG DA SAN SANG PHUC VU TREN TRINH DUYET!
echo.
echo  - Thu muc du lieu nguon (92 xa): data\
echo  - Thu muc ket qua xuat ra:      outputs\
echo  - So cai tong hop 33 bieu:       outputs\SO_CAI_TONG_HOP_PCGD_2026_THAI_NGUYEN.xlsx
echo  - Bao cao doi soat 63 quy tac:   outputs\BAO_CAO_DOI_SOAT_63_QUY_TAC_PCGD_2026.xlsx
echo =========================================================================================
echo.

:MENU
echo -----------------------------------------------------------------------------------------
echo BANG DIEU KHIEN NHANH:
echo   [1] Mo lai trang Web tren trinh duyet (http://localhost:8501)
echo   [2] Mo thu muc Du lieu nguon (data)
echo   [3] Mo thu muc Tep ket qua xuat ra (outputs)
echo   [4] Chay kiem tra doi soat nhanh toan bo 92 xa
echo   [5] Dung tat hoan toan may chu he thong
echo   [0] Dong cua so nay (May chu Web van tiep tuc chay ngam)
echo -----------------------------------------------------------------------------------------
set "opt="
set /p "opt=Nhap lua chon (0-5) roi bam Enter: "

if "%opt%"=="1" goto ACT_BROWSER
if "%opt%"=="2" goto ACT_DATA
if "%opt%"=="3" goto ACT_OUTPUTS
if "%opt%"=="4" goto ACT_CHECK
if "%opt%"=="5" goto ACT_STOP
if "%opt%"=="0" goto ACT_EXIT
if not defined opt goto MENU

echo [!] Lua chon khong hop le, vui long thu lai.
goto MENU

:ACT_BROWSER
start http://localhost:8501
goto MENU

:ACT_DATA
explorer "%~dp0data"
goto MENU

:ACT_OUTPUTS
explorer "%~dp0outputs"
goto MENU

:ACT_CHECK
echo.
echo [*] Dang chay kiem tra doi soat 92 xa...
"%PYTHON_EXE%" scripts/run_all_ready_sheets.py
echo.
echo [*] Hoan thanh doi soat.
goto MENU

:ACT_STOP
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\stop_service.ps1" -Port 8501 -ProjectDir "%~dp0"
echo.
timeout /t 2 /nobreak > nul
exit /b 0

:ACT_EXIT
exit /b 0
