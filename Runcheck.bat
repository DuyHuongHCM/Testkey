@echo off
:: Thiet lap code page UTF-8 de tranh loi font tieng Viet trong CMD
chcp 65001 >nul
title Trinh chay Kiem tra Ban quyen - Thu ky AI

:: Thiet lap thu muc lam viec hien tai ve dung thu muc chua file bat
cd /d "%~dp0"

:: Kiem tra quyen Admin bang lenh he thong fsutil
fsutil dirty query %systemdrive% >nul 2>&1
if %errorLevel% == 0 (
    goto :runScript
)

:: Khoi tao bien de tim tệp tin PowerShell phu hop
set "ps_file="
if exist "Check_Key.ps1" set "ps_file=Check_Key.ps1"
if not defined ps_file if exist "CheckLicensing.ps1" set "ps_file=CheckLicensing.ps1"
if not defined ps_file if exist "Kiem tra ban quyen windows_v_office.ps1" set "ps_file=Kiem tra ban quyen windows_v_office.ps1"
if not defined ps_file (
    for %%f in (*.ps1) do (
        set "ps_file=%%f"
    )
)

:: Neu khong tim thay file .ps1 thi thong bao loi va dung lai
if "%ps_file%"=="" (
    echo [LOI] Khong tim thay file kich ban PowerShell (.ps1) nao trong thu muc nay!
    echo Vui long de file "RunCheck.bat" nay chung mot thu muc voi file script PowerShell cua anh.
    echo.
    pause
    exit /b
)

echo ==========================================================
echo  DANG TU DONG YEU CAU QUYEN ADMINISTRATOR...
echo ==========================================================

:: Nang quyen truc tiep len PowerShell Admin va thuc thi file .ps1 (Boc nhay kep tranh loi khoang trang)
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dp0%ps_file%""' -Verb RunAs"
exit /b

:runScript
:: Neu nguoi dung da mo san CMD voi quyen Admin tu truoc thi chay luon o day
set "ps_file="
if exist "Check_Key.ps1" set "ps_file=Check_Key.ps1"
if not defined ps_file if exist "CheckLicensing.ps1" set "ps_file=CheckLicensing.ps1"
if not defined ps_file if exist "Kiem tra ban quyen windows_v_office.ps1" set "ps_file=Kiem tra ban quyen windows_v_office.ps1"
if not defined ps_file (
    for %%f in (*.ps1) do (
        set "ps_file=%%f"
    )
)

if "%ps_file%"=="" (
    echo [LOI] Khong tim thay file kich ban PowerShell (.ps1) nao!
    pause
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%ps_file%"
pause