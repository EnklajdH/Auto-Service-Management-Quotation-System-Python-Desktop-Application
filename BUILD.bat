@echo off
echo ============================================
echo   DUKE NDERTUAR PREVENTIV INSTALLER...
echo ============================================
echo.

:: Hapi 1 - PyInstaller
echo [1/2] Duke krijuar Preventiv.exe me PyInstaller...
cd /d "C:\Users\enkla\OneDrive\Desktop\PREVENTIVAPP"
pyinstaller --onefile --windowed --name Preventiv main.py

if not exist "dist\Preventiv.exe" (
    echo.
    echo [GABIM] Preventiv.exe nuk u krijua!
    echo Sigurohu qe PyInstaller eshte instaluar: pip install pyinstaller
    pause
    exit /b 1
)

echo [OK] Preventiv.exe u krijua me sukses!
echo.

:: Hapi 2 - Inno Setup
echo [2/2] Duke krijuar Installer me Inno Setup...

set INNO="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

if not exist %INNO% (
    echo [GABIM] Inno Setup nuk u gjet!
    echo Shkarko nga: https://jrsoftware.org/isdl.php
    pause
    exit /b 1
)

%INNO% "Preventiv_Installer.iss"

echo.
echo ============================================
echo   GATI! Installer ndodhet te:
echo   Output\Preventiv_Setup.exe
echo ============================================
echo.
pause
