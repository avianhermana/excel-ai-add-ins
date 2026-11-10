@echo off
title Uninstall Excel AI Add-ins
color 0C

echo =====================================================================
echo                 UNINSTALL EXCEL AI ADD-INS
echo =====================================================================
echo.
echo Script ini akan menghapus pendaftaran folder add-in dari Excel
echo dan menghapus shortcut Desktop.
echo.

set /p "CONFIRM=Yakin ingin mencopot Excel AI Add-ins? (Y/N): "
if /i not "%CONFIRM%"=="Y" (
    echo Dibatalkan.
    exit /b 0
)

:: Hapus Registry TrustedCatalogs
set "CATALOG_GUID={A5F19B8E-1C2D-4E3F-8A9B-0C1D2E3F4A5B}"
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /f >nul 2>nul
echo [OK] Pendaftaran katalog di Excel berhasil dihapus.

:: Hapus Shortcut Desktop
set "SHORTCUT_PATH=%USERPROFILE%\Desktop\Mulai Excel AI Add-ins.lnk"
if exist "%SHORTCUT_PATH%" (
    del /f /q "%SHORTCUT_PATH%" >nul 2>nul
    echo [OK] Shortcut di Desktop berhasil dihapus.
)
set "OLD_SHORTCUT_PATH=%USERPROFILE%\Desktop\Mulai Telkomsel Copilot.lnk"
if exist "%OLD_SHORTCUT_PATH%" (
    del /f /q "%OLD_SHORTCUT_PATH%" >nul 2>nul
)

echo.
echo Excel AI Add-ins telah berhasil dicopot dari Excel Anda.
pause
