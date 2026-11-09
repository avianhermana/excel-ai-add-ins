@echo off
title Uninstall Telkomsel Copilot for Excel
color 0C

echo =====================================================================
echo           UNINSTALL TELKOMSEL COPILOT FOR EXCEL
echo =====================================================================
echo.
echo Script ini akan menghapus pendaftaran folder add-in dari Excel
echo dan menghapus shortcut Desktop.
echo.

set /p "CONFIRM=Yakin ingin mencopot Copilot for Excel? (Y/N): "
if /i not "%CONFIRM%"=="Y" (
    echo Dibatalkan.
    exit /b 0
)

:: Hapus Registry TrustedCatalogs
set "CATALOG_GUID={A5F19B8E-1C2D-4E3F-8A9B-0C1D2E3F4A5B}"
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /f >nul 2>nul
echo [OK] Pendaftaran katalog di Excel berhasil dihapus.

:: Hapus Shortcut Desktop
set "SHORTCUT_PATH=%USERPROFILE%\Desktop\Mulai Telkomsel Copilot.lnk"
if exist "%SHORTCUT_PATH%" (
    del /f /q "%SHORTCUT_PATH%" >nul 2>nul
    echo [OK] Shortcut di Desktop berhasil dihapus.
)

echo.
echo Copilot for Excel telah berhasil dicopot dari Excel Anda.
pause
