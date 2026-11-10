@echo off
setlocal enabledelayedexpansion
title Setup Excel AI Add-ins
color 0A

echo =====================================================================
echo       INSTALLER EXCEL AI ADD-INS (1-CLICK SETUP)
echo =====================================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

:: 1. Cek Node.js
echo [1/5] Memeriksa instalasi Node.js...
where node >nul 2>nul
if %ERRORLEVEL% equ 0 (
    for /f "tokens=*" %%v in ('node -v') do set "NODE_VER=%%v"
    echo   [OK] Node.js terdeteksi: !NODE_VER!
) else (
    echo   [PERINGATAN] Node.js belum terinstall di komputer ini.
    echo.
    echo   Apakah Anda ingin menginstal Node.js sekarang via Windows Package Manager?
    set /p "INSTALL_NODE=Ketik Y untuk install otomatis, atau N untuk buka web download (Y/N): "
    if /i "!INSTALL_NODE!"=="Y" (
        echo   Mengunduh dan memasang Node.js LTS...
        winget install OpenJS.NodeJS.LTS -e --silent --accept-package-agreements --accept-source-agreements
        if !ERRORLEVEL! neq 0 (
            echo   Gagal install via winget. Membuka halaman resmi nodejs.org...
            start https://nodejs.org/en/download
        )
    ) else (
        start https://nodejs.org/en/download
    )
    echo   Silakan selesaikan instalasi Node.js terlebih dahulu, lalu jalankan installer ini kembali.
    pause
    exit /b 1
)

:: 2. Instalasi Dependensi Node.js
echo.
echo [2/5] Memeriksa dependensi npm (node_modules)...
if not exist "%SCRIPT_DIR%\node_modules" (
    echo   Mengunduh paket dependensi...
    cd /d "%SCRIPT_DIR%"
    call npm install
    if !ERRORLEVEL! neq 0 (
        echo   [ERROR] Gagal menjalankan npm install.
        pause
        exit /b 1
    )
    echo   [OK] Dependensi npm berhasil dipasang.
) else (
    echo   [OK] Dependensi npm sudah terpasang.
)

:: 3. Generate & Daftarkan Sertifikat SSL Localhost ke Windows Trust Store
echo.
echo [3/5] Memeriksa & Mendaftarkan Sertifikat SSL Localhost...
cd /d "%SCRIPT_DIR%"
if not exist "%SCRIPT_DIR%\certs\cert.pem" (
    node generate-cert.js
)

if exist "%SCRIPT_DIR%\certs\cert.pem" (
    certutil -addstore -user "Root" "%SCRIPT_DIR%\certs\cert.pem" >nul 2>nul
    echo   [OK] Sertifikat SSL berhasil didaftarkan ke Windows Trust Store.
) else (
    echo   [INFO] Sertifikat belum siap, akan dibuat otomatis saat server dijalankan.
)

:: 4. Konfigurasi WebView2 Excel agar tidak memblokir localhost SSL
echo.
echo [4/5] Mengatur izin WebView2 Excel untuk Localhost Dev...
setx WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS "--ignore-certificate-errors" >nul 2>nul
echo   [OK] Izin WebView2 berhasil dikonfigurasi.

:: 5. Daftarkan Folder Add-in ke Katalog Resmi Excel (Trusted Catalogs)
echo.
echo [5/5] Mendaftarkan Manifest ke Excel (Trusted Shared Folder)...
set "CATALOG_GUID={A5F19B8E-1C2D-4E3F-8A9B-0C1D2E3F4A5B}"
set "REG_KEY=HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%"

reg add "%REG_KEY%" /v "Id" /t REG_SZ /d "%CATALOG_GUID%" /f >nul 2>nul
reg add "%REG_KEY%" /v "Url" /t REG_SZ /d "%SCRIPT_DIR%" /f >nul 2>nul
reg add "%REG_KEY%" /v "Flags" /t REG_DWORD /d 1 /f >nul 2>nul

echo   [OK] Folder add-in berhasil didaftarkan ke katalog Excel!

:: 6. Buat Shortcut di Desktop
echo.
echo Membuat shortcut di Desktop...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([System.IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), 'Mulai Excel AI Add-ins.lnk')); $s.TargetPath = [System.IO.Path]::Combine('%SCRIPT_DIR%', '2-Jalankan-Server.bat'); $s.WorkingDirectory = '%SCRIPT_DIR%'; $s.Save()" >nul 2>nul
echo   [OK] Shortcut 'Mulai Excel AI Add-ins' berhasil dibuat di Desktop Anda.

echo.
echo =====================================================================
echo              INSTALASI BERHASIL! (SELESAI 100%%)
echo =====================================================================
echo.
echo LANGKAH SELANJUTNYA:
echo 1. Jalankan server dengan klik 2x file '2-Jalankan-Server.bat' 
echo    (atau melalui shortcut di Desktop).
echo 2. Buka Excel Desktop:
echo    - Klik menu Insert (Sisipkan) ^> My Add-ins (Add-in Saya).
echo    - Pilih tab 'SHARED FOLDER' (Folder Bersama).
echo    - Klik 'Excel AI Add-ins' ^> Add (Tambah).
echo 3. Selesai! Panel Excel AI Add-ins akan langsung muncul di samping Excel.
echo.
echo (Untuk Excel Web / Office 365 di browser:
echo  Masuk menu Insert ^> Add-ins ^> Manage My Add-ins ^> Upload My Add-in
echo  lalu pilih file 'manifest.xml' di folder ini).
echo =====================================================================
echo.
pause
