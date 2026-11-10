@echo off
setlocal enabledelayedexpansion
title Excel AI Add-ins - Backend Server
color 0B

echo =====================================================================
echo             EXCEL AI ADD-INS - BACKEND SERVER
echo =====================================================================
echo.

cd /d "%~dp0"

:: Cek Node.js
where node >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Node.js tidak ditemukan!
    echo Silakan jalankan '1-Install-Addin.bat' terlebih dahulu untuk setup.
    echo.
    pause
    exit /b 1
)

:: Cek apakah port 3000 sedang berjalan
netstat -ano | findstr :3000 | findstr LISTENING >nul 2>nul
if %ERRORLEVEL% equ 0 (
    echo [INFO] Server Excel AI Add-ins tampaknya sudah berjalan di port 3000.
    echo.
)

echo Menjalankan server HTTPS Excel AI Add-ins di port 3000...
echo.
echo Pastikan jendela ini TETAP TERBUKA selama Anda menggunakan Excel AI Add-ins.
echo (Tekan Ctrl+C untuk menghentikan server jika sudah selesai).
echo.

node server.js
pause
