@echo off
title PdfToWord Build Tool

echo ========================================
echo   PdfToWord Build Script
echo ========================================
echo.

cd /d "%~dp0"

echo [1/5] Checking virtual environment...
if not exist "venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo Failed to create virtual environment!
        pause
        exit /b 1
    )
    echo Virtual environment created!
) else (
    echo Virtual environment exists, skip creation
)
echo.

echo [2/5] Installing dependencies...
call venv\Scripts\pip.exe install -r requirements.txt --quiet
if errorlevel 1 (
    echo Failed to install dependencies!
    pause
    exit /b 1
)
call venv\Scripts\pip.exe install "pyinstaller>=6.0,<7.0" --quiet
if errorlevel 1 (
    echo Failed to install pyinstaller!
    pause
    exit /b 1
)
echo Dependencies installed!
echo.

echo [3/5] Generating icon...
call venv\Scripts\python.exe generate_icon.py
if errorlevel 1 (
    echo Failed to generate icon!
    pause
    exit /b 1
)
echo.

echo [4/5] Cleaning old build...
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
echo Clean completed!
echo.

echo [5/5] Building exe...
call venv\Scripts\python.exe -m PyInstaller --noconfirm --clean PdfToWord.spec

if errorlevel 1 (
    echo.
    echo ========================================
    echo   Build FAILED! Check error messages
    echo ========================================
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Build SUCCESS!
echo ========================================
echo.
echo Output: dist\PdfToWord.exe
echo.

if exist "dist\PdfToWord.exe" (
    for %%A in ("dist\PdfToWord.exe") do echo Size: %%~zA bytes
)

echo.
echo [+] Copying 说明.txt to dist...
copy /y "说明.txt" "dist\说明.txt" >nul 2>nul
if not exist "dist\说明.txt" (
    echo Warning: 说明.txt not found in project root, skip copy
) else (
    echo 说明.txt copied to dist\
)

echo.
echo ========================================
echo   Press any key to exit...
echo ========================================
pause >nul
