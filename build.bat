@echo off
REM ============================================================
REM  Italikin Build Script for Windows
REM  Run this to create .exe distribution
REM ============================================================

echo.
echo ========================================
echo   Italikin Build Script
echo ========================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found! Please install Python 3.8+
    pause
    exit /b 1
)

REM Check/Install dependencies
echo [1/4] Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies!
    pause
    exit /b 1
)

REM Install PyInstaller
echo [2/4] Installing PyInstaller...
pip install pyinstaller>=6.0
if errorlevel 1 (
    echo [ERROR] Failed to install PyInstaller!
    pause
    exit /b 1
)

REM Clean previous builds
echo [3/4] Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM Build .exe
echo [4/4] Building .exe...
pyinstaller italikin.spec --clean --noconfirm
if errorlevel 1 (
    echo [ERROR] Build failed!
    pause
    exit /b 1
)

REM Done
echo.
echo ========================================
echo   Build Complete!
echo ========================================
echo.
echo Output: dist\Italikin\Italikin.exe
echo.
pause
