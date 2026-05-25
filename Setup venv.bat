@echo off
title ChangeGear Setup venv
cd /d "%~dp0"

echo ================================================
echo  ChangeGear venv Setup / Reinstall
echo  Rebuilds .venv from scratch with all dependencies
echo ================================================
echo.

REM 若 venv 不存在 → 建立；若已存在 → 直接用，只裝缺的套件
if not exist ".venv\Scripts\python.exe" (
    echo [1/4] Creating .venv ...
    python -m venv .venv
    if errorlevel 1 (
        echo.
        echo ERROR: python not found. Please install Python 3.11+ first.
        pause
        exit /b 1
    )
) else (
    echo [1/4] .venv already exists, skipping creation
)

echo.
echo [2/4] Bootstrapping pip ...
.venv\Scripts\python.exe -m ensurepip --default-pip
.venv\Scripts\python.exe -m pip install --upgrade pip

echo.
echo [3/4] Installing Python packages from requirements.txt ...
.venv\Scripts\python.exe -m pip install -r requirements.txt

echo.
echo [4/4] Installing Playwright Chromium browser ...
.venv\Scripts\python.exe -m playwright install chromium

echo.
echo ================================================
echo  Setup complete!
echo  You can now run:
echo    - Auto mission start.bat
echo    - Data base build.bat
echo    - Build CMDB DB.bat
echo    - Teach Bot.bat
echo ================================================
echo.
pause
