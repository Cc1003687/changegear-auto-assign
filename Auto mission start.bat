@echo off
title ChangeGear Auto Assign

cd /d "%~dp0"

echo ================================================
echo  ChangeGear Auto Assign v6
echo  Press Ctrl+C to stop at any time
echo ================================================
echo.

python changegear_auto_assign_v6.py

echo.
echo Done. Press any key to close...
pause >nul
