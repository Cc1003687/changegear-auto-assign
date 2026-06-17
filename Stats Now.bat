@echo off
title ChangeGear Stats Now
cd /d "%~dp0"

echo ================================================
echo  ChangeGear Dispatch Accuracy Report
echo  Overall + field-level + source breakdown + trend
echo ================================================
echo.

.venv\Scripts\python.exe stats.py %*

echo.
pause
