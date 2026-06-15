@echo off
title ChangeGear Reflect Now
cd /d "%~dp0"

echo ================================================
echo  ChangeGear Reflection - Manual Trigger
echo  Synthesize dispatch principles from recent
echo  feedback + auto-detected corrections
echo ================================================
echo.

.venv\Scripts\python.exe reflect.py %*

echo.
pause
