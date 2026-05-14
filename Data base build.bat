@echo off
title ChangeGear DB Build

cd /d "%~dp0"

echo ================================================
echo  ChangeGear History DB Builder
echo  Scrapes All Incidents into SQLite DB
echo  Press Ctrl+C to stop (existing data is safe)
echo ================================================
echo.

python build_history_db.py

echo.
echo Done. Press any key to close...
pause >nul
