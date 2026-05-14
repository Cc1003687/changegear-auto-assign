@echo off
title ChangeGear CMDB DB Build

cd /d "%~dp0"

echo ================================================
echo  ChangeGear CMDB Owner DB Builder
echo  Scrapes All Managed Items into SQLite DB
echo  Press Ctrl+C to stop (existing data is safe)
echo ================================================
echo.

python build_cmdb_db.py

echo.
echo Done. Press any key to close...
pause >nul
