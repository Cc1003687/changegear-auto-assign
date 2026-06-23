@echo off
title ChangeGear Backfill Wen
cd /d "%~dp0"

echo ================================================
echo  ChangeGear "Wen Blessed" Backfill
echo  Scrapes each ticket's audit log to find the
echo  last person who did Save/Accept. Marks records
echo  where wen.hsieh was the last saver as authority-
echo  blessed, so they get x2 weight in future matches.
echo  Press Ctrl+C anytime; already-saved progress is kept.
echo ================================================
echo.

.venv\Scripts\python.exe backfill_wen.py %*

echo.
pause
