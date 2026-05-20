@echo off
title ChangeGear Teach Bot
cd /d "%~dp0"

echo ================================================
echo  ChangeGear Teach Bot - Feedback Recorder
echo  Tell the bot which ticket should be assigned to whom
echo ================================================
echo.

python teach.py %*

echo.
pause
