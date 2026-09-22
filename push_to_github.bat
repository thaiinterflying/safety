@echo off
title Push to GitHub
cd /d "%~dp0"
echo ======================================================
echo Pushing TIF Safety to https://github.com/thaiinterflying/safety.git
echo ======================================================
git push -u origin main
echo.
echo ======================================================
pause
