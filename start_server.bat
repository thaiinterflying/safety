@echo off
title TIF Safety Form Server
cd /d "%~dp0"
echo Starting TIF Safety Form Server...
start "" http://localhost:3000/index.html
node server.js
pause
