@echo off
REM  Double-click this to start the rehearsal bridge.
REM  It opens the deck in PowerPoint, starts the slide show, and lets the
REM  rehearsal page drive it. Keep this window open while you practise.

cd /d "%~dp0"
title Speech rehearsal bridge
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0rehearsal-bridge.ps1" %*

echo.
echo   The bridge has stopped.
pause
