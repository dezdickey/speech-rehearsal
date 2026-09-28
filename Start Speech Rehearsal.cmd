@echo off
REM  Start here: double-click this to start Speech Rehearsal with your slides.
REM  It uses the presentation you have open in PowerPoint or LibreOffice Impress,
REM  or asks which one to open, then opens the app already connected to it.
REM  Keep this window open while you practise. (No slides? Just open
REM  speech-rehearsal.html instead.)

cd /d "%~dp0"
title Speech rehearsal bridge
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0rehearsal-bridge.ps1" %*

echo.
echo   The bridge has stopped.
pause
