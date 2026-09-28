@echo off
REM  Windows tags files downloaded from the internet, and PowerShell refuses to run
REM  them until that tag is cleared. Put this in the folder with the project files
REM  and double-click it once. It only touches this folder, and it only removes the
REM  download tag - nothing is modified, moved, or deleted.

cd /d "%~dp0"
title Unblock rehearsal files
echo.
echo   Folder: %~dp0
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "$n=0; Get-ChildItem -LiteralPath '%~dp0' -File | Where-Object { $_.Extension -in '.ps1','.cmd','.html','.pptx','.pdf' } | ForEach-Object { if (Get-Item -LiteralPath $_.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue) { Unblock-File -LiteralPath $_.FullName; Write-Host ('   unblocked   ' + $_.Name) -ForegroundColor Green; $n++ } else { Write-Host ('   already ok  ' + $_.Name) -ForegroundColor DarkGray } }; Write-Host ''; if ($n) { Write-Host ('   Cleared ' + $n + ' file(s). You can start the bridge now.') -ForegroundColor Cyan } else { Write-Host '   Nothing needed clearing. You are good to go.' -ForegroundColor Cyan }"

echo.
pause
