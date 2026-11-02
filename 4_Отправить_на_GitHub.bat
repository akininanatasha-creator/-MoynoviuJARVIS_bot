@echo off
title Push to GitHub
powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "$env:Path = 'C:\Users\Natali\.local\mingit\cmd;' + $env:Path; Set-Location 'F:\Мой Джарвис'; Write-Host 'Connecting to GitHub...' -ForegroundColor Cyan; git push -u origin main; Write-Host 'Done!' -ForegroundColor Green; Read-Host 'Press Enter to exit'"
