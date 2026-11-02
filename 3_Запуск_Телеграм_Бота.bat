@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
title OpenJarvis - Telegram Bot (@MoynoviuJARVIS_bot)

cd /d "%~dp0"

if "%NIM_API_KEY%"=="" (
    if exist api_key.txt (
        for /f "usebackq delims=" %%k in ("api_key.txt") do set "NIM_API_KEY=%%k"
    )
)

echo ===================================================================
echo   Запуск Telegram-бота Джарвиса (@MoynoviuJARVIS_bot)
echo   Модель: nvidia/nemotron-3.5-bolt-e4b (NVIDIA NIM)
echo ===================================================================
echo.
"OpenJarvis\.venv\Scripts\python.exe" telegram_bot.py
pause
