@echo off
chcp 65001 >nul
title OpenJarvis - Консольный чат

cd /d "%~dp0"

if "%NIM_API_KEY%"=="" (
    if exist api_key.txt (
        for /f "usebackq delims=" %%k in ("api_key.txt") do set "NIM_API_KEY=%%k"
    )
)

if "%NIM_API_KEY%"=="" (
    echo ===================================================================
    echo  NVIDIA NIM API ключ не задан!
    echo  Бесплатный ключ можно взять на: https://build.nvidia.com
    echo ===================================================================
    echo.
    set /p "USER_KEY=Введите ваш API ключ (начинается с nvapi-...): "
    call set "NIM_API_KEY=%%USER_KEY%%"
    > api_key.txt echo %%USER_KEY%%
)

cd /d "%~dp0OpenJarvis"
.venv\Scripts\python.exe -m openjarvis.cli chat
pause