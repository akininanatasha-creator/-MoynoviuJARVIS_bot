@echo off
chcp 65001 >nul
title OpenJarvis - Веб-интерфейс

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

echo Открываем браузер...
start http://127.0.0.1:8000

cd /d "%~dp0OpenJarvis"
echo Сервер OpenJarvis запускается...
.venv\Scripts\python.exe -m openjarvis.cli serve --host 127.0.0.1 --port 8000
pause