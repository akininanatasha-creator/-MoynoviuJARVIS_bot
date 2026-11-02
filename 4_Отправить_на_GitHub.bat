@echo off
chcp 65001 >nul
title Отправка проекта на GitHub
cd /d "%~dp0"

set "PATH=C:\Users\Natali\.local\mingit\cmd;%PATH%"

echo ===================================================================
echo   Отправка проекта Джарвиса на GitHub
echo   Репозиторий: https://github.com/akininanatasha-creator/-MoynoviuJARVIS_bot.git
echo ===================================================================
echo.

git status
echo.
echo Отправляем изменения в ветку main...
echo (Если откроется окно браузера - подтвердите вход в GitHub)
echo.

git push -u origin main

echo.
if %ERRORLEVEL% EQU 0 (
    echo ===================================================================
    echo   [УСПЕХ] Проект успешно сохранен на GitHub!
    echo   Ссылка: https://github.com/akininanatasha-creator/-MoynoviuJARVIS_bot
    echo ===================================================================
) else (
    echo [!] Произошла ошибка при отправке. Проверьте авторизацию в GitHub.
)

echo.
pause
