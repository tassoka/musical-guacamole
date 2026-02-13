@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo --- Запуск приложения Склад ---

if not exist .venv (
    echo [1/3] Создание виртуального окружения...
    python -m venv .venv
) else (
    echo [1/3] Виртуальное окружение найдено.
)

echo [2/3] Активация и проверка зависимостей...
call .venv\Scripts\activate
if exist requirements.txt (
    pip install -r requirements.txt
) else (
    echo Внимание: requirements.txt не найден.
)

echo [3/3] Запуск сервера...
python app.py

pause