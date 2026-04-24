@echo off
echo ========================================
echo Запуск инвестиционного калькулятора
echo ========================================
echo.

cd /d "C:\Users\Viktoria\PycharmProjects\PythonProject1"

echo Активация виртуального окружения...
call .venv\Scripts\activate.bat

echo Запуск приложения...
streamlit run main.py

pause