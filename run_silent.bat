@echo off
:: Запуск программы без окна консоли (тихий режим)

:: Обновляем библиотеки в фоновом режиме
python -m pip install pandas openpyxl --quiet --upgrade >nul 2>&1

:: Запускаем программу без консоли
start "" pythonw materials.py