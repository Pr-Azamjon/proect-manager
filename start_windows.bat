@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .env copy .env.example .env
python -m pip install -r requirements.txt
if errorlevel 1 goto error
python -m flask --app app init-db
if errorlevel 1 goto error
python app.py
pause
exit /b 0
:error
echo Ошибка установки или инициализации. Проверьте Python и интернет-соединение.
pause
exit /b 1
