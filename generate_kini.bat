@echo off
chcp 65001 >nul
cd /d "%~dp0"
python generate_kini.py --config kini_config.json
if errorlevel 1 echo Не удалось сформировать расчёт. Смотрите сообщение выше.
pause
