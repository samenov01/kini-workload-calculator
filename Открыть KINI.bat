@echo off
chcp 65001 >nul
cd /d "%~dp0"
python webapp.py --open
if errorlevel 1 pause
