@echo off
chcp 65001 >nul
cd /d "%~dp0"
python "app\control.py" stop
timeout /t 3 >nul
