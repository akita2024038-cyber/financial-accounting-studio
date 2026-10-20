@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .env (
 echo Create a .env file from .env.example and configure APP_SECRET and SETUP_TOKEN.
 pause
 exit /b 1
)
if not exist .venv\Scripts\python.exe (
 py -3 -m venv .venv
 .venv\Scripts\python.exe -m pip install -r requirements.txt
)
.venv\Scripts\python.exe run.py
pause
