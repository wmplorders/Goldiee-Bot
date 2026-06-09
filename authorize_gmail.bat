@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" -m src.gmail_auth
pause
