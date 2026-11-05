@echo off
cd /d "%~dp0"
echo Installing Zoho desktop sync helper. This is a one-time setup.
py -m pip install -r requirements-desktop.txt
if errorlevel 1 pause & exit /b 1
py -m playwright install chromium
if errorlevel 1 pause & exit /b 1
py setup_sync.py
pause
