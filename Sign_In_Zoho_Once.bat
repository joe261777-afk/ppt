@echo off
cd /d "%~dp0"
echo Sign in to Zoho in this browser. This is the SAME profile used by the sync worker.
py zoho_login.py
pause
