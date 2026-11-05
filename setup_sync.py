"""One-time Windows installation. Store GitHub token in Windows Credential Manager."""
import getpass
import json
import os
import sys
from pathlib import Path
import keyring
from zoho_agent import CONFIG, CONFIG_DIR, SERVICE
from zoho_library import Library

if sys.platform != 'win32':
    raise SystemExit('This installer is intended for Windows. For other systems run zoho_agent.py with keyring configured.')
repo = input('Private GitHub repository (owner/repo): ').strip()
token = getpass.getpass('GitHub token (hidden): ').strip()
Library(repo, token).list_projects()  # verify permissions before saving
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG.write_text(json.dumps({'repo': repo}), encoding='utf-8')
keyring.set_password(SERVICE, repo, token)
print('Saved token in Windows Credential Manager. Starting persistent helper at Windows sign-in.')
# Launcher installed as shortcut in Startup using PowerShell to avoid storing token in command line.
import subprocess
pythonw = Path(sys.executable).with_name('pythonw.exe')
script = Path(__file__).resolve().with_name('zoho_agent.py')
if not pythonw.exists():
    raise SystemExit('pythonw.exe missing. Install regular Windows Python.')
ps = r'''$startup=[Environment]::GetFolderPath('Startup'); $lnk=Join-Path $startup 'Workshop Zoho AutoSync.lnk'; $w=New-Object -ComObject WScript.Shell; $s=$w.CreateShortcut($lnk); $s.TargetPath=$env:PYTHONW; $s.Arguments='"'+$env:AGENT+'"'; $s.WorkingDirectory=$env:WORKDIR; $s.Save()'''
env = os.environ.copy();env.update({'PYTHONW':str(pythonw),'AGENT':str(script),'WORKDIR':str(script.parent)})
subprocess.run(['powershell','-NoProfile','-Command',ps],env=env,check=True)
subprocess.Popen([str(pythonw),str(script)],cwd=str(script.parent),creationflags=subprocess.DETACHED_PROCESS)
print('Installed. The helper runs now and starts when you sign in to Windows. Log:', CONFIG_DIR/'worker.log')
print('Use the Streamlit website to request a fresh Zoho export. Do not put the token in your code.')
