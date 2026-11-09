ZOHO EXPORT RETRY UPDATE — 30-second timeout
==============================================
This updates ONLY the Windows Zoho helper. Do not change your Streamlit application, tokens or GitHub configuration.

1. Stop all running zoho_agent.py processes (PowerShell):
Get-CimInstance Win32_Process | Where-Object { $_.Name -match '^pythonw?\.exe$' -and $_.CommandLine -like '*zoho_agent.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId }

2. Copy the two files from this ZIP into your existing helper folder, replacing the originals:
   zoho_agent.py
   zoho_sync.py

Helper folder on your computer:
C:\Workshop_Evaluation_One_Click_Zoho_Refresh\Workshop_Evaluation_One_Click

3. Open PowerShell in that folder. Run this one-time test:
py zoho_agent.py --once

4. Since earlier requests already failed, make a NEW refresh request in Streamlit with the correct survey URL. While testing, launch the visible foreground worker:
py zoho_agent.py

5. Check sync status in Streamlit. If an error occurs, read:
Get-Content "$env:APPDATA\WorkshopEvaluationSync\worker.log" -Tail 40

Behavior:
* Both PDF and spreadsheet use the top Export menu on Summary of Responses.
* Select As PDF / As Spreadsheet, confirm the black Export button in the popup.
* Wait up to 30 seconds for a new download notification; click "Click here" to download.
* If no fresh download is captured, reload Zoho and repeat, up to 3 attempts per file.
* No old completed exports are blindly downloaded from the Exports history.
* 2FA/login challenges must still be completed by you.

A note on Windows startup: If you have an existing auto-start helper task, it may also launch a hidden worker. Stop that worker when testing in the foreground and allow only one active process.

This code has been syntax-checked. Live Zoho flows have not been verified here.
