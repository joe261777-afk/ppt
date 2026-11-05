# Workshop Evaluation — One-click Zoho refresh

This package upgrades the existing Streamlit application without changing its PowerPoint templates, response calculations, attendance image/Excel review, or report outputs.

## What is and is not automatic

The Streamlit website creates a request in the **private GitHub project-data repository**. A small helper running on **your own Windows PC** checks for requests, opens Zoho Survey using an existing signed-in Chromium profile, attempts the Summary PDF and Individual Responses Excel exports, and uploads them to the private GitHub library. The website can then retrieve them and generate the report.

**The Windows PC must be on, connected to the internet, and signed into Windows.** The helper is not a cloud service. Zoho may require you to sign in / solve 2FA in the local browser, or may change export dialogs; such requests will show an error rather than silently claiming success. A first-time survey often needs its dashboard URL. This version does not guarantee unattended exporting for every Zoho survey.

## A. Update existing Streamlit website

1. Unzip the package and update the existing GitHub repository (`ppt`) with `app.py`, `report_engine.py`, `zoho_library.py`, `sync_queue.py`, `requirements.txt`, `packages.txt`, and `templates/` preserving the directory structure. Do **not** upload the zip as one file.
2. In Streamlit Community Cloud > Manage app > Settings > Secrets add/retain:

```toml
PROJECT_LIBRARY_REPO = "joe261777-afk/zoho-evaluation-data"
GITHUB_TOKEN = "your-private-fine-grained-GitHub-token"
SYNC_BUTTON_PASSWORD = "choose-a-long-unique-passphrase"
```

3. Redeploy/reboot. The `SYNC_BUTTON_PASSWORD` is used only to authorize new sync requests from the website. Keep the site access restricted too. Never commit tokens or passwords to GitHub.

## B. One-time Windows install (no PowerShell each time)

1. Copy the extracted package to a **permanent local folder** (for example `C:\WorkshopEvaluationSync`). **Do not move it after installation.**
2. Install standard Python for Windows if necessary; ensure `py` is available.
3. Double-click `Install_AutoSync_Windows.bat`.
4. Enter the data repository name (`joe261777-afk/zoho-evaluation-data`) and the existing GitHub token when prompted. The token is stored via Windows Credential Manager (`keyring`) rather than in a plaintext config file.
5. The helper starts immediately and creates a Windows Startup shortcut so that it also runs when you next sign in.

6. Before requesting the first survey, double-click `Sign_In_Zoho_Once.bat`, sign in to Zoho in its browser, then press Enter in that window. If the background helper has already opened that profile, finish its current request or stop the helper first (two Chromium instances cannot own the same profile).

The helper maintains its own Chromium profile under `%APPDATA%\WorkshopEvaluationSync\chromium_profile`. A Zoho login or 2FA may be required in that browser before unattended export requests can work. To troubleshoot, check `%APPDATA%\WorkshopEvaluationSync\worker.log` and the sync status shown on the website.

## C. Everyday use

1. Open your Streamlit website.
2. Under **Saved Zoho project**, enter the exact Zoho Survey name and, preferably, its dashboard URL (copied from the survey in Zoho). The dashboard URL is used only for navigation and must belong to an official Zoho Survey domain.
3. Enter your sync access code and click **Request fresh Zoho PDF + Excel**. A request is saved in the private GitHub repository.
4. Click **Check Zoho sync status**. When `DONE`, click **Refresh project list**, select the project and **Retrieve saved Zoho PDF and spreadsheet**.
5. Upload attendance Excel or photograph, review statuses, generate and download the PowerPoint.

If a request says `ERROR` or `LOGIN_REQUIRED`, fix the local Zoho session or export controls; the previously saved survey data stays intact. Manual PDF/Excel upload remains available. The refresh request is not evidence of new data until `DONE` appears.

## Security

- Don't store Zoho credentials or browser cookies on GitHub or Streamlit.
- Keep the data repository private, limit the GitHub token to that repository with Contents read/write, and protect the Streamlit website.
- The website's sync button requires a separate passphrase; use a strong value not reused elsewhere.
- Uploaded survey responses and attendance data may include confidential personal information. Follow your company's retention and access rules.
- The local helper does not bypass CAPTCHA or two-factor authentication.
