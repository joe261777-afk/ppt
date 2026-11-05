# Workshop Evaluation Generator — Zoho Free export bridge

This package preserves the existing Streamlit evaluation report generator, PowerPoint templates, attendance spreadsheet/image review and PowerPoint output. It adds **lookup by project name** from privately stored Zoho Survey export files.

## How it works

Zoho Free does not provide a verified public one-click report export API for this exact PDF + Excel workflow. Therefore, an authenticated **desktop browser helper** captures the exports in your own signed-in browser and uploads them into a **separate PRIVATE GitHub repository**. Your Streamlit website then looks up the project, downloads the two export files securely, and generates the same report. There is no requirement to upload the PDF or Excel to the website again.

The helper attempts to open the named survey and now automatically clicks **Reports → Summary of Responses → Export → As PDF**, then **Individual Responses → Export → As Spreadsheet**, based on the supplied Zoho screenshots. You still complete login or security challenges if requested; unusual export dialogs may need a manual click. The download capture, validation, private upload, website search and report retrieval are automated. **This is an assisted export bridge, not a fully unattended server-side Zoho integration.** It cannot bypass subscription restrictions, 2FA or security challenges.

## Initial setup (once)

1. Create a **second, private GitHub repository** such as `myname/zoho-evaluation-data`. Do not use a public repository: files contain participant details and comments.
2. Create a fine-grained GitHub personal access token restricted to that data repository with **Contents: Read and write**. Set an expiration and keep it private. Never paste it in your repository or source code.
3. Upload the app files (`app.py`, `report_engine.py`, `zoho_library.py`, `requirements.txt`, `packages.txt`, `templates/`) into your existing **Streamlit application** repository, replacing older files.
4. In Streamlit Cloud → Manage app → Settings → Secrets, add:

   ```toml
   PROJECT_LIBRARY_REPO = "your-github-username/zoho-evaluation-data"
   GITHUB_TOKEN = "your-fine-grained-github-token"
   ```

   Secrets remain on Streamlit Cloud; **do not commit them to GitHub**. Reboot the app if needed.
5. On your Windows PC install Python, then in this extracted folder run:

   ```powershell
   py -m pip install -r requirements-desktop.txt
   py -m playwright install chromium
   ```

## Sync a survey from your own PC

In PowerShell, inside this folder:

```powershell
$env:PROJECT_LIBRARY_REPO = "your-github-username/zoho-evaluation-data"
$env:GITHUB_TOKEN = "your-fine-grained-github-token"
py zoho_sync.py --project "Arab Bank - Flex Your Agility Muscles"
```

The script opens Zoho Survey in a locally authenticated Chromium profile. Sign in if prompted. The helper attempts to locate the survey and click Zoho's export controls to download:

- the **survey summary/report PDF** (not a PDF of just one respondent)
- the **individual responses XLSX** with the `Responses` tab

For recent Zoho versions, exports can take time to prepare; open the **Exports** area and press Download once ready. Both downloads are captured in the local helper session. Press Enter in PowerShell once both have downloaded. The helper validates file signatures and pushes them to the private GitHub data repository under `zoho_exports/<project>/`.

The helper intentionally **does not automate login, password entry, 2FA or CAPTCHA**. It does attempt the export controls, but cannot guarantee them if the live Zoho UI changes. Zoho's UI can change, so complete those steps yourself in the opened browser. The browser login profile remains on your PC.

## If the helper says `Captured files: []`

This means Playwright did not observe a download event. It does **not** necessarily
mean Zoho failed to export. Check the ordinary Windows **Downloads** folder and
Zoho **Reports → Exports**, where pending exports may need an additional Download
click. Ensure the PDF is the overall report and the XLSX contains the survey
responses. Do **not** share Zoho cookies or your GitHub token.

You can now sync already-downloaded files, without opening Zoho again:

```powershell
py zoho_sync.py --project "Suez Canal Bank - The Age of CX for Corporate Run 2 - Virtual - Ashraf Afifi - 20-21-6-2026" --pdf "C:\Users\yousef\Downloads\report.pdf" --excel "C:\Users\yousef\Downloads\responses.xlsx"
```

Replace those sample filenames with the **actual full paths**. In File Explorer,
right-click each file while holding Shift and choose **Copy as path**.
Alternatively, run the original command: when capture fails, the helper now asks
for the PDF and Excel paths before exiting.

## Each time you generate a report

1. Open your existing Streamlit website URL.
2. Select **Saved Zoho project (automatic lookup)**.
3. Type part of a project title, click **Refresh project list** if needed, select its name, then **Retrieve saved Zoho PDF and spreadsheet**.
4. Upload attendance Excel or photo. Review/edit attendance and click **Generate PowerPoint**.
5. Download the PowerPoint and audit JSON.

For a survey with new responses, run `zoho_sync.py` again to refresh its two stored exports. The website will pull updated files when you click Retrieve again.

## Troubleshooting

- **No project appears**: ensure the export helper completed successfully and the Streamlit token is valid for the *private data repo*. Then refresh.
- **GitHub 403/404**: verify repository name and token permissions. Fine-grained tokens need Contents permission.
- **Old data loaded**: sync again, then click Retrieve again (and refresh the app if needed).
- **Missing one export**: return to Zoho's Exports area and download both; the helper only uploads when both valid files are captured.
- **PDF is incorrect**: select the survey summary PDF; exporting an individual response PDF is not equivalent.
- **Metadata incorrect**: original engine derives trainer/course/date from export filename. The project title must follow your existing client/training/trainer/date naming convention (e.g. `Client - Training - Trainer - DD-MM-YYYY`). Otherwise edit the metadata extraction logic or use the manual workflow.
- **Unexpected plan limitation**: desktop automation cannot unlock disabled Zoho export features.

## Privacy

Use a private data repository and private/restricted Streamlit app. Data is transmitted to GitHub and Streamlit, so obtain appropriate authorization for participant information and signed attendance. GitHub history retains old versions until explicitly purged; overwriting an export does not delete the old contents from Git history. Do not use this storage method if your organization's privacy requirements prohibit storing survey exports in GitHub. For those cases use an approved private object store instead.

## Optional future improvement

If you can provide screenshots of the exact Zoho Survey export screens (without respondent information), the local helper's export navigation can be adapted to perform the clicks end-to-end after interactive login. Do not send passwords or 2FA codes.

## New automatic export mode (from Zoho screenshots)

The helper now attempts the PDF and Excel export clicks after your normal sign-in, with an Exports-area fallback. For fastest project lookup, give it the direct survey dashboard URL as `--url` (rather than a link to an exported document). The sample below uses the survey URL visible in your screenshots:

```powershell
py zoho_sync.py --project "Suez Canal Bank - The Age of CX for Corporate Run 2 - Virtual - Ashraf Afifi - 20-21-6-2026" --url "https://survey.zoho.com/survey/newui/#/portal/798606363/department/D3RoIG/survey/931394000007193624/dashboard"
```

The helper needs a **one-time Enter after Zoho login** because it cannot bypass interactive authentication. After that, it attempts both export clicks and the Exports page itself. If Zoho asks for confirmation or prepares reports asynchronously, it will leave the browser open for you to complete that step. This is a best-effort automation based on screenshots, **not yet tested against your signed-in Zoho account**.


## Zoho invalid-route page (`URL_RULE_NOT_CONFIGURED`)

This error is returned by Zoho before the browser can access any export controls.
The helper now opens the Zoho homepage first instead of directly entering a
`/survey/newui/#/portal/...` route. Sign in, locate the matching survey,
and press Enter. If the homepage itself shows this error, test the same
site in your usual browser and check the site domain, account sign-in,
VPN/proxy or extension restrictions. The helper stops instead of pretending
that file exports succeeded.
