"""Desktop Zoho Survey export assistant.

Requires a supported desktop browser. Sign in normally in the opened Chromium
window (session stays locally under ~/.workshop_zoho_browser). Never asks for a
Zoho password. Download PDF summary report and Responses XLSX using Zoho's UI.
This helper monitors downloads and sends two validated files to a private GitHub
repository so the Streamlit website can access them by project name.

Zoho changes its UI periodically; if automated clicks fail, complete the export
manually in the same browser window. Download capture and GitHub upload remain
automatic. This approach does NOT bypass plan restrictions or CAPTCHA/2FA.
"""
import argparse
import getpass
import os
import sys
import time
from pathlib import Path
from urllib.parse import urlparse
from playwright.sync_api import sync_playwright
from zoho_library import Library

ALLOWED_HOSTS = ('survey.zoho.com', 'survey.zoho.eu', 'survey.zoho.in', 'survey.zoho.com.au', 'survey.zoho.jp', 'survey.zoho.com.cn')


def click_one(page, variants, timeout=1200):
    for text in variants:
        for locator in [page.get_by_role('button', name=text, exact=True),
                        page.get_by_role('link', name=text, exact=True),
                        page.get_by_text(text, exact=True)]:
            try:
                if locator.count() and locator.first.is_visible():
                    locator.first.click(timeout=timeout)
                    return True
            except Exception:
                pass
    return False


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--project', required=True, help='Exact Zoho survey/project name')
    parser.add_argument('--url', default='https://survey.zoho.com/', help='Zoho Survey dashboard or survey URL')
    parser.add_argument('--downloads', default='./zoho_downloads')
    parser.add_argument('--repo', default=os.environ.get('PROJECT_LIBRARY_REPO', ''))
    args = parser.parse_args()
    u = urlparse(args.url)
    if u.scheme != 'https' or u.hostname not in ALLOWED_HOSTS:
        parser.error('Use a legitimate https://survey.zoho.* URL')
    folder = Path(args.downloads).resolve() / ''.join(c if c.isalnum() or c in '-_ ' else '_' for c in args.project)
    folder.mkdir(parents=True, exist_ok=True)
    print('\nProject:', args.project)
    print('Save location:', folder)
    print('Zoho login is completed in your local browser; do not share login cookies.')
    saved = []
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(Path.home()/'.workshop_zoho_browser'),
            headless=False, accept_downloads=True, viewport={'width': 1400, 'height': 900})
        def record(download):
            suggested = Path(download.suggested_filename).name
            # Prevent treating a stale export from a different project as a new one.
            dest = folder / f'{len(saved)+1:02d}_{suggested}'
            try:
                download.save_as(str(dest))
                saved.append(dest)
                print('Captured:', dest.name)
            except Exception as exc:
                print('Could not save download:', exc)
        ctx.on('page', lambda pg: pg.on('download', record))
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        for existing in ctx.pages: existing.on('download', record)
        pg.goto(args.url, wait_until='domcontentloaded', timeout=60000)
        print('\nSign into Zoho in this browser if prompted, including 2FA.')
        input('Once the Zoho Survey dashboard is visible, press ENTER here to attempt survey lookup...')
        try:
            # Conservative best effort: only exact text matches; do not click unrelated surveys.
            if not click_one(pg, [args.project]):
                print('Automatic survey name click unavailable. Find and open the survey yourself.')
            else:
                print('Opened matching survey name.')
            time.sleep(1)
            click_one(pg, ['Reports', 'Report'])
        except Exception as exc:
            print('Automatic navigation unavailable:', exc)
        print('\nIn the browser, make sure you are inside:', args.project)
        print('Download BOTH: (1) summary REPORT PDF and (2) individual RESPONSES Excel.')
        print('Zoho may place completed exports in the Exports area, requiring a second Download click.')
        print('All completed downloads in the opened browser are captured automatically.')
        input('After downloading both files, press ENTER here to validate the capture...')
        # Some downloads may finish in the background shortly after the dialog closes.
        time.sleep(1)
        ctx.close()
    pdfs = [p for p in saved if p.suffix.lower()=='.pdf']
    spreadsheets = [p for p in saved if p.suffix.lower() in ('.xlsx','.xlsm')]
    if not pdfs or not spreadsheets:
        print('Captured files:', [p.name for p in saved])
        sys.exit('Missing one or both required exports. Re-run and use Zoho Reports > Exports to download them.')
    pdf, xlsx = pdfs[-1], spreadsheets[-1]
    if not pdf.read_bytes().startswith(b'%PDF') or xlsx.read_bytes()[:2] != b'PK':
        sys.exit('Unexpected export format. Inspect downloaded files locally; nothing uploaded.')
    print('Verified PDF:',pdf.name,'and Excel:',xlsx.name)
    repo = args.repo or input('Private GitHub repository (owner/repo): ').strip()
    token = os.environ.get('GITHUB_TOKEN') or getpass.getpass('GitHub fine-grained token (hidden): ').strip()
    if not repo or not token: sys.exit('GitHub repo and token required to sync; local downloads were preserved.')
    Library(repo,token).save_project(args.project, pdf, xlsx)
    print('\nSUCCESS. Open your Streamlit website and select:',args.project)

if __name__ == '__main__':
    main()
