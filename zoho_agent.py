"""Windows background worker. Runs on user's PC; no Zoho password is stored."""
import argparse
import json
import logging
import os
import re
import tempfile
import time
from pathlib import Path
from urllib.parse import urlparse
from playwright.sync_api import sync_playwright
from zoho_library import Library
from sync_queue import list_requests, update_request, valid_url
from zoho_sync import (ALLOWED_HOSTS, open_zoho_safely, zoho_route_error, goto_survey,
                       ensure_summary, export_with_retry, get_export_files)

CONFIG_DIR = Path(os.environ.get('APPDATA', str(Path.home()))) / 'WorkshopEvaluationSync'
CONFIG = CONFIG_DIR / 'config.json'
LOG = CONFIG_DIR / 'worker.log'
SERVICE = 'WorkshopEvaluationSync_GitHub'


def get_credentials():
    import keyring
    cfg = json.loads(CONFIG.read_text(encoding='utf-8'))
    token = keyring.get_password(SERVICE, cfg['repo'])
    if not token:
        raise ValueError('No GitHub token saved in Windows Credential Manager. Run setup_sync.py.')
    return cfg, token



def verify_survey_identity(page, project, requested_url=''):
    """Prevent wrong-survey exports; tolerate filename-vs-dashboard title differences.

    Prefer the exact survey ID if supplied via the website. Otherwise require
    visible distinctive title text. Do not accept generic client names alone.
    """
    from urllib.parse import urlparse

    def survey_id(url):
        # IDs are in the SPA hash route, not in urllib.parse(url).path.
        match = re.search(r'/survey/(\d+)(?:/|$)', urlparse(url).fragment)
        return match.group(1) if match else None

    current_id = survey_id(page.url)
    requested_id = survey_id(requested_url) if requested_url else None
    if requested_url:
        if not requested_id:
            raise RuntimeError('Requested URL does not contain a Zoho survey ID. Paste its dashboard URL from Zoho.')
        if current_id != requested_id:
            raise RuntimeError('Survey ID mismatch: opened page is not the requested Zoho survey. Refusing export.')
        return True

    def normalized(text):
        return ' '.join(re.findall(r'[\w]+', str(text).casefold(), flags=re.UNICODE))

    text = normalized(page.locator('body').inner_text(timeout=12000) or '')
    full_title = normalized(project)
    # Require substantial, specific title information, not just the customer.
    chunks = [normalized(x) for x in re.split(r'\s+[-–]\s+', project) if x.strip()]
    meaningful_chunks = [x for x in chunks if len(x) >= 20]
    expected_parts = ([full_title] if len(full_title) >= 25 else []) + meaningful_chunks
    if not current_id or not expected_parts or not any(part in text for part in expected_parts):
        raise RuntimeError(
            'Could not safely verify the survey title. Enter the exact Zoho survey dashboard URL '
            'in the Streamlit refresh form and submit a new request. '
            'The helper will verify its numeric survey ID before exporting.'
        )
    return True


def capture_latest(project, target_url):
    """Attempt automatic exports inside a local signed-in browser.

    Stops with a recoverable error if Zoho requires manual login or changes UI.
    Does not create or supply credentials or bypass CAPTCHA/2FA.
    """
    import uuid
    folder = CONFIG_DIR / 'downloads' / uuid.uuid4().hex
    folder.mkdir(parents=True)
    saved = []
    target_url = valid_url(target_url) if target_url else 'https://survey.zoho.com/'
    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            user_data_dir=str(CONFIG_DIR / 'chromium_profile'), headless=False,
            accept_downloads=True, viewport={'width': 1500, 'height': 950})
        def record(download):
            dest = folder / (f'{len(saved)+1:02d}_' + Path(download.suggested_filename).name)
            download.save_as(dest)
            saved.append(dest)
        def attach(page): page.on('download', record)
        context.on('page', attach)
        page = context.pages[0] if context.pages else context.new_page()
        for p in context.pages: attach(p)
        try:
            if not open_zoho_safely(page, target_url):
                raise RuntimeError('Zoho homepage refused navigation. Open the local browser and sign in.')
            if target_url != 'https://survey.zoho.com/':
                # Only validated Zoho URLs are allowed, and only after homepage initialization.
                page.goto(target_url, wait_until='domcontentloaded', timeout=60000)
                page.wait_for_timeout(2000)
            if zoho_route_error(page):
                raise RuntimeError('Zoho URL_RULE_NOT_CONFIGURED. Open the survey from the Zoho dashboard.')
            if not goto_survey(page, project):
                raise RuntimeError('Survey not found automatically. Open it once in the local helper or provide its exact Zoho URL.')
            # Verify the selected survey without demanding that its short on-screen
            # title match the entire exported project filename (which often includes
            # trainer, run and dates). A user-provided URL identifies the Zoho survey
            # by its stable numeric survey ID. Never infer that ID from page text.
            verify_survey_identity(page, project, target_url if target_url != 'https://survey.zoho.com/' else '')
            # Each export is fully confirmed, then its new download captured.
            # If Zoho doesn't deliver within 30 seconds, refresh and retry.
            pdf = export_with_retry(page, saved, 'pdf', max_attempts=3, timeout_seconds=30)
            excel = export_with_retry(page, saved, 'xlsx', max_attempts=3, timeout_seconds=30)
            if not (pdf and excel):
                raise RuntimeError('Both fresh Zoho exports were not captured.')
            if pdf.read_bytes()[:4] != b'%PDF' or excel.read_bytes()[:2] != b'PK':
                raise RuntimeError('Downloaded files do not have valid PDF/XLSX signatures.')
            return pdf, excel
        finally:
            context.close()


def process(lib, item):
    item = update_request(lib, item, 'processing', 'Opening Zoho on your Windows computer')
    try:
        pdf, xlsx = capture_latest(item['project'], item.get('url', ''))
        lib.save_project(item['project'], pdf, xlsx)
        update_request(lib, item, 'done', 'Latest PDF and XLSX saved to the private Project Library')
    except Exception as exc:
        logging.exception('Sync failed for request %s', item['id'])
        msg = str(exc)
        status = 'login_required' if re.search(r'login|sign in|authentication|2fa|captcha', msg, re.I) else 'error'
        update_request(lib, item, status, msg)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--once', action='store_true', help='Check queued requests once and exit')
    parser.add_argument('--interval', type=int, default=20, help='Queue polling frequency in seconds')
    args = parser.parse_args()
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(filename=str(LOG), level=logging.INFO,
                        format='%(asctime)s %(levelname)s %(message)s')
    cfg, token = get_credentials()
    lib = Library(cfg['repo'], token)
    print('Windows Zoho sync helper active. Leave Windows signed in.')
    while True:
        try:
            for item in reversed(list_requests(lib)):
                if item['status'] == 'pending':
                    process(lib, item)
        except Exception:
            logging.exception('Queue processing failed')
        if args.once: break
        time.sleep(max(15, args.interval))


if __name__ == '__main__': main()
