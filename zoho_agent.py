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
                       ensure_summary, export_one, export_spreadsheet,
                       fetch_queued_exports, get_export_files)

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
            # Make sure the selected survey belongs to this project; avoid blindly exporting some other survey.
            heading = ' '.join((page.locator('body').inner_text(timeout=8000) or '').split()).casefold()
            expected = ' '.join(project.split()).casefold()
            if expected not in heading:
                raise RuntimeError('Survey title could not be verified on page; refusing to export the wrong survey.')
            ensure_summary(page)
            export_one(page, saved, ['As PDF', 'PDF'], 'summary report PDF')
            export_spreadsheet(page, saved)
            fetch_queued_exports(page, saved)
            # Existing exports may take time to finish in Zoho's queue.
            pdf, excel = get_export_files(saved, folder)
            if not (pdf and excel):
                for _ in range(4):
                    page.wait_for_timeout(3000)
                    fetch_queued_exports(page, saved)
                    pdf, excel = get_export_files(saved, folder)
                    if pdf and excel: break
            if not (pdf and excel):
                raise RuntimeError('Zoho did not provide both PDF and XLSX. Check browser Exports and worker.log.')
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
