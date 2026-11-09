"""Windows background worker for Zoho Survey exports and GitHub Project Library.

Key reliability changes:
- Cache each request's verified PDF/XLSX under %APPDATA%/WorkshopEvaluationSync/downloads/<request-id>.
- If an export already exists locally, do not repeat that Zoho export.
- Retry transient GitHub uploads independently from Zoho exports.
- If GitHub is offline, leave the request pending so the next worker cycle uploads
  the cached files instead of exporting them again.
- Recover stale "processing" requests after an interrupted run.
- Allow an explicit retry of one failed request with --retry-request <id>.
"""
import argparse
import json
import logging
import os
import re
import shutil
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import urlparse

import requests
from playwright.sync_api import sync_playwright

from zoho_library import Library
from sync_queue import (
    get_request, list_requests, update_request, valid_url
)
from zoho_sync import (
    ALLOWED_HOSTS, open_zoho_safely, zoho_route_error, goto_survey,
    ensure_summary, export_with_retry, get_export_files
)

CONFIG_DIR = Path(os.environ.get('APPDATA', str(Path.home()))) / 'WorkshopEvaluationSync'
CONFIG = CONFIG_DIR / 'config.json'
LOG = CONFIG_DIR / 'worker.log'
SERVICE = 'WorkshopEvaluationSync_GitHub'
DOWNLOADS_DIR = CONFIG_DIR / 'downloads'
UPLOAD_ATTEMPTS = 5
STALE_PROCESSING_MINUTES = 5


def get_credentials():
    import keyring
    cfg = json.loads(CONFIG.read_text(encoding='utf-8'))
    token = keyring.get_password(SERVICE, cfg['repo'])
    if not token:
        raise ValueError(
            'No GitHub token saved in Windows Credential Manager. Run setup_sync.py.'
        )
    return cfg, token


def verify_survey_identity(page, project, requested_url=''):
    """Prevent wrong-survey exports; prefer the numeric survey ID in a supplied URL."""
    def survey_id(url):
        match = re.search(r'/survey/(\d+)(?:/|$)', urlparse(url).fragment)
        return match.group(1) if match else None

    current_id = survey_id(page.url)
    requested_id = survey_id(requested_url) if requested_url else None
    if requested_url:
        if not requested_id:
            raise RuntimeError(
                'Requested URL does not contain a Zoho survey ID. '
                'Paste its dashboard URL from Zoho.'
            )
        if current_id != requested_id:
            raise RuntimeError(
                'Survey ID mismatch: opened page is not the requested Zoho survey. '
                'Refusing export.'
            )
        return True

    def normalized(text):
        return ' '.join(re.findall(r'[\w]+', str(text).casefold(), flags=re.UNICODE))

    text = normalized(page.locator('body').inner_text(timeout=12000) or '')
    full_title = normalized(project)
    chunks = [
        normalized(x) for x in re.split(r'\s+[-–]\s+', project) if x.strip()
    ]
    meaningful_chunks = [x for x in chunks if len(x) >= 20]
    expected_parts = ([full_title] if len(full_title) >= 25 else []) + meaningful_chunks
    if not current_id or not expected_parts or not any(part in text for part in expected_parts):
        raise RuntimeError(
            'Could not safely verify the survey title. Enter the exact Zoho survey '
            'dashboard URL in the Streamlit refresh form and submit a new request. '
            'The helper verifies its numeric survey ID before exporting.'
        )
    return True


def request_download_folder(request_id):
    if not re.fullmatch(r'[a-f0-9]{32}', str(request_id or '')):
        raise ValueError('Invalid request ID for local download cache.')
    folder = DOWNLOADS_DIR / request_id
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def is_valid_pdf(path):
    try:
        path = Path(path)
        return path.is_file() and path.stat().st_size > 100 and path.open('rb').read(4) == b'%PDF'
    except OSError:
        return False


def is_valid_excel(path):
    try:
        path = Path(path)
        return path.is_file() and path.stat().st_size > 100 and path.open('rb').read(2) == b'PK'
    except OSError:
        return False


def cached_exports(folder):
    """Return only the deterministic, request-specific cache files."""
    pdf = folder / 'survey_report.pdf'
    excel = folder / 'survey_responses.xlsx'
    return (
        pdf if is_valid_pdf(pdf) else None,
        excel if is_valid_excel(excel) else None
    )


def adopt_legacy_downloads(request_id, source_folder):
    """Adopt files captured by an older worker's random-UUID download folder."""
    source_folder = Path(source_folder).expanduser().resolve()
    if not source_folder.is_dir():
        raise ValueError(f'Legacy download folder does not exist: {source_folder}')

    pdf_candidates = [p for p in source_folder.rglob('*.pdf') if is_valid_pdf(p)]
    excel_candidates = [
        p for p in source_folder.rglob('*.xlsx') if is_valid_excel(p)
    ]
    if not pdf_candidates or not excel_candidates:
        raise RuntimeError(
            'The selected legacy folder must contain at least one valid PDF and one valid XLSX.'
        )

    pdf_source = max(pdf_candidates, key=lambda p: p.stat().st_mtime)
    excel_source = max(excel_candidates, key=lambda p: p.stat().st_mtime)
    destination = request_download_folder(request_id)
    pdf = persist_export(
        pdf_source, destination / 'survey_report.pdf', is_valid_pdf
    )
    excel = persist_export(
        excel_source, destination / 'survey_responses.xlsx', is_valid_excel
    )
    print('Adopted existing local exports; Zoho will not be opened:', flush=True)
    print('  PDF:  ', pdf, flush=True)
    print('  Excel:', excel, flush=True)
    return pdf, excel


def persist_export(source, destination, validator):
    """Copy a verified download into its stable request-specific cache location."""
    source = Path(source)
    destination = Path(destination)
    if not validator(source):
        raise RuntimeError(f'Download did not pass file validation: {source.name}')
    if source.resolve() != destination.resolve():
        temp = destination.with_suffix(destination.suffix + '.tmp')
        shutil.copy2(source, temp)
        if not validator(temp):
            temp.unlink(missing_ok=True)
            raise RuntimeError(f'Cached file failed validation: {destination.name}')
        temp.replace(destination)
    return destination


def capture_latest(project, target_url, request_id):
    """Capture missing exports only; retain completed files across upload retries."""
    folder = request_download_folder(request_id)
    pdf_path = folder / 'survey_report.pdf'
    excel_path = folder / 'survey_responses.xlsx'
    pdf, excel = cached_exports(folder)

    if pdf and excel:
        print(
            f'Using both previously captured files for request {request_id}. '
            'Zoho will not be opened again.',
            flush=True
        )
        return pdf, excel

    saved = []
    target_url = valid_url(target_url) if target_url else 'https://survey.zoho.com/'

    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            user_data_dir=str(CONFIG_DIR / 'chromium_profile'),
            headless=False,
            accept_downloads=True,
            viewport={'width': 1500, 'height': 950}
        )

        def record(download):
            suggested = Path(download.suggested_filename).name
            dest = folder / f'{len(saved) + 1:02d}_{suggested}'
            try:
                download.save_as(str(dest))
                saved.append(dest)
                print(f'Captured locally: {dest.name}', flush=True)
                logging.info('Captured download for request %s: %s', request_id, dest)
            except Exception:
                logging.exception('Could not save browser download for request %s', request_id)
                raise

        def attach(page):
            page.on('download', record)

        context.on('page', attach)
        page = context.pages[0] if context.pages else context.new_page()
        for existing_page in context.pages:
            attach(existing_page)

        try:
            if not open_zoho_safely(page, target_url):
                raise RuntimeError(
                    'Zoho homepage refused navigation. Open the local browser and sign in.'
                )

            if target_url != 'https://survey.zoho.com/':
                page.goto(target_url, wait_until='domcontentloaded', timeout=60000)
                page.wait_for_timeout(2000)

            if zoho_route_error(page):
                raise RuntimeError(
                    'Zoho URL_RULE_NOT_CONFIGURED. Open the survey from the Zoho dashboard.'
                )

            if not goto_survey(page, project):
                raise RuntimeError(
                    'Survey not found automatically. Open it once in the local helper '
                    'or provide its exact Zoho URL.'
                )

            verify_survey_identity(
                page, project,
                target_url if target_url != 'https://survey.zoho.com/' else ''
            )

            # If one file survived an earlier attempt, reuse it and export only the missing one.
            if not pdf or not is_valid_pdf(pdf_path):
                print('Starting summary PDF export...', flush=True)
                downloaded_pdf = export_with_retry(
                    page, saved, 'pdf', max_attempts=3, timeout_seconds=30
                )
                if not downloaded_pdf:
                    raise RuntimeError('Zoho did not return a new summary PDF.')
                pdf = persist_export(downloaded_pdf, pdf_path, is_valid_pdf)
                print(f'PDF safely cached: {pdf}', flush=True)
            else:
                print('PDF already cached; skipping PDF export.', flush=True)

            if not excel or not is_valid_excel(excel_path):
                print('Starting responses spreadsheet export...', flush=True)
                downloaded_excel = export_with_retry(
                    page, saved, 'xlsx', max_attempts=3, timeout_seconds=30
                )
                if not downloaded_excel:
                    raise RuntimeError('Zoho did not return a new responses spreadsheet.')
                excel = persist_export(downloaded_excel, excel_path, is_valid_excel)
                print(f'Excel safely cached: {excel}', flush=True)
            else:
                print('Excel already cached; skipping spreadsheet export.', flush=True)

            if not (is_valid_pdf(pdf_path) and is_valid_excel(excel_path)):
                raise RuntimeError('Both fresh Zoho exports were not captured and validated.')

            return pdf_path, excel_path
        finally:
            try:
                context.close()
            except Exception:
                logging.exception('Could not close Chromium context for request %s', request_id)


def is_transient_github_error(exc):
    """Identify network/rate-limit/server errors worth retrying; never retry bad credentials."""
    if isinstance(exc, requests.exceptions.RequestException):
        return True
    message = str(exc).casefold()
    transient_markers = (
        'timed out', 'timeout', 'name resolution', 'failed to resolve',
        'connection reset', 'connection aborted', 'connection error',
        'temporarily unavailable', 'remote end closed', '429', '500',
        '502', '503', '504', 'rate limit'
    )
    return any(marker in message for marker in transient_markers)


def github_retry(operation, label, attempts=UPLOAD_ATTEMPTS):
    delays = (3, 6, 12, 20, 30)
    for attempt in range(1, attempts + 1):
        try:
            result = operation()
            if attempt > 1:
                print(f'{label}: succeeded on retry {attempt}/{attempts}.', flush=True)
            return result
        except Exception as exc:
            if not is_transient_github_error(exc) or attempt >= attempts:
                raise
            delay = delays[min(attempt - 1, len(delays) - 1)]
            print(
                f'{label}: temporary GitHub/network error on attempt '
                f'{attempt}/{attempts}: {str(exc)[:180]}',
                flush=True
            )
            print(f'Will retry in {delay} seconds. Local exports are preserved.', flush=True)
            logging.warning(
                '%s failed on attempt %s/%s: %s',
                label, attempt, attempts, exc
            )
            time.sleep(delay)


def update_status_safely(lib, item, status, message):
    return github_retry(
        lambda: update_request(lib, item, status, message),
        f'Updating request status to {status}',
        attempts=3
    )


def process(lib, item):
    request_id = item['id']
    project = item['project']
    print(f'\nProcessing request {request_id}: {project}', flush=True)

    try:
        item = update_status_safely(
            lib, item, 'processing', 'Opening Zoho or reusing locally cached exports'
        )
    except Exception:
        logging.exception('Could not mark request %s as processing', request_id)
        print(
            'Could not update the GitHub queue. No Zoho export was started. '
            'The request can be retried when GitHub connectivity returns.',
            flush=True
        )
        return

    # Stage 1: capture/reuse the files. A Zoho/UI error is not treated as an upload error.
    try:
        pdf, excel = capture_latest(project, item.get('url', ''), request_id)
        print('Both exports are safely stored locally:', flush=True)
        print('  PDF:  ', pdf, flush=True)
        print('  Excel:', excel, flush=True)
    except Exception as exc:
        logging.exception('Zoho capture failed for request %s', request_id)
        message = str(exc)
        status = (
            'login_required'
            if re.search(r'login|sign in|authentication|2fa|captcha', message, re.I)
            else 'error'
        )
        try:
            update_status_safely(lib, item, status, message)
        except Exception:
            logging.exception('Could not update failed capture status for %s', request_id)
        print(f'Zoho capture failed: {message}', flush=True)
        return

    # Stage 2: upload. If GitHub/network is unavailable, keep files and requeue this
    # request. The next pass sees the request-specific cache and skips Zoho entirely.
    print('Uploading PDF, Excel and project manifest to private GitHub...', flush=True)
    try:
        github_retry(
            lambda: lib.save_project(project, pdf, excel),
            'Uploading project files to GitHub',
            attempts=UPLOAD_ATTEMPTS
        )
        print('GitHub project upload completed.', flush=True)
    except Exception as exc:
        logging.exception('GitHub upload failed for request %s', request_id)
        message = (
            'PDF and Excel are safely cached on this PC. GitHub upload has not completed; '
            'the helper will retry the upload without repeating Zoho exports. '
            + str(exc)
        )
        try:
            update_status_safely(lib, item, 'pending', message)
            print(
                'Upload is pending. Local PDF/Excel are preserved; the next worker '
                'cycle will retry GitHub only.',
                flush=True
            )
        except Exception:
            logging.exception(
                'Could not requeue upload for request %s; stale-processing recovery will retry it',
                request_id
            )
            print(
                'GitHub is unreachable and queue status could not be updated. '
                'Local exports remain cached; the worker will recover this request later.',
                flush=True
            )
        return

    # Stage 3: mark complete. If this status write fails, stale-processing recovery
    # will re-run the idempotent upload from the local cache, not the browser.
    try:
        update_status_safely(
            lib, item, 'done',
            'Latest PDF and XLSX saved to the private Project Library'
        )
        print('\nSUCCESS. Project files are in the private GitHub library:', project, flush=True)
    except Exception:
        logging.exception('Files uploaded but could not mark request %s done', request_id)
        print(
            'Files were uploaded, but the completion status could not reach GitHub. '
            'The helper will recover the status on a later cycle.',
            flush=True
        )


def processing_is_stale(item):
    if item.get('status') != 'processing':
        return False
    stamp = item.get('updated_at') or item.get('created_at')
    if not stamp:
        return True
    try:
        when = datetime.fromisoformat(stamp.replace('Z', '+00:00'))
        if when.tzinfo is None:
            when = when.replace(tzinfo=timezone.utc)
        return datetime.now(timezone.utc) - when > timedelta(minutes=STALE_PROCESSING_MINUTES)
    except (TypeError, ValueError):
        return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--once', action='store_true', help='Check queued requests once and exit')
    parser.add_argument('--interval', type=int, default=20, help='Queue polling frequency in seconds')
    parser.add_argument(
        '--retry-request', default='',
        help='Explicitly requeue one failed request ID; cached exports are reused if available'
    )
    parser.add_argument(
        '--adopt-download-folder', default='',
        help='Copy the newest valid PDF/XLSX from an older local download folder before retrying'
    )
    args = parser.parse_args()

    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=str(LOG), level=logging.INFO,
        format='%(asctime)s %(levelname)s %(message)s'
    )

    cfg, token = get_credentials()
    lib = Library(cfg['repo'], token)
    print('Windows Zoho sync helper active. Leave Windows signed in.', flush=True)
    print('Local export cache:', DOWNLOADS_DIR, flush=True)

    if args.retry_request:
        item = get_request(lib, args.retry_request.strip())
        if not item:
            parser.error('Request ID was not found in the GitHub queue.')
        if item.get('status') not in ('error', 'pending', 'processing', 'login_required'):
            parser.error(
                f"Request status is {item.get('status')!r}; only failed/pending requests can be retried."
            )
        if args.adopt_download_folder:
            adopt_legacy_downloads(item['id'], args.adopt_download_folder)
        item = update_status_safely(
            lib, item, 'pending', 'Manually retried; reuse any locally cached exports'
        )
        process(lib, item)
        if args.once:
            return

    while True:
        try:
            requests_list = list_requests(lib)
            work = []
            for item in reversed(requests_list):
                if item.get('status') == 'pending':
                    work.append(item)
                elif processing_is_stale(item):
                    logging.warning('Recovering stale processing request %s', item.get('id'))
                    try:
                        item = update_status_safely(
                            lib, item, 'pending',
                            'Recovered after an interrupted worker; reuse cached exports'
                        )
                        work.append(item)
                    except Exception:
                        logging.exception(
                            'Could not recover stale request %s', item.get('id')
                        )

            if not work:
                if args.once:
                    print('No pending requests found. Failed requests require --retry-request ID.', flush=True)
                else:
                    print('No pending Zoho sync requests.', flush=True)

            for item in work:
                process(lib, item)

        except Exception:
            logging.exception('Queue processing failed')
            print(
                'Queue check failed. See worker.log; the helper will retry on its next cycle.',
                flush=True
            )

        if args.once:
            break
        time.sleep(max(15, args.interval))


if __name__ == '__main__':
    main()
