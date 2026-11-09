"""Zoho Survey desktop export automation for the Impact evaluation generator.

Uses the user's own signed-in Chromium profile. Attempts Zoho Survey's
Reports > Summary of Responses > Export > As PDF, then Individual Responses >
Export > As Spreadsheet. Zoho may queue files under Exports, which this helper
also checks. Keeps manual fallback if Zoho changes its UI or requires input.
Never automates passwords, two-factor challenges or CAPTCHA.
"""
import argparse
import getpass
import os
import re
import sys
import time  # required by the 30-second export notification polling loop
from pathlib import Path
from urllib.parse import urlparse
from playwright.sync_api import sync_playwright
from zoho_library import Library

ALLOWED_HOSTS = ('survey.zoho.com', 'survey.zoho.eu', 'survey.zoho.in',
                 'survey.zoho.com.au', 'survey.zoho.jp', 'survey.zoho.com.cn')
ROOT = Path(__file__).resolve().parent


def visible(locator):
    try:
        return locator.count() and locator.first.is_visible()
    except Exception:
        return False


def click_first(page, names, timeout=2500):
    """Try accessible controls first and then exact text, without arbitrary CSS."""
    for name in names:
        locators = [page.get_by_role('button', name=name, exact=True),
                    page.get_by_role('link', name=name, exact=True),
                    page.get_by_text(name, exact=True)]
        for loc in locators:
            try:
                if visible(loc):
                    loc.first.click(timeout=timeout)
                    return True
            except Exception:
                continue
    return False


def click_top_export(page):
    """Prefer header Export over the separate Export button on chart cards."""
    candidates = page.get_by_text('Export', exact=True)
    try:
        positions = []
        for idx in range(min(candidates.count(), 12)):
            el = candidates.nth(idx)
            if not el.is_visible():
                continue
            rect = el.bounding_box()
            if rect:
                # Header Export in screenshot is at top right at ~y=157.
                positions.append((0 if rect['y'] < 230 else 1,
                                  rect['y'], -rect['x'], idx))
        for _, _, _, idx in sorted(positions):
            try:
                candidates.nth(idx).click(timeout=3500)
                return True
            except Exception:
                pass
    except Exception:
        pass
    return click_first(page, ['Export'])


def wait_for_new_downloads(page, saved, prior_count, seconds=9):
    # Playwright sync API needs its own wait cycle to dispatch download events.
    for _ in range(int(seconds * 2)):
        if len(saved) > prior_count:
            return True
        page.wait_for_timeout(500)
    return len(saved) > prior_count


def goto_survey(page, project):
    """Return true only if current page points to a survey (not merely the portal)."""
    if re.search(r'/survey/\d+/', page.url):
        return True
    page.wait_for_timeout(1500)
    # A project may be visible as a card, link or heading on the survey list.
    for loc in [page.get_by_text(project, exact=True),
                page.get_by_role('link', name=project, exact=True)]:
        try:
            if visible(loc):
                loc.first.click(timeout=4500)
                page.wait_for_timeout(1500)
                if re.search(r'/survey/\d+/', page.url):
                    return True
        except Exception:
            pass
    # Attempt a local filter/search, if present.
    for placeholder in ['Search surveys', 'Search', 'Search Surveys']:
        loc = page.get_by_placeholder(placeholder, exact=False)
        try:
            if visible(loc):
                loc.first.fill(project, timeout=2500)
                page.wait_for_timeout(1200)
                match = page.get_by_text(project, exact=True)
                if visible(match):
                    match.first.click(timeout=4500)
                    page.wait_for_timeout(1500)
                    return bool(re.search(r'/survey/\d+/', page.url))
        except Exception:
            continue
    return False


def zoho_route_error(page):
    """Detect Zoho's plain-text invalid-route response before any export clicks."""
    try:
        return 'URL_RULE_NOT_CONFIGURED' in page.locator('body').inner_text(timeout=3000)
    except Exception:
        return False


def open_zoho_safely(page, target_url):
    """Open the service root first; avoid direct deep links into an unauthenticated SPA."""
    from urllib.parse import urlparse
    host = urlparse(target_url).netloc
    root = f'https://{host}/'
    print('Opening Zoho homepage:', root)
    page.goto(root, wait_until='commit', timeout=90000)
    page.locator('body').wait_for(state='visible', timeout=30000)
    page.wait_for_timeout(3000)
    page.wait_for_timeout(1200)
    if zoho_route_error(page):
        print('Zoho returned URL_RULE_NOT_CONFIGURED on its homepage.')
        print('Try the normal Zoho Survey site in your regular browser; check domain,')
        print('sign-in status, extensions, VPN or network restrictions before retrying.')
        return False
    if '/survey/' in urlparse(target_url).path:
        print('Open the named survey from the dashboard, or use the target URL if needed:')
        print(target_url)
    return True


def survey_page_base(url):
    m = re.search(r'^(https://[^/]+/survey/newui/#/portal/[^/]+/department/[^/]+/survey/\d+)', url)
    return m.group(1) if m else None


def ensure_summary(page):
    base = survey_page_base(page.url)
    if base:
        # The Zoho report screenshot includes this exact SPA route.
        page.goto(base + '/reports/summaryresponse', wait_until='domcontentloaded', timeout=60000)
        page.wait_for_timeout(1500)
    else:
        click_first(page, ['Reports', 'REPORTS'])
        page.wait_for_timeout(700)
        click_first(page, ['Summary of Responses', 'SUMMARY OF RESPONSES'])
        page.wait_for_timeout(1000)


def _download_of_type(saved, extension, after=0):
    """Only count downloads captured after the current export action."""
    return next((p for p in saved[after:] if p.is_file() and
                 p.suffix.lower() in extension), None)


def _handle_export_modal(page):
    """Submit Zoho's second (black Export) button, not the menu button."""
    try:
        # Zoho uses a styled DIV modal which may not have aria role=dialog.
        heading = page.get_by_text(re.compile(r'^Export as (PDF|Spreadsheet|Excel)$', re.I))
        if heading.count():
            # Find the nearest ancestor that contains both the heading and an Export button.
            panel = heading.first.locator('xpath=ancestor::*[.//button[normalize-space()="Export"]][1]')
            if panel.count():
                btn = panel.get_by_role('button', name='Export', exact=True)
                if visible(btn):
                    btn.first.click(timeout=6000)
                    print('  Confirmed Export in settings dialog.', flush=True)
                    return True
        for modal_selector in ['[role="dialog"]', '.modal.show', '.modal-dialog', '[class*="modal-content"]']:
            panel = page.locator(modal_selector)
            for index in range(min(panel.count(), 4)):
                if panel.nth(index).is_visible():
                    btn = panel.nth(index).get_by_role('button', name='Export', exact=True)
                    if visible(btn):
                        btn.first.click(timeout=6000)
                        print('  Confirmed Export in settings dialog.', flush=True)
                        return True
    except Exception as exc:
        print('  Export dialog handling:', str(exc)[:120])
    return False


def _dismiss_scheduled_dialog(page):
    """Acknowledgement of the queued export; not a download confirmation."""
    try:
        popup = page.get_by_text('The file export request has been scheduled', exact=False)
        if visible(popup):
            btn = page.get_by_role('button', name='OK', exact=True)
            if visible(btn):
                btn.first.click(timeout=3000)
                return True
    except Exception:
        pass
    return False


def _click_ready_notification(page):
    """Click Zoho's completed-export toast only when its success text is present."""
    try:
        body_text = page.locator('body').inner_text(timeout=2500).casefold()
    except Exception:
        return False

    # Do not click arbitrary "Click here" links elsewhere on the report.
    if 'export has been completed successfully' not in body_text:
        return False

    candidates = [
        page.get_by_role('link', name=re.compile(r'click here', re.I)),
        page.get_by_text('Click here', exact=True),
        page.get_by_text(re.compile(r'click here to download the file', re.I)),
    ]
    for group in candidates:
        try:
            for i in range(min(group.count(), 5)):
                locator = group.nth(i)
                if locator.is_visible():
                    locator.click(timeout=3500)
                    print('  Clicked Zoho completed-export download link.', flush=True)
                    return True
        except Exception:
            continue
    return False


def _poll_export_ready(page, saved, prior, extension, seconds=30):
    """Wait for the ready toast/download, printing progress and clicking the toast."""
    deadline = time.monotonic() + seconds
    next_status = seconds
    clicked_toast = False

    while time.monotonic() < deadline:
        if _download_of_type(saved, extension, prior):
            print('  Download captured successfully.', flush=True)
            return True

        _dismiss_scheduled_dialog(page)
        if not clicked_toast and _click_ready_notification(page):
            clicked_toast = True

        # Give the browser time to dispatch the download event after the toast click.
        if clicked_toast:
            for _ in range(10):
                if _download_of_type(saved, extension, prior):
                    print('  Download captured successfully.', flush=True)
                    return True
                page.wait_for_timeout(300)
            clicked_toast = False  # If no download began, look for a reappearing toast.

        remaining = max(0, int(deadline - time.monotonic()))
        if remaining <= next_status:
            print(f'  Waiting for Zoho download notification... {remaining}s remaining.', flush=True)
            next_status = remaining - 5
        page.wait_for_timeout(500)

    return bool(_download_of_type(saved, extension, prior))

def export_with_retry(page, saved, kind, max_attempts=3, timeout_seconds=30):
    """Complete Zoho's two-stage export, retrying after 30s without a new file."""
    if kind not in ('pdf', 'xlsx'):
        raise ValueError('Unsupported export type')
    ext = {'.pdf'} if kind == 'pdf' else {'.xlsx', '.xlsm'}
    label = 'summary PDF' if kind == 'pdf' else 'responses spreadsheet'
    options = ['As PDF', 'PDF'] if kind == 'pdf' else ['As Spreadsheet', 'Spreadsheet', 'As Excel', 'Excel']
    for attempt in range(1, max_attempts + 1):
        prior = len(saved)
        print(f'Exporting {label}: attempt {attempt}/{max_attempts}', flush=True)
        ensure_summary(page)
        if not click_top_export(page):
            print('  Could not open the top Export menu.', flush=True)
        elif not click_first(page, options):
            print(f'  Could not select {label} in Export menu.', flush=True)
        else:
            page.wait_for_timeout(500)
            if not _handle_export_modal(page):
                print('  No settings dialog Export confirmation found; checking notification anyway.', flush=True)
            if _poll_export_ready(page, saved, prior, ext, timeout_seconds):
                return _download_of_type(saved, ext, prior)
        if attempt < max_attempts:
            print(f'  No fresh {label} after {timeout_seconds} seconds. Refreshing page and restarting export.', flush=True)
            page.reload(wait_until='domcontentloaded', timeout=60000)
            page.wait_for_timeout(1300)
    raise RuntimeError(f'Zoho did not download a new {label} after {max_attempts} attempts.')


def export_one(page, saved, option_names, label):
    """Compatibility wrapper used by older scripts."""
    return bool(export_with_retry(page, saved, 'pdf' if any('PDF' in x for x in option_names) else 'xlsx'))


def export_spreadsheet(page, saved):
    """Export spreadsheet using the same Summary of Responses top-right Export."""
    return bool(export_with_retry(page, saved, 'xlsx'))


def fetch_queued_exports(page, saved, max_clicks=8):
    """Compatibility fallback: each new export is handled by export_with_retry."""
    return None


def get_export_files(saved, folder):
    paths = list(dict.fromkeys([p for p in saved if p.is_file()] +
                              list(folder.glob('*.pdf')) + list(folder.glob('*.xlsx'))))
    pdfs = [p for p in paths if p.suffix.lower() == '.pdf' and p.read_bytes()[:4] == b'%PDF']
    excels = [p for p in paths if p.suffix.lower() in ('.xlsx', '.xlsm') and p.read_bytes()[:2] == b'PK']
    pdfs.sort(key=lambda p: p.stat().st_mtime)
    excels.sort(key=lambda p: p.stat().st_mtime)
    return (pdfs[-1] if pdfs else None, excels[-1] if excels else None)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--project', required=True, help='Exact Zoho survey/project name')
    parser.add_argument('--url', default='https://survey.zoho.com/',
                        help='Zoho dashboard URL or a specific survey URL (recommended)')
    parser.add_argument('--downloads', default='./zoho_downloads')
    parser.add_argument('--pdf', default='', help='Existing summary PDF; skip browser')
    parser.add_argument('--excel', default='', help='Existing response XLSX; skip browser')
    parser.add_argument('--repo', default=os.environ.get('PROJECT_LIBRARY_REPO', ''))
    args = parser.parse_args()
    url = urlparse(args.url)
    if url.scheme != 'https' or url.hostname not in ALLOWED_HOSTS:
        parser.error('Use a legitimate https://survey.zoho.* URL')
    if bool(args.pdf) != bool(args.excel):
        parser.error('Provide both --pdf and --excel, or neither.')

    folder = Path(args.downloads).resolve() / ''.join(
        c if c.isalnum() or c in '-_ ' else '_' for c in args.project)
    folder.mkdir(parents=True, exist_ok=True)
    print('\nProject:', args.project, '\nSave location:', folder)
    if args.pdf:
        pdf, excel = Path(args.pdf).expanduser().resolve(), Path(args.excel).expanduser().resolve()
    else:
        saved = []
        with sync_playwright() as pw:
            ctx = pw.chromium.launch_persistent_context(
                user_data_dir=str(Path.home()/'.workshop_zoho_browser'),
                headless=False, accept_downloads=True, viewport={'width': 1500, 'height': 950})

            def record(download):
                suggested = Path(download.suggested_filename).name
                dest = folder / f'{len(saved)+1:02d}_{suggested}'
                try:
                    download.save_as(str(dest))
                    saved.append(dest)
                    print('Captured:', dest.name)
                except Exception as exc:
                    print('Could not save download:', exc)

            def attach(pg):
                pg.on('download', record)

            ctx.on('page', attach)
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            for pg in ctx.pages:
                attach(pg)
            if not open_zoho_safely(page, args.url):
                print('Please confirm Zoho works in your regular browser before attempting exports.')
                input('After fixing the Zoho page in this browser, press ENTER to retry...')
                if zoho_route_error(page):
                    ctx.close()
                    sys.exit('Zoho still returns URL_RULE_NOT_CONFIGURED. No export requested.')
            print('\nSign in normally in this browser if requested, including 2FA.')
            input('When Zoho is signed in, press ENTER to automate both exports...')
            if zoho_route_error(page):
                ctx.close()
                sys.exit('Zoho is showing URL_RULE_NOT_CONFIGURED. Fix navigation before export.')
            # Do not trust a project URL alone; require the correct survey title
            # when navigating manually through the dashboard.
            if not goto_survey(page, args.project):
                print('Project lookup could not open the survey automatically.')
                print('Open the matching survey in the SAME browser window.')
                input('After opening the survey dashboard, press ENTER ...')
                if zoho_route_error(page):
                    ctx.close()
                    sys.exit('Zoho survey navigation failed: URL_RULE_NOT_CONFIGURED.')
            if not re.search(r'/survey/\d+/', page.url):
                print('WARNING: Current URL does not appear to be inside a survey.')
            try:
                ensure_summary(page)
                export_one(page, saved, ['As PDF', 'PDF'], 'summary report PDF')
                export_spreadsheet(page, saved)
                fetch_queued_exports(page, saved)
            except Exception as exc:
                print('Zoho changed its export UI or an export was blocked:', exc)
            pdf, excel = get_export_files(saved, folder)
            if not pdf or not excel:
                print('Automatic capture is incomplete. Current files:', [p.name for p in saved])
                print('Review the open browser. Finish any export confirmation or download')
                print('the queued items from Exports, without closing this browser.')
                input('After both downloads finish, press ENTER to re-check ...')
                page.wait_for_timeout(1500)
                pdf, excel = get_export_files(saved, folder)
            ctx.close()

        if not pdf or not excel:
            print('Zoho did not deliver both files to the automated browser session.')
            pdf_txt = input('Existing PDF path (blank to stop): ').strip().strip('"')
            excel_txt = input('Existing XLSX path (blank to stop): ').strip().strip('"')
            if not pdf_txt or not excel_txt:
                sys.exit('No upload performed. Try --url with the exact survey dashboard URL.')
            pdf, excel = Path(pdf_txt).expanduser().resolve(), Path(excel_txt).expanduser().resolve()

    if not pdf.is_file() or not excel.is_file():
        sys.exit(f'File not found: PDF={pdf} Excel={excel}')
    if pdf.suffix.lower() != '.pdf' or excel.suffix.lower() not in ('.xlsx', '.xlsm'):
        sys.exit('Expected one .pdf report and one .xlsx/.xlsm spreadsheet.')
    if pdf.read_bytes()[:4] != b'%PDF' or excel.read_bytes()[:2] != b'PK':
        sys.exit('Unexpected file format. No files uploaded.')
    print('Verified:', pdf.name, 'and', excel.name)
    repo = args.repo or input('Private GitHub repository (owner/repo): ').strip()
    token = os.getenv('GITHUB_TOKEN') or getpass.getpass('GitHub token (input hidden): ').strip()
    if not repo or not token:
        sys.exit('GitHub credentials missing. Local exports remain saved.')
    Library(repo, token).save_project(args.project, pdf, excel)
    print('\nSUCCESS. In Streamlit select this project:', args.project)


if __name__ == '__main__':
    main()
