"""One-time manual sign-in to the exact Chromium profile used by AutoSync."""
from playwright.sync_api import sync_playwright
from zoho_agent import CONFIG_DIR
CONFIG_DIR.mkdir(parents=True,exist_ok=True)
with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(
        user_data_dir=str(CONFIG_DIR/'chromium_profile'),
        headless=False, accept_downloads=True,
        viewport={'width':1500,'height':950})
    page = context.pages[0] if context.pages else context.new_page()
    page.goto('https://survey.zoho.com/', wait_until='domcontentloaded',timeout=60000)
    print('Sign into Zoho Survey here, including any two-factor checks.')
    input('When the survey dashboard is visible, return here and press ENTER to export the cloud session...')
    state_file = CONFIG_DIR / 'zoho_storage_state.json'
    context.storage_state(path=str(state_file))
    print(f'Zoho session state exported to: {state_file}')
    print('IMPORTANT: This file contains active Zoho session cookies. Do not share it or commit it to GitHub.')
    print('To copy its contents for a GitHub Actions secret, run this in PowerShell:')
    print(f'Get-Content "{state_file}" -Raw | Set-Clipboard')
    context.close()
