"""Find the exact Zoho Survey by its project title, without a survey URL.

The company name is only a search filter; the complete title must match before
opening a survey. Never fall back to the first or closest search result.
"""
import re
import time
import unicodedata


def normalize(value):
    value = unicodedata.normalize('NFKC', str(value or '')).casefold()
    return ' '.join(re.findall(r'\w+', value, flags=re.UNICODE))


def company_prefix(project):
    """The part before the first spaced dash in the supplied full project name."""
    first = re.split(r'\s+[-–—]\s+', project, maxsplit=1)[0].strip()
    if not first or len(first) < 2:
        raise RuntimeError('Cannot derive a company name from this project name.')
    return first


def visible(locator):
    try:
        return locator.count() > 0 and locator.first.is_visible()
    except Exception:
        return False


def survey_id(url):
    hit = re.search(r'/survey/(\d+)(?:/|$)', url)
    return hit.group(1) if hit else None



def open_department(page, department='Impact for L&D'):
    """Open Zoho department without waiting for every resource."""
    target = 'https://survey.zoho.com/survey/newui#/home'
    print('Opening Zoho department:', department, flush=True)

    try:
        page.goto(target, wait_until='commit', timeout=90000)
    except Exception as exc:
        if 'survey.zoho.com/survey/newui' not in page.url:
            raise RuntimeError(
                f'Zoho navigation failed: {str(exc)[:240]}'
            ) from exc

    page.locator('body').wait_for(
        state='visible', timeout=30000
    )

    if '/mysurveys' in page.url:
        return

    selector = page.get_by_text(department, exact=True)

    for _ in range(120):
        if '/mysurveys' in page.url:
            return
        if visible(selector):
            selector.first.click(timeout=6000)
            break
        page.wait_for_timeout(500)
    else:
        body = page.locator('body').inner_text(
            timeout=2500
        ).casefold()

        if any(x in body for x in (
            'sign in', 'sign-in', 'log in', 'login'
        )):
            raise RuntimeError(
                'Zoho session expired. Refresh the '
                'ZOHO_STORAGE_STATE_JSON secret.'
            )

        raise RuntimeError(
            f'Department {department!r} not found '
            'within 60 seconds.'
        )

    for _ in range(120):
        if '/mysurveys' in page.url:
            return
        page.wait_for_timeout(500)

    raise RuntimeError(
        'My Surveys page did not open after selecting '
        'the department.'
    )


def search_input(page):
    # Zoho's My Surveys search is an input in the main header (screenshots).
    for loc in [
        page.get_by_placeholder('Search', exact=False),
        page.locator('input[type="search"]'),
        page.locator('input[placeholder*="Search" i]'),
        page.locator('input[aria-label*="Search" i]'),
    ]:
        try:
            for i in range(min(loc.count(), 5)):
                candidate = loc.nth(i)
                if candidate.is_visible() and candidate.is_enabled():
                    return candidate
        except Exception:
            pass
    # The search control may initially be represented by an icon and expand on click.
    for loc in [page.get_by_text('Search', exact=True), page.get_by_role('button', name='Search')]:
        try:
            if visible(loc):
                loc.first.click(timeout=3000)
                page.wait_for_timeout(450)
                return search_input_expanded(page)
        except Exception:
            pass
    raise RuntimeError('Cannot find the Zoho My Surveys search input.')


def search_input_expanded(page):
    for loc in [page.get_by_placeholder('Search', exact=False), page.locator('input[type="search"]'), page.locator('input[placeholder*="Search" i]')]:
        for i in range(min(loc.count(), 5)):
            if loc.nth(i).is_visible():
                return loc.nth(i)
    raise RuntimeError('Zoho search input did not expand.')


def matching_title_nodes(page, project):
    """Find exact rendered text, not fuzzy fragments or the page's aggregate body text."""
    expected = normalize(project)
    # Results use a separate survey card title; limiting to leafish elements avoids
    # treating the entire results panel as a match.
    return page.locator('body *').evaluate_all('''(nodes, expected) => {
        const norm = s => (s || '').normalize('NFKC').toLocaleLowerCase()
            .replace(/[^\\p{L}\\p{N}_]+/gu, ' ').trim().replace(/\\s+/g,' ');
        return nodes.filter(el => {
            if (!el.getClientRects().length) return false;
            if (norm(el.textContent) !== expected) return false;
            return !Array.from(el.children).some(child => norm(child.textContent) === expected);
        }).slice(0, 20).map(el => ({tag: el.tagName, text: el.textContent.trim()}));
    }''', expected)


def locate_exact_title(page, project):
    # Matching via a regex with whitespace flexibility handles wrapped result titles.
    # Verify the normalized whole visible text, never substring match.
    pattern = re.compile(r'^\s*' + r'\s+'.join(re.escape(part) for part in project.split()) + r'\s*$', re.I)
    loc = page.get_by_text(pattern)
    candidates = []
    try:
        for idx in range(min(loc.count(), 30)):
            item = loc.nth(idx)
            if item.is_visible() and normalize(item.inner_text(timeout=1000)) == normalize(project):
                candidates.append(item)
    except Exception:
        pass
    if len(candidates) > 1:
        # Some apps duplicate the title (tooltip etc). Use the smallest visible text
        # element from the first result rather than clicking a fuzzy match.
        return candidates[0]
    if candidates:
        return candidates[0]
    return None


def find_survey_by_name(page, project, department='Impact for L&D', max_scrolls=70):
    """Search by company, scroll the results, then open only exact title match."""
    company = company_prefix(project)
    print(f'Opening Zoho department: {department}', flush=True)
    open_department(page, department)
    field = search_input(page)
    field.fill(company, timeout=5000)
    field.press('Enter', timeout=3000)
    print(f'Searching Zoho surveys for company: {company}', flush=True)
    page.wait_for_timeout(1200)
    # Explicit screenshot-like workflow: scroll results until exact title appears.
    seen_positions = set()
    unchanged = 0
    for turn in range(max_scrolls + 1):
        title = locate_exact_title(page, project)
        if title is not None:
            title.scroll_into_view_if_needed(timeout=5000)
            print(f'Exact project title found after {turn} scroll(s): {project}', flush=True)
            title.click(timeout=8000)
            for _ in range(30):
                if survey_id(page.url):
                    print(f'Opened matching Zoho survey ID: {survey_id(page.url)}', flush=True)
                    return page.url
                page.wait_for_timeout(500)
            raise RuntimeError('Clicked matching survey title, but Zoho did not open the survey dashboard.')
        # Scroll results panel, not a hardcoded coordinate. Mouse wheel bubbles to
        # whichever div (or page) currently owns scrolling in Zoho's UI.
        page.mouse.move(1200, 680)
        page.mouse.wheel(0, 670)
        page.wait_for_timeout(550)
        try:
            position = page.evaluate('''() => {
                const candidates = [document.scrollingElement, ...document.querySelectorAll('main, [class*="list"], [class*="scroll"], [class*="result"]')];
                return candidates.filter(x => x && x.scrollHeight > x.clientHeight + 30)
                    .map(x => [x.scrollTop, x.scrollHeight, x.clientHeight]).join('|');
            }''')
        except Exception:
            position = str(turn)
        if position in seen_positions:
            unchanged += 1
        else:
            unchanged = 0
            seen_positions.add(position)
        if turn and turn % 10 == 0:
            print(f'  Scanned {turn} result-scroll steps; looking for exact title...', flush=True)
        if unchanged >= 8:
            break
    raise RuntimeError(
        f'Could not locate exact survey title in {department!r} after searching for '
        f'{company!r}: {project!r}. No export was attempted. '
        'Check the spelling, department, and whether the survey is visible in Zoho.'
    )
