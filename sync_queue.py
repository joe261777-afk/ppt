"""Private GitHub-backed command queue for the user's own local Zoho sync agent."""
import json
import re
import uuid
from datetime import datetime, timezone
from urllib.parse import urlparse

ALLOWED_HOSTS = ('survey.zoho.com', 'survey.zoho.eu', 'survey.zoho.in',
                 'survey.zoho.com.au', 'survey.zoho.jp', 'survey.zoho.com.cn')
STATUS = ('pending', 'processing', 'done', 'error', 'login_required')


def valid_url(url):
    if not url:
        return ''
    parsed = urlparse(url)
    if parsed.scheme != 'https' or parsed.hostname not in ALLOWED_HOSTS:
        raise ValueError('Survey URL must be an official https://survey.zoho.* address.')
    return url


def normalize_project(project):
    project = ' '.join(project.split()).strip()
    if not project or len(project) > 220 or any(ord(c) < 32 for c in project):
        raise ValueError('Enter a project name of 1–220 printable characters.')
    return project


def new_request(library, project, url=''):
    project = normalize_project(project)
    url = valid_url(url.strip())
    identifier = uuid.uuid4().hex
    data = {'id': identifier, 'project': project, 'url': url, 'status': 'pending',
            'created_at': datetime.now(timezone.utc).isoformat(), 'message': 'Waiting for the GitHub Actions cloud worker'}
    library.write(f'sync_requests/{identifier}.json', json.dumps(data).encode('utf-8'), f'Queue survey sync {identifier[:8]}')
    return data


def list_requests(library):
    entries = library.request('GET', 'sync_requests') or []
    out = []
    for row in entries:
        if row.get('type') != 'file' or not re.fullmatch(r'[a-f0-9]{32}\.json', row.get('name', '')):
            continue
        try:
            data = json.loads(library.read(row['path']))
            if data.get('id') == row['name'][:-5] and data.get('status') in STATUS:
                out.append(data)
        except (ValueError, TypeError, KeyError):
            pass
    return sorted(out, key=lambda x: x.get('created_at', ''), reverse=True)


def get_request(library, identifier):
    if not re.fullmatch(r'[a-f0-9]{32}', identifier):
        raise ValueError('Invalid request identifier')
    raw = library.read(f'sync_requests/{identifier}.json')
    return json.loads(raw) if raw else None


def update_request(library, data, status, message):
    if status not in STATUS:
        raise ValueError('Invalid status')
    updated = dict(data)
    updated['status'] = status
    updated['message'] = str(message)[:400]
    updated['updated_at'] = datetime.now(timezone.utc).isoformat()
    library.write(f"sync_requests/{updated['id']}.json", json.dumps(updated, ensure_ascii=False).encode(),
                  f"Sync status {updated['id'][:8]}: {status}")
    return updated
