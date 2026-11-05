"""Private GitHub-backed file library. Keep GitHub credentials only in secrets/env."""
import base64
import json
import re
from urllib.parse import quote
from pathlib import Path
import requests

BASE = 'https://api.github.com'


def slug(text):
    result = re.sub(r'[^a-z0-9]+', '-', text.casefold()).strip('-')
    return result[:90] or 'survey'


class Library:
    def __init__(self, repo, token):
        if not re.fullmatch(r'[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+', repo):
            raise ValueError('Repository must be owner/repository')
        if not token:
            raise ValueError('Missing GitHub token')
        self.repo = repo
        self.session = requests.Session()
        self.session.headers.update({'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github+json',
                                     'X-GitHub-Api-Version': '2022-11-28', 'User-Agent': 'workshop-evaluation-sync'})

    def request(self, method, path, **kwargs):
        path = quote(path, safe='/')
        r = self.session.request(method, f'{BASE}/repos/{self.repo}/contents/{path}', timeout=75, **kwargs)
        if r.status_code == 404:
            return None
        if not r.ok:
            raise RuntimeError(f'GitHub {r.status_code}: {r.text[:350]}')
        return r.json()

    def read(self, path):
        obj = self.request('GET', path)
        if obj is None: return None
        if isinstance(obj, list) or obj.get('type') != 'file':
            raise ValueError(f'Not a file: {path}')
        # The GitHub Contents API sometimes omits content for larger files.
        if obj.get('encoding') == 'base64' and obj.get('content'):
            return base64.b64decode(obj['content'])
        rr = self.session.get(obj['download_url'], timeout=90)
        rr.raise_for_status()
        return rr.content

    def write(self, path, data, message):
        old = self.request('GET', path)
        body = {'message': message, 'content': base64.b64encode(data).decode('ascii')}
        if old and old.get('sha'):
            body['sha'] = old['sha']
        return self.request('PUT', path, json=body)

    def list_projects(self):
        directories = self.request('GET', 'zoho_exports') or []
        result = []
        for entry in directories:
            if entry.get('type') != 'dir': continue
            raw = self.read(f"{entry['path']}/manifest.json")
            if not raw: continue
            try:
                m = json.loads(raw)
                if m.get('project') and m.get('pdf') and m.get('excel'):
                    m['folder'] = entry['path']
                    result.append(m)
            except (ValueError, KeyError):
                pass
        return sorted(result, key=lambda x: x['project'].casefold())

    def save_project(self, project, pdf_path, excel_path):
        s = slug(project)
        root = f'zoho_exports/{s}'
        pdf, excel = Path(pdf_path), Path(excel_path)
        if pdf.suffix.lower() != '.pdf' or excel.suffix.lower() not in ('.xlsx', '.xlsm'):
            raise ValueError('Expected one PDF and one XLSX/XLSM file')
        if pdf.stat().st_size > 95_000_000 or excel.stat().st_size > 95_000_000:
            raise ValueError('A file exceeds the 100 MB GitHub file limit.')
        self.write(f'{root}/survey_report.pdf', pdf.read_bytes(), f'Update PDF for {project}')
        self.write(f'{root}/survey_responses.xlsx', excel.read_bytes(), f'Update responses for {project}')
        manifest = {'project': project, 'pdf': 'survey_report.pdf', 'excel': 'survey_responses.xlsx'}
        self.write(f'{root}/manifest.json', json.dumps(manifest, ensure_ascii=False, indent=2).encode(), f'Index {project}')
        return manifest

    def load_project(self, manifest):
        root = manifest['folder']
        return self.read(f"{root}/{manifest['pdf']}"), self.read(f"{root}/{manifest['excel']}")
