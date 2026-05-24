from pathlib import Path

path = Path('main.py')
text = path.read_text(encoding='utf-8')

if 'import urllib.request\n' not in text:
    anchor = 'import webbrowser\n'
    if anchor not in text:
        raise SystemExit('webbrowser import anchor not found')
    text = text.replace(anchor, anchor + 'import urllib.request\n', 1)

if 'LATEST_RELEASE_API' not in text:
    anchor = 'CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".chungbuk_auto_config.json")\n'
    if anchor not in text:
        raise SystemExit('CONFIG_FILE anchor not found')
    text = text.replace(
        anchor,
        anchor + "LATEST_RELEASE_API = 'https://api.github.com/repos/codersongpro/sotong/releases/latest'\n",
        1,
    )

if 'def _version_parts(version: str) -> tuple:' not in text:
    helper = '''

def _version_parts(version: str) -> tuple:
    version = (version or '').strip().lstrip('vV')
    parts = []
    for part in version.split('.'):
        digits = ''.join(ch for ch in part if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts or [0])


def _is_newer_version(latest: str, current: str) -> bool:
    latest_parts = _version_parts(latest)
    current_parts = _version_parts(current)
    size = max(len(latest_parts), len(current_parts))
    latest_parts += (0,) * (size - len(latest_parts))
    current_parts += (0,) * (size - len(current_parts))
    return latest_parts > current_parts
'''
    anchor = '\nclass App:\n'
    if anchor not in text:
        raise SystemExit('class App anchor not found')
    text = text.replace(anchor, helper + anchor, 1)

init_old = '        self._refresh_calib_labels()\n        self._check_deps()\n'
init_new = '        self._refresh_calib_labels()\n        self._check_deps()\n        self._check_for_update_async()\n'
if 'self._check_for_update_async()' not in text:
    if init_old not in text:
        raise SystemExit('App __init__ anchor not found')
    text = text.replace(init_old, init_new, 1)

if 'def _check_for_update_async(self):' not in text:
    methods = '''
    # update check
    def _check_for_update_async(self):
        threading.Thread(target=self._check_for_update_worker, daemon=True).start()

    def _check_for_update_worker(self):
        try:
            req = urllib.request.Request(
                LATEST_RELEASE_API,
                headers={'User-Agent': f'{APP_NAME}/{APP_VERSION}'}
            )
            with urllib.request.urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read().decode('utf-8'))
            latest = (data.get('tag_name') or data.get('name') or '').strip().lstrip('vV')
            url = data.get('html_url') or 'https://github.com/codersongpro/sotong/releases/latest'
            if latest and _is_newer_version(latest, APP_VERSION):
                self.root.after(0, lambda: self._show_update_notice(latest, url))
        except Exception:
            pass

    def _show_update_notice(self, latest: str, url: str):
        open_page = messagebox.askyesno(
            '새 버전 알림',
            f'{APP_NAME} 새 버전이 나왔습니다.\\n\\n'
            f'현재 버전: {APP_VERSION}\\n'
            f'최신 버전: {latest}\\n\\n'
            '다운로드 페이지를 열까요?'
        )
        if open_page:
            webbrowser.open(url)

'''
    anchor = '    def _check_deps(self):\n'
    if anchor not in text:
        raise SystemExit('_check_deps anchor not found')
    text = text.replace(anchor, methods + anchor, 1)

path.write_text(text, encoding='utf-8')
