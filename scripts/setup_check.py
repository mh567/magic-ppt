#!/usr/bin/env python3
"""Check local dependencies; optionally install missing packages in a skill-local venv."""
import argparse
import importlib.util
import json
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PACKAGES = {'pptx': 'python-pptx', 'PIL': 'Pillow', 'lxml': 'lxml',
            'docx': 'python-docx', 'openpyxl': 'openpyxl', 'pypdf': 'pypdf'}
REQUIRED = {'pptx', 'PIL', 'lxml'}


def inspect(python):
    code = 'import importlib.util,json; print(json.dumps({k:importlib.util.find_spec(k) is not None for k in ' + repr(list(PACKAGES)) + '}))'
    result = subprocess.run([str(python), '-c', code], capture_output=True, text=True, check=True)
    return json.loads(result.stdout)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--install', action='store_true', help='Install missing packages into .venv; requires network')
    args = parser.parse_args()
    python = ROOT / '.venv' / ('Scripts/python.exe' if sys.platform == 'win32' else 'bin/python')
    if not python.is_file():
        python = Path(sys.executable)
    try:
        found = inspect(python)
        if args.install and not all(found.values()):
            local = ROOT / '.venv' / ('Scripts/python.exe' if sys.platform == 'win32' else 'bin/python')
            if not local.is_file():
                subprocess.run([sys.executable, '-m', 'venv', str(ROOT / '.venv')], check=True)
            python = local
            found = inspect(python)
            missing = [package for module, package in PACKAGES.items() if not found[module]]
            if missing:
                subprocess.run([str(python), '-m', 'pip', 'install', *missing], check=True)
            found = inspect(python)
        missing = [PACKAGES[k] for k in REQUIRED if not found[k]]
        report = {'status': 'ready' if not missing else 'missing-dependencies',
                  'python': str(python), 'packages': found,
                  'missing_required': missing,
                  'missing_optional': [v for k, v in PACKAGES.items() if k not in REQUIRED and not found[k]],
                  'local_renderer': bool(shutil.which('soffice') and shutil.which('pdftoppm')),
                  'host_checks': 'Agent must check image tools, final renderer and template fonts before production'}
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 1 if missing else 0
    except (OSError, subprocess.SubprocessError, ValueError) as error:
        print(json.dumps({'status': 'failed', 'error': str(error)}, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    sys.exit(main())
