#!/usr/bin/env python3
"""Validate a clean committed revision and build a reproducible offline package."""
import argparse
import hashlib
import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def run(*args):
    return subprocess.check_output(args, cwd=ROOT, text=True).strip()


def build(out, tag=None):
    if run('git', 'status', '--porcelain', '--untracked-files=normal'):
        raise ValueError('Commit or preserve local changes before building a release')
    version = (ROOT / 'VERSION').read_text().strip()
    if not re.fullmatch(r'0|[1-9]\d*', version.split('.')[0]) or not re.fullmatch(r'(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)', version):
        raise ValueError('VERSION must contain a semantic version')
    if tag is not None and tag != 'v' + version:
        raise ValueError('Release tag must match VERSION')
    skill = (ROOT / 'SKILL.md').read_text()
    if not re.search(r'^name: magic-ppt$', skill, re.M) or not re.search(r'^  version: "' + re.escape(version) + r'"$', skill, re.M):
        raise ValueError('Skill name/version mismatch')
    changelog = (ROOT / 'CHANGELOG.md').read_text()
    match = re.search(r'^## ' + re.escape(version) + r' \(\d{4}-\d{2}-\d{2}\)\n(.*?)(?=^## |\Z)', changelog, re.M | re.S)
    if not match or not match.group(1).strip():
        raise ValueError('Versioned changelog entry required')
    subprocess.run([sys.executable, str(ROOT / 'scripts/self_test.py')], cwd=ROOT, check=True)
    if run('git', 'status', '--porcelain', '--untracked-files=normal'):
        raise ValueError('Checks modified the checkout')
    commit = run('git', 'rev-parse', 'HEAD')
    entries = subprocess.check_output(['git', 'ls-tree', '-r', '-z', 'HEAD'], cwd=ROOT).split(b'\0')
    out = Path(out).resolve()
    if out == ROOT or ROOT in out.parents:
        raise ValueError('Output directory must be outside the checkout')
    out.mkdir(parents=True, exist_ok=True)
    package = out / ('magic-ppt-' + version + '.zip')
    payloads = {}
    for entry in entries:
        if not entry:
            continue
        meta, raw_path = entry.split(b'\t', 1)
        mode, kind, oid = meta.decode().split()
        path = raw_path.decode()
        if mode not in {'100644', '100755'} or kind != 'blob':
            raise ValueError('Unsupported tracked entry: ' + path)
        if path.startswith('.github/') or path == '.gitignore':
            continue
        if any(part in {'.git', '.venv', '__pycache__', '..'} for part in Path(path).parts) or path.endswith('.pyc'):
            raise ValueError('Unexpected package file: ' + path)
        payloads[path] = (mode, subprocess.check_output(['git', 'cat-file', 'blob', oid], cwd=ROOT))
    with zipfile.ZipFile(package, 'w', zipfile.ZIP_DEFLATED) as archive:
        for path, (mode, data) in sorted(payloads.items()):
            info = zipfile.ZipInfo('magic-ppt/' + path, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = (int(mode, 8) << 16)
            archive.writestr(info, data)
    with zipfile.ZipFile(package) as archive:
        if archive.testzip() is not None:
            raise ValueError('ZIP integrity failure')
        for path, (_, data) in payloads.items():
            if archive.read('magic-ppt/' + path) != data:
                raise ValueError('ZIP content mismatch: ' + path)
    sha = hashlib.sha256(package.read_bytes()).hexdigest()
    (out / 'SHA256SUMS.txt').write_text(sha + '  ' + package.name + '\n')
    receipt = {'version': version, 'commit': commit, 'package': package.name,
               'sha256': sha, 'files': len(payloads), 'checks': 'passed'}
    (out / 'release-manifest.json').write_text(json.dumps(receipt, indent=2) + '\n')
    (out / 'release-notes.md').write_text(match.group(1).strip() + '\n')
    print(json.dumps(receipt))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--out', required=True)
    parser.add_argument('--tag')
    args = parser.parse_args()
    try:
        build(args.out, args.tag)
    except (ValueError, OSError, subprocess.SubprocessError) as error:
        parser.exit(1, str(error) + '\n')


if __name__ == '__main__':
    main()
