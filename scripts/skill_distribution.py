#!/usr/bin/env python3
"""Incremental distribution index for a skill release.

A release ships a per-file manifest (name + sha256 + size) and an update index
whose history records, per version, which files were added, modified or removed.
An installed skill can then pull only what changed instead of a whole archive.
"""
import json
import hashlib
import shutil
from pathlib import Path

SCHEMA = 'skill-updates/v1'
MANIFEST_SCHEMA = 'skill-manifest/v1'
SKIP_DIRS = {'.venv', '__pycache__', '.git', '.github'}
SKIP_NAMES = {'.DS_Store', '.gitignore'}
SKIP_SUFFIXES = ('.pyc',)


def digest_bytes(data):
    return hashlib.sha256(data).hexdigest()


def shipped_files(root):
    """Relative POSIX paths of files that belong in a release, sorted."""
    root = Path(root)
    for path in sorted(root.rglob('*')):
        if not path.is_file():
            continue
        relative = path.relative_to(root)
        if (SKIP_DIRS & set(relative.parts) or relative.name in SKIP_NAMES
                or relative.name.endswith(SKIP_SUFFIXES)):
            continue
        yield relative.as_posix()


def build_manifest(root):
    """Per-file sha256 and size, for verifying a downloaded or installed copy."""
    root = Path(root)
    files = {}
    for relative in shipped_files(root):
        data = (root / relative).read_bytes()
        files[relative] = {'sha256': digest_bytes(data), 'size': len(data)}
    return {'schema': MANIFEST_SCHEMA, 'files': files}


def delta(old, new):
    """Compare two {path: content} snapshots into added / modified / removed lists."""
    return {
        'added': sorted(set(new) - set(old)),
        'modified': sorted(path for path in set(old) & set(new) if old[path] != new[path]),
        'removed': sorted(set(old) - set(new)),
    }


def parse_version(value):
    parts = [int(part) for part in str(value).strip().split('.')]
    return tuple(parts + [0] * (3 - len(parts)))[:3]


def plan_update(index, local_version):
    """Files to add, replace or delete when moving from local_version to the latest.

    Replays every release newer than local in version order, tracking each path's
    final state. A path added then removed ends up removed; removed then re-added
    ends up added. A naive set union gets both cases wrong.
    """
    target = index.get('latest_version')
    if not target or parse_version(target) <= parse_version(local_version):
        return {'needs_update': False, 'target_version': local_version,
                'added': [], 'modified': [], 'removed': []}
    state = {}
    later = sorted((entry for entry in index['history']
                    if parse_version(entry['version']) > parse_version(local_version)),
                   key=lambda entry: parse_version(entry['version']))
    for entry in later:
        for name in entry.get('added', []):
            state[name] = 'added'
        for name in entry.get('modified', []):
            state[name] = 'modified'
        for name in entry.get('removed', []):
            state[name] = 'removed'
    def names(status):
        return sorted(name for name, value in state.items() if value == status)
    return {'needs_update': True, 'target_version': target,
            'added': names('added'), 'modified': names('modified'), 'removed': names('removed')}


def new_index(source):
    """An empty update index pointing at the mirror an installed skill should read."""
    return {'schema': SCHEMA, 'latest_version': None, 'update_source': source, 'history': []}


def record_release(index, version, changes, summary):
    """Record one version's file delta, replacing any entry for the same version."""
    entry = {'version': version, 'summary': summary,
             'added': sorted(changes['added']), 'modified': sorted(changes['modified']),
             'removed': sorted(changes['removed'])}
    index['history'] = [h for h in index['history'] if h['version'] != version] + [entry]
    index['latest_version'] = version
    return index


def apply_update(root, plan, fetch, remote_manifest, version):
    """Write the changed files, prune removed ones, then report the new version.

    Every download is verified against the remote manifest before anything
    touches the install, so a tampered or truncated mirror leaves no half state.
    VERSION is written by the release itself when it is part of the delta; it is
    never stamped separately, or it would disagree with its own manifest hash.
    """
    root = Path(root)
    wanted_hashes = remote_manifest.get('files', {})
    staged, failures = {}, []
    for name in plan['added'] + plan['modified']:
        expected = wanted_hashes.get(name, {}).get('sha256')
        if not expected:
            failures.append(f'{name}: 缺少 manifest 条目，无法确认完整性')
            continue
        data = fetch(name)
        actual = digest_bytes(data)
        if expected != actual:
            failures.append(f'{name}: sha256 {actual} != {expected}')
        else:
            staged[name] = data
    if failures:
        raise ValueError('下载的文件未通过校验：\n' + '\n'.join(failures))
    for name, data in staged.items():
        destination = root / name
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(data)
    for name in plan['removed']:
        target = root / name
        if target.exists():
            target.unlink()
    return {'version': version, 'written': sorted(staged), 'removed': sorted(plan['removed'])}


def dump(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True) + '\n',
                    encoding='utf-8')


def build_release(source, out, version, summary, index):
    """Publish a manifest and the next revision of the update index.

    The delta comes from the previous manifest in ``out``, so the first release
    reports every shipped file as added and later ones stay minimal. Only the
    files in that delta are mirrored under ``files/``, which is what an installed
    copy downloads.
    """
    source, out = Path(source), Path(out)
    manifest = build_manifest(source)
    out.mkdir(parents=True, exist_ok=True)
    previous_path = out / 'manifest.json'
    if previous_path.is_file():
        previous = json.loads(previous_path.read_text(encoding='utf-8')).get('files', {})
    else:
        previous = {}
    changes = delta({name: entry['sha256'] for name, entry in previous.items()},
                    {name: entry['sha256'] for name, entry in manifest['files'].items()})
    record_release(index, version, changes, summary)
    for name in changes['added'] + changes['modified']:
        destination = out / 'files' / name
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source / name, destination)
    for name in changes['removed']:
        stale = out / 'files' / name
        if stale.exists():
            stale.unlink()
    dump(out / 'manifest.json', manifest)
    dump(out / 'updates.json', index)
    return {'version': version, **{k: len(v) for k, v in changes.items()}}


def publish(source, out, version, summary, source_url):
    """Build the mirror an installed skill can pull from: manifest, index, files.

    An existing index keeps its recorded ``update_source`` unless a new one is
    given, so republishing does not silently point installs at the wrong mirror.
    """
    index_path = Path(out) / 'updates.json'
    index = (json.loads(index_path.read_text(encoding='utf-8')) if index_path.is_file()
             else new_index(source_url))
    if source_url:
        index['update_source'] = source_url
    return build_release(source, out, version, summary, index)


def check(root, index):
    """Report what an installed copy is missing, without changing it."""
    version = (Path(root) / 'VERSION').read_text(encoding='utf-8').strip()
    return {'local_version': version, **plan_update(index, version)}


def update(root, index, fetch, remote_manifest):
    """Apply the plan for an installed copy against a remote mirror."""
    version = (Path(root) / 'VERSION').read_text(encoding='utf-8').strip()
    plan = plan_update(index, version)
    if not plan['needs_update']:
        return {'version': version, 'written': [], 'removed': [], 'up_to_date': True}
    return {**apply_update(root, plan, fetch, remote_manifest, plan['target_version']),
            'up_to_date': False}


def fetch_http(source):
    """A fetch(name) callable plus manifest loader for a plain HTTP mirror."""
    import urllib.request

    def fetch(name):
        url = source.rstrip('/') + '/files/' + name
        with urllib.request.urlopen(url, timeout=120) as response:
            return response.read()

    with urllib.request.urlopen(source.rstrip('/') + '/manifest.json', timeout=30) as response:
        manifest = json.loads(response.read().decode('utf-8'))
    return fetch, manifest


def main():
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    p = sub.add_parser('publish')
    p.add_argument('--source', required=True)
    p.add_argument('--out', required=True)
    p.add_argument('--version', required=True)
    p.add_argument('--summary', required=True)
    p.add_argument('--source-url', required=True)
    p = sub.add_parser('check')
    p.add_argument('--root', required=True)
    p.add_argument('--index', required=True)
    p = sub.add_parser('update')
    p.add_argument('--root', required=True)
    p.add_argument('--index', required=True)
    args = parser.parse_args()
    try:
        if args.command == 'publish':
            print(json.dumps(publish(args.source, args.out, args.version,
                                     args.summary, args.source_url), ensure_ascii=False))
            return 0
        index = json.loads(Path(args.index).read_text(encoding='utf-8'))
        if args.command == 'check':
            print(json.dumps(check(args.root, index), ensure_ascii=False))
            return 0
        fetch, manifest = fetch_http(index['update_source'])
        print(json.dumps(update(args.root, index, fetch, manifest), ensure_ascii=False))
        return 0
    except (ValueError, OSError, KeyError) as error:
        parser.exit(2, str(error) + '\n')


if __name__ == '__main__':
    raise SystemExit(main())
