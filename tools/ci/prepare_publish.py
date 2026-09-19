"""Retain only unpublished files after verifying already-published SHA256s."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import urllib.error
import urllib.request
from validate_release import project_version


def pending_files(files, published):
    local = {p.name: p for p in files}
    if set(published) - set(local):
        raise ValueError('PyPI contains unexpected files for this version')
    for name, digest in published.items():
        if hashlib.sha256(local[name].read_bytes()).hexdigest() != digest:
            raise ValueError(f'Published file differs: {name}; never replace a released version')
    return [p for name, p in local.items() if name not in published]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('directory', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    version = project_version()
    request = urllib.request.Request(f'https://pypi.org/pypi/turboxl/{version}/json', headers={'User-Agent': 'turboxl-release'})
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            published = {p['filename']: p['digests']['sha256'] for p in json.load(response)['urls']}
    except urllib.error.HTTPError as error:
        if error.code != 404:
            raise
        published = {}
    files = sorted([*args.directory.glob('*.whl'), *args.directory.glob('*.tar.gz')])
    pending = pending_files(files, published)
    args.output.mkdir(exist_ok=False)
    for path in pending:
        shutil.copy2(path, args.output / path.name)
    if os.environ.get('GITHUB_OUTPUT'):
        with open(os.environ['GITHUB_OUTPUT'], 'a') as output:
            print(f'count={len(pending)}', file=output)
    print(f'{len(pending)} new files; {len(published)} matching published files')


if __name__ == '__main__':
    main()
