"""Fail closed before a version tag can start release builds."""
import argparse
from pathlib import Path
import re
import subprocess

VERSION = re.compile(r'^project\(turboxl VERSION (\d+\.\d+\.\d+) LANGUAGES CXX\)$', re.M)


def project_version(root=Path('.')):
    matches = VERSION.findall((root / 'CMakeLists.txt').read_text())
    if len(matches) != 1:
        raise ValueError('Expected exactly one CMake project version')
    return matches[0]


def validate_tag(tag, version):
    if not re.fullmatch(r'v(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)', tag):
        raise ValueError(f'Invalid release tag: {tag}')
    if tag != f'v{version}':
        raise ValueError(f'Tag {tag} does not match CMake version {version}')


def validate_commit(tag):
    tagged = subprocess.check_output(['git', 'rev-parse', f'refs/tags/{tag}^{{commit}}'], text=True).strip()
    head = subprocess.check_output(['git', 'rev-parse', 'HEAD'], text=True).strip()
    if tagged != head:
        raise ValueError('Checkout does not match the release tag')
    subprocess.run(['git', 'merge-base', '--is-ancestor', tagged, 'origin/main'], check=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('tag')
    args = parser.parse_args()
    version = project_version()
    validate_tag(args.tag, version)
    validate_commit(args.tag)
    print(f'Validated {args.tag} on main history')
