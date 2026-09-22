"""Check the exact release inventory, archive contents, and package metadata."""
import argparse
from email.parser import BytesParser
import hashlib
from pathlib import Path
import tarfile
import zipfile
from packaging.utils import parse_wheel_filename
from packaging.version import Version
from validate_release import project_version

PLATFORMS = {'linux-x64', 'windows-x64', 'macos-arm64', 'macos-x64'}
VARIANTS = {('cp310', 'cp310'), ('cp311', 'cp311'), ('cp312', 'abi3')}


def platform_id(tags):
    names = {t.platform for t in tags}
    if all(p.startswith('manylinux_') and p.endswith('_x86_64') for p in names) and 'manylinux_2_28_x86_64' in names:
        return 'linux-x64'
    if names == {'win_amd64'}:
        return 'windows-x64'
    for arch, key in [('arm64', 'macos-arm64'), ('x86_64', 'macos-x64')]:
        if names == {f'macosx_15_0_{arch}'}:
            return key
    raise ValueError(f'Unexpected platform tags: {names}')


def metadata(data, version):
    fields = BytesParser().parsebytes(data)
    if fields['Name'] != 'turboxl' or fields['Version'] != version or fields['Requires-Python'] != '>=3.10':
        raise ValueError(f'Unexpected metadata: {fields["Name"]}, {fields["Version"]}, {fields["Requires-Python"]}')


def validate(directory, version, platform=None):
    seen = set()
    for path in sorted(directory.glob('*.whl')):
        name, wheel_version, _, tags = parse_wheel_filename(path.name)
        if name != 'turboxl' or wheel_version != Version(version):
            raise ValueError(f'Unexpected wheel version: {path.name}')
        variants = {(t.interpreter, t.abi) for t in tags}
        if len(variants) != 1 or not variants <= VARIANTS:
            raise ValueError(f'Unexpected wheel ABI: {path.name}')
        key = (platform_id(tags), next(iter(variants)))
        if key in seen:
            raise ValueError(f'Duplicate wheel: {key}')
        seen.add(key)
        with zipfile.ZipFile(path) as wheel:
            names = wheel.namelist()
            meta = [n for n in names if n.endswith('.dist-info/METADATA')]
            if len(meta) != 1:
                raise ValueError('Missing or duplicate wheel metadata')
            metadata(wheel.read(meta[0]), version)
            if any(n.startswith(('include/', 'lib/', 'bin/')) or n.endswith(('.a', '.lib', '.h', '.hpp', '.cmake')) for n in names):
                raise ValueError(f'Native development files in {path.name}')
            if not {'turboxl/__init__.py', 'turboxl/pandas.py'} <= set(names):
                raise ValueError(f'Missing Python package or pandas adapter in {path.name}')
            if not any(n.startswith('turboxl/_turboxl') and n.endswith(('.so', '.pyd')) for n in names):
                raise ValueError(f'Missing extension in {path.name}')
            if not any('license' in n.lower() for n in names):
                raise ValueError(f'Missing license in {path.name}')
    expected = {(p, v) for p in ({platform} if platform else PLATFORMS) for v in VARIANTS}
    if seen != expected:
        raise ValueError(f'Wheel inventory mismatch: missing={expected - seen}, extra={seen - expected}')
    if not platform:
        sdists = list(directory.glob('*.tar.gz'))
        if len(sdists) != 1 or sdists[0].name != f'turboxl-{version}.tar.gz':
            raise ValueError('Expected exactly one version-matched sdist')
        with tarfile.open(sdists[0]) as archive:
            prefix = f'turboxl-{version}/'
            names = {n.removeprefix(prefix) for n in archive.getnames()}
            required = {'CMakeLists.txt', 'pyproject.toml', 'README.md', 'LICENSE', 'PKG-INFO',
                        'src/python/module.cpp', 'src/python/turboxl/__init__.py',
                        'src/python/turboxl/pandas.py', 'tools/wheels/test_pandas_engine.py',
                        'include/xlsxcsv.hpp', 'cmake/turboxlConfig.cmake.in',
                        'tools/wheels/fixtures.py', 'tools/wheels/smoke_test.py', 'tools/ci/constraints.txt',
                        'tools/ci/install_deps.sh', 'tests/fixture_config.hpp.in', 'tests/fixture_helpers.hpp'}
            if not required <= names:
                raise ValueError(f'Incomplete sdist: {required - names}')
            if any(n.startswith(('.git/', '.venv/', 'build/')) or
                   '/__pycache__/' in n or n.endswith('.pyc') for n in names):
                raise ValueError('Local state leaked into sdist')
            metadata(archive.extractfile(prefix + 'PKG-INFO').read(), version)
    print(f'Validated {len(seen)} wheels' + (' and sdist' if not platform else ''))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('directory', type=Path)
    parser.add_argument('--platform', choices=sorted(PLATFORMS))
    parser.add_argument('--version')
    args = parser.parse_args()
    validate(args.directory, args.version or project_version(), args.platform)
    if not args.platform:
        files = sorted([*args.directory.glob('*.whl'), *args.directory.glob('*.tar.gz')])
        (args.directory / 'SHA256SUMS').write_text(''.join(f'{hashlib.sha256(p.read_bytes()).hexdigest()}  {p.name}\n' for p in files))
