"""Resolve build paths once; keep wheel policy in the sdist's pyproject.toml."""
import os
from pathlib import Path
import tomllib

root = Path('source').resolve()
platform = {'Linux': 'linux', 'macOS': 'macos', 'Windows': 'windows'}[os.environ['RUNNER_OS']]
config = tomllib.loads((root / 'pyproject.toml').read_text())['tool']['cibuildwheel']
environment = config.get(platform, {}).get('environment', {}).copy()
project = '/project/source' if platform == 'linux' else root.as_posix()
environment['UV_BUILD_CONSTRAINT'] = f'{project}/tools/ci/constraints.txt'
environment['PIP_CONSTRAINT'] = f'{project}/tools/ci/constraints.txt'
if platform == 'macos':
    environment['CCACHE_BASEDIR'] = project
if platform == 'windows':
    toolchain = os.environ['CMAKE_TOOLCHAIN_FILE'].replace('\\', '/')
    installed = os.environ['VCPKG_INSTALLED_DIR']
    environment.update(CMAKE_BUILD_PARALLEL_LEVEL='4', CMAKE_TOOLCHAIN_FILE=toolchain,
                       CMAKE_ARGS=f'-DCMAKE_TOOLCHAIN_FILE={toolchain} -DVCPKG_TARGET_TRIPLET=x64-windows-static-md -DVCPKG_INSTALLED_DIR={installed} -DVCPKG_MANIFEST_INSTALL=OFF -DTURBOXL_STATIC_WINDOWS_DEPS=ON')
with open(os.environ['GITHUB_ENV'], 'a') as output:
    values = ' '.join(f'{key}="{value}"' for key, value in environment.items())
    print(f'CIBW_ENVIRONMENT_{platform.upper()}={values}', file=output)
    # cibuildwheel resolves a custom constraints path from the invocation directory.
    print(f'CIBW_DEPENDENCY_VERSIONS={root.as_posix()}/tools/ci/constraints.txt', file=output)
