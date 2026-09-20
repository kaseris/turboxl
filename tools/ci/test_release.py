import hashlib
import io
import tarfile
import zipfile
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch
from validate_release import validate_tag, validate_commit
from validate_dist import validate, platform_id
from prepare_publish import pending_files
from packaging.tags import Tag


class ReleaseTests(unittest.TestCase):
    def test_tag(self):
        validate_tag('v1.2.3', '1.2.3')
        for tag in ['1.2.3', 'v1.2', 'v1.2.3rc1', 'v01.2.3', 'v1.2.4', 'v1.2.3\n']:
            with self.subTest(tag=tag), self.assertRaises(ValueError):
                validate_tag(tag, '1.2.3')

    def test_ancestry(self):
        with patch('subprocess.check_output', return_value='abc\n'), patch('subprocess.run', side_effect=subprocess.CalledProcessError(1, 'git')):
            with self.assertRaises(subprocess.CalledProcessError):
                validate_commit('v1.2.3')
        with patch('subprocess.check_output', side_effect=['abc\n', 'def\n']):
            with self.assertRaises(ValueError):
                validate_commit('v1.2.3')

    def test_missing_wheels(self):
        with tempfile.TemporaryDirectory() as tmp, self.assertRaisesRegex(ValueError, 'inventory'):
            validate(Path(tmp), '1.2.3')

    def test_platforms(self):
        self.assertEqual(platform_id({Tag('cp312', 'abi3', 'win_amd64')}), 'windows-x64')
        for platform in ['win32', 'macosx_16_0_arm64', 'linux_x86_64']:
            with self.assertRaises(ValueError):
                platform_id({Tag('cp312', 'abi3', platform)})

    def test_complete_inventory_and_corrupt_metadata(self):
        version = '1.2.3'
        meta = b'Name: turboxl\nVersion: 1.2.3\nRequires-Python: >=3.10\n'
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            wheels = []
            for platform in ['manylinux_2_28_x86_64', 'win_amd64', 'macosx_15_0_arm64', 'macosx_15_0_x86_64']:
                for python, abi in [('cp310', 'cp310'), ('cp311', 'cp311'), ('cp312', 'abi3')]:
                    path = directory / f'turboxl-{version}-{python}-{abi}-{platform}.whl'
                    with zipfile.ZipFile(path, 'w') as wheel:
                        extension = 'turboxl/_turboxl.pyd' if platform == 'win_amd64' else 'turboxl/_turboxl.so'
                        wheel.writestr(extension, b'fixture')
                        wheel.writestr('turboxl/__init__.py', b'from ._turboxl import *\n')
                        wheel.writestr(f'turboxl-{version}.dist-info/METADATA', meta)
                        wheel.writestr(f'turboxl-{version}.dist-info/licenses/LICENSE', b'MIT')
                    wheels.append(path)
            required = ['CMakeLists.txt', 'pyproject.toml', 'README.md', 'LICENSE', 'PKG-INFO',
                        'src/python/module.cpp', 'src/python/turboxl/__init__.py',
                        'include/xlsxcsv.hpp', 'cmake/turboxlConfig.cmake.in',
                        'tools/wheels/fixtures.py', 'tools/wheels/smoke_test.py', 'tools/ci/constraints.txt',
                        'tools/ci/install_deps.sh', 'tests/fixture_config.hpp.in', 'tests/fixture_helpers.hpp']
            with tarfile.open(directory / f'turboxl-{version}.tar.gz', 'w:gz') as archive:
                for name in required:
                    data = meta if name == 'PKG-INFO' else b'fixture'
                    entry = tarfile.TarInfo(f'turboxl-{version}/{name}')
                    entry.size = len(data)
                    archive.addfile(entry, io.BytesIO(data))
            validate(directory, version)
            with zipfile.ZipFile(wheels[0], 'a') as wheel:
                wheel.writestr('include/native.hpp', b'not a runtime dependency')
            with self.assertRaisesRegex(ValueError, 'development files'):
                validate(directory, version)
            wheels[0].unlink()
            with self.assertRaisesRegex(ValueError, 'inventory'):
                validate(directory, version)

    def test_publish_recovery(self):
        with tempfile.TemporaryDirectory() as tmp:
            a, b = Path(tmp) / 'a.whl', Path(tmp) / 'b.whl'
            a.write_bytes(b'first'); b.write_bytes(b'second')
            digest = hashlib.sha256(a.read_bytes()).hexdigest()
            self.assertEqual(pending_files([a, b], {a.name: digest}), [b])
            self.assertEqual(pending_files([a], {a.name: digest}), [])
            for published in [{a.name: 'wrong'}, {'unexpected.whl': digest}]:
                with self.assertRaises(ValueError):
                    pending_files([a, b], published)


if __name__ == '__main__':
    unittest.main()
