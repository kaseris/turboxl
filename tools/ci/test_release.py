import hashlib
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
