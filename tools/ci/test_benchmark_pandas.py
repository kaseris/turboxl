"""Integrity and gate checks for the pandas benchmark."""

import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest import mock
from types import SimpleNamespace
import zipfile

BENCHMARK = Path(__file__).resolve().parents[1] / "benchmark_pandas.py"
SPEC = importlib.util.spec_from_file_location("benchmark_pandas", BENCHMARK)
benchmark = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(benchmark)


class BenchmarkPandasTests(unittest.TestCase):
    def test_manifest_rejects_changed_input_and_duplicate_ids(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = root / "sample.xlsx"
            workbook.write_bytes(b"original")
            entry = {
                "id": "sample",
                "path": workbook.name,
                "sha256": hashlib.sha256(b"original").hexdigest(),
                "source": "https://example.org/sample",
                "publisher": "example",
                "license": "CC BY 4.0",
                "kind": "real",
            }
            manifest = root / "manifest.json"
            manifest.write_text(json.dumps({"workbooks": [entry]}))
            self.assertEqual(benchmark.validate_manifest(manifest)[0]["id"], "sample")
            workbook.write_bytes(b"changed")
            with self.assertRaisesRegex(ValueError, "SHA256 mismatch"):
                benchmark.validate_manifest(manifest)
            workbook.write_bytes(b"original")
            manifest.write_text(json.dumps({"workbooks": [entry, entry]}))
            with self.assertRaisesRegex(ValueError, "duplicate"):
                benchmark.validate_manifest(manifest)

    def test_real_gate_requires_parity_and_threshold(self):
        records = [
            {
                "kind": "real",
                "publisher": f"publisher-{index % 3}",
                "advantage": 0.19 if index % 2 else 0.21,
                "parity": True,
            }
            for index in range(12)
        ]
        records.append({"kind": "synthetic", "advantage": -1, "parity": False})
        self.assertEqual(benchmark.performance_gate(records), (0.2, True))
        records[1]["parity"] = False
        self.assertEqual(benchmark.performance_gate(records), (0.2, False))
        self.assertFalse(benchmark.performance_gate(records[:11])[1])
        self.assertEqual(benchmark.performance_gate([]), (None, False))

    def test_worker_failure_contains_workbook_and_engine(self):
        result = SimpleNamespace(returncode=1, stderr="bad archive", stdout="")
        with mock.patch.object(benchmark.subprocess, "run", return_value=result):
            with self.assertRaisesRegex(
                RuntimeError, "sample turboxl worker failed: bad archive"
            ):
                benchmark.run_worker(
                    SimpleNamespace(python="python"),
                    {"id": "sample"},
                    "turboxl",
                    Path("unused.pickle"),
                )

    def test_azure_bundle_preserves_frozen_inputs(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = root / "sample.xlsx"
            workbook.write_bytes(b"frozen workbook")
            manifest = root / "manifest.json"
            manifest.write_text(
                json.dumps(
                    {
                        "workbooks": [
                            {
                                "id": "sample",
                                "path": "sample.xlsx",
                                "sha256": hashlib.sha256(
                                    workbook.read_bytes()
                                ).hexdigest(),
                                "source": "https://example.org/sample.xlsx",
                                "publisher": "example",
                                "license": "CC BY 4.0",
                                "kind": "real",
                            }
                        ]
                    }
                )
            )
            wheel = root / "turboxl-0.3.0-cp312-abi3-manylinux_2_28_x86_64.whl"
            with zipfile.ZipFile(wheel, "w") as archive:
                archive.writestr(
                    "turboxl-0.3.0.dist-info/WHEEL",
                    "Tag: cp312-abi3-manylinux_2_28_x86_64\n",
                )
            bundle = root / "bundle.zip"
            script = BENCHMARK.parent / "ci/prepare_pandas_azure.py"
            subprocess.run(
                [
                    sys.executable,
                    str(script),
                    "--manifest",
                    str(manifest),
                    "--wheel",
                    str(wheel),
                    "--bundle",
                    str(bundle),
                ],
                check=True,
                capture_output=True,
                text=True,
            )
            with zipfile.ZipFile(bundle) as archive:
                self.assertEqual(archive.read("data/000.xlsx"), b"frozen workbook")
                packed = json.loads(archive.read("manifest.json"))
                self.assertEqual(packed["workbooks"][0]["path"], "data/000.xlsx")

    @unittest.skipUnless(importlib.util.find_spec("pandas"), "pandas not installed")
    def test_frame_parity_detects_dtype_and_missing_value_differences(self):
        import pandas as pd
        import pickle

        with tempfile.TemporaryDirectory() as directory:
            left = Path(directory) / "left.pickle"
            right = Path(directory) / "right.pickle"
            left.write_bytes(pickle.dumps(pd.DataFrame({"a": [1, None]})))
            right.write_bytes(pickle.dumps(pd.DataFrame({"a": [1, ""]})))
            self.assertIn("different", benchmark.compare_frames(left, right))


if __name__ == "__main__":
    unittest.main()
