#!/usr/bin/env python3
"""Fetch a frozen public XLSX corpus, rejecting changed remote files."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from urllib.request import Request, urlopen
import zipfile


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", type=Path)
    args = parser.parse_args()
    manifest = args.manifest.resolve()
    entries = json.loads(manifest.read_text())["workbooks"]
    for entry in entries:
        target = (manifest.parent / entry["path"]).resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        if (
            target.is_file()
            and hashlib.sha256(target.read_bytes()).hexdigest() == entry["sha256"]
        ):
            print(f"verified {entry['id']}")
            continue
        url = entry["download_url"]
        if not url.startswith("https://"):
            raise ValueError(f"non-HTTPS download URL for {entry['id']}")
        request = Request(
            url, headers={"User-Agent": "TurboXL corpus verification/1.0"}
        )
        with urlopen(request, timeout=60) as response:
            contents = response.read()
        actual = hashlib.sha256(contents).hexdigest()
        if actual != entry["sha256"]:
            raise ValueError(f"remote file changed for {entry['id']}: {actual}")
        temporary = target.with_suffix(".tmp")
        temporary.write_bytes(contents)
        if not zipfile.is_zipfile(temporary):
            temporary.unlink()
            raise ValueError(f"not an XLSX ZIP archive: {entry['id']}")
        temporary.replace(target)
        print(f"downloaded {entry['id']}")


if __name__ == "__main__":
    main()
