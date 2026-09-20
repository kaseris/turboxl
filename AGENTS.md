# TurboXL Agent Guide

Use [CONTRIBUTING.md](CONTRIBUTING.md) as the source of truth for feature and
release procedures. Keep it synchronized when CI or packaging behavior changes.

## Feature changes

- Work from current `main` on a focused branch.
- Add native coverage under `tests/` for behavior changes.
- Keep `include/xlsxcsv.hpp`, `src/python/module.cpp`, and `README.md` aligned
  when public behavior changes.
- Run the relevant native tests. For packaging, CI, or release changes, also run
  `python -m unittest discover -s tools/ci -p 'test_*.py' -v` and validate an
  sdist.
- Do not bump the version in an ordinary feature change.
- Do not merge until the required **CI passed** check succeeds.

## Skipping CI for documentation-only changes

For a documentation-only change that cannot affect builds, tests, packaging, or
runtime behavior, a direct push to `main` may intentionally skip GitHub Actions.
End the commit subject with `[skip ci]`, for example:

```bash
git commit -m "docs: clarify contribution workflow [skip ci]"
git push origin main
```

Use `[skip ci]` only for Markdown or other non-executable documentation changes.
Do not use it for source code, tests, build configuration, dependencies,
packaging, workflows, version bumps, or release tags. Before pushing, stage
explicit file paths and inspect the staged diff so unrelated files are not
included.

## Current release method

- `.github/workflows/release.yml` is triggered only by a pushed tag matching
  `v*`; workflows do not create release tags.
- `CMakeLists.txt` is the version source of truth. `pyproject.toml` extracts the
  CMake `project()` version; this project does not use `setuptools-scm`.
- Release tags must use exact SemVer form `vX.Y.Z`, match the CMake version,
  point at the checked-out commit, and be contained in `origin/main`.
- The release workflow validates the tag, runs native tests, builds and validates
  one sdist plus 12 wheels, publishes them to PyPI, then publishes a GitHub
  Release with the same files and `SHA256SUMS`.
- The PyPI publish step uses `PYPI_API_TOKEN` and supports safe retries of partial
  uploads by comparing SHA256 hashes.

## Release procedure

1. Update the version in `project(turboxl VERSION X.Y.Z LANGUAGES CXX)` in
   `CMakeLists.txt`, review and merge it, and wait for **CI passed** on `main`.
2. Optionally rehearse with **Actions → CI → Run workflow**; it builds the full
   distribution set without publishing.
3. Fast-forward local `main` and confirm `HEAD` equals `origin/main`.
4. Confirm `vX.Y.Z` does not exist locally or remotely.
5. Create an annotated tag and validate it before pushing:

   ```bash
   git tag -a vX.Y.Z -m "Release vX.Y.Z"
   python tools/ci/validate_release.py vX.Y.Z
   git show --no-patch --decorate vX.Y.Z
   git push origin refs/tags/vX.Y.Z
   ```

6. Verify the **Release** workflow, PyPI files, and GitHub Release assets.

An incorrect local tag may be deleted only before it is pushed. Never move or
reuse a pushed release tag and never overwrite a PyPI version; correct the issue
and use a new version. For a partial publication, re-run failed jobs while the
validated artifacts are retained (30 days). Existing PyPI files are accepted
only when their hashes match.
