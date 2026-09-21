# Contributing to TurboXL

This guide describes how a change moves from a development branch to a shipped
TurboXL release. For build prerequisites and platform-specific dependency
installation, see [Building](README.md#building).

## Ship a feature

### 1. Start from the current `main`

Create a focused branch from an up-to-date checkout. Use a descriptive prefix
such as `feature/`, `fix/`, `perf/`, or `docs/`.

```bash
git switch main
git pull --ff-only
git switch -c feature/<short-description>
```

Keep each pull request limited to one coherent change. Do not change the project
version as part of an ordinary feature pull request; version bumps are release
changes.

### 2. Implement the complete change

A feature is ready for review when its implementation, tests, and user-facing
documentation agree.

- Add or update native tests under `tests/` for C++ behavior.
- Keep the public C++ API in `include/xlsxcsv.hpp` and the Python bindings in
  `src/python/module.cpp` consistent when a feature affects both interfaces.
- Keep `src/python/turboxl/__init__.py` re-exports synchronized with the native
  bindings so installed-package imports remain backward compatible.
- Update `README.md` for public API, behavior, compatibility, or build changes.
- Preserve the read-only conversion scope documented in the README.
- Add representative XLSX fixtures through the existing fixture helpers rather
  than committing incidental generated build output.
- Treat performance-sensitive parsing and serialization changes as performance
  work: check output parity and benchmark the affected path.

Generated directories such as `build/`, `out/`, `dist/`, and virtual
environments must not be committed. The source-distribution configuration also
excludes generated build and output directories so local artifacts cannot leak
into a release archive.

### 3. Verify locally

Install the platform dependencies from the README, then run the checks relevant
to the change. The following matches the native configuration used by CI:

```bash
cmake -S . -B out/native \
  -DCMAKE_BUILD_TYPE=Debug \
  -DBUILD_PYTHON=OFF \
  -DBUILD_TESTS=ON \
  -DTURBOXL_ENABLE_IPO=OFF \
  -DTURBOXL_PORTABLE_BUILD=ON
cmake --build out/native --config Debug --parallel 4
ctest --test-dir out/native -C Debug --output-on-failure --no-tests=error
```

When changing packaging, CI, or release code, also run the CI helper tests and
build an sdist:

```bash
python -m pip install -c tools/ci/constraints.txt \
  build scikit-build-core nanobind packaging twine
python -m unittest discover -s tools/ci -p 'test_*.py' -v
python -m build --sdist --no-isolation
python -m twine check --strict dist/*
```

Run the narrower test or benchmark for the code you touched as well. The full
wheel matrix is intentionally left to GitHub Actions because it covers four
platform targets and all supported Python variants.

For typed worksheet extraction work, compare end-to-end Python materialization
against the pinned python-calamine version and retain the machine-readable
result when the change is intended to make or update a performance claim:

```bash
python tools/benchmark_typed.py --rows 60000 --warmups 2 --rounds 9 \
  --json-output typed-benchmark-results.json
```

For a performance-sensitive change, first capture the same command on the
unmodified base commit, then run the candidate with
`--baseline-json <base-result.json>`. The comparison fails if any typed fixture
loses the existing 10% advantage over calamine or regresses more than 5% from
the same-machine baseline.

### 4. Open and merge a pull request

Explain the user-visible behavior, important implementation choices, and the
checks you ran. Call out API or compatibility changes explicitly.

Every pull request to `main` runs `.github/workflows/verify.yml`. Do not merge
until the required **CI passed** check succeeds. That check covers:

- CI helper tests and workflow linting;
- native Debug builds and tests on Linux, Windows, macOS Intel, and macOS Apple
  Silicon;
- one source distribution and 12 wheels across the four supported platform
  targets;
- package, wheel inventory, and stable-ABI validation; and
- Windows output-parity and performance checks against pinned
  `python-calamine`.

Merge the reviewed change into `main`. A merge or push to `main` verifies the
commit but does not publish a release.

## Ship a new version

Releases are deliberate, tag-triggered operations. The version has one source
of truth: the `project(turboxl VERSION X.Y.Z LANGUAGES CXX)` declaration in
`CMakeLists.txt`. `pyproject.toml` reads that value for Python package metadata.

Only a pushed tag matching `v*` starts `.github/workflows/release.yml`; no
workflow creates a tag automatically.

### 1. Prepare the release commit

Choose the next semantic version and update only the version in the CMake
`project()` declaration. Review and merge that change into `main`, then wait for
**CI passed**.

Before tagging, use **Actions → CI → Run workflow** on the intended release
branch if you want a fresh full-matrix rehearsal. This builds and tests the
complete release set without publishing it.

### 2. Tag the exact release commit

Update your local `main`, verify its relationship to `origin/main`, and confirm
that the version tag does not already exist. Replace `X.Y.Z` below with the
version already present in `CMakeLists.txt`.

```bash
git switch main
git pull --ff-only
git status --short
git rev-parse HEAD
git rev-parse origin/main
git tag --list vX.Y.Z
git ls-remote --exit-code --tags origin refs/tags/vX.Y.Z
```

`git ls-remote` exits with status 2 when the remote tag is absent, which is the
expected result for a new release. Create an annotated tag, run the same release
input validation used by GitHub Actions, and push that tag explicitly:

```bash
git tag -a vX.Y.Z -m "Release vX.Y.Z"
python tools/ci/validate_release.py vX.Y.Z
git show --no-patch --decorate vX.Y.Z
git push origin refs/tags/vX.Y.Z
```

The validator requires an exact `vX.Y.Z` tag, a matching CMake version, a tag
pointing at the checked-out commit, and a commit contained in `origin/main`.

If an unpushed local tag is wrong, delete it with `git tag -d vX.Y.Z` and create
it again on the correct commit. Once a tag has been pushed, do not move or reuse
it. Correct the problem and ship a new version instead.

### 3. Verify publication

Watch the **Release** workflow triggered by the tag. It:

1. validates the tag and release commit;
2. runs the native test matrix;
3. builds and validates the sdist and 12 wheels;
4. publishes the validated files to PyPI; and
5. creates a GitHub Release with generated notes, the same distributions, and
   `SHA256SUMS`.

Confirm that the workflow succeeded, the new version and expected files appear
on PyPI, and the GitHub Release is public with its assets attached.

### Failed or partial releases

- Use **Re-run failed jobs** while the workflow artifacts are available; release
  artifacts are retained for 30 days.
- A retry checks already-published PyPI files against the retained artifacts and
  uploads only missing files. Any hash mismatch fails closed.
- Re-running the GitHub Release job is safe after PyPI publication succeeds.
- Never overwrite a PyPI version or move a published tag. If artifacts have
  expired or the released input itself is wrong, fix the issue and release a new
  version.
