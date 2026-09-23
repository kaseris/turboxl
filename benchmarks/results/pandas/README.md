# Issue #109: pandas engine evaluation

**Decision: no-go for an upstream first-party engine proposal at this time.**
The local adapter works after explicit registration, but the upstream XLSX
reader suite still has failures and the real-workbook median does not meet the
20% speed threshold. The adapter merged after the latest TurboXL release, so
external adoption evidence is not yet available.

## Evidence

The [frozen corpus](../../corpus/README.md) contains 13 public XLSX workbooks
from three publishers. The two macOS arm64 runs used Python 3.12, pandas 3.0.0,
python-calamine 0.8.2, an installed candidate TurboXL wheel, two warmups, and
nine timed rounds per engine and workbook. Engine order alternated. Timing
covers `pd.read_excel` through completed DataFrame construction; imports and
registration were outside the clock. Individual samples, peak process RSS,
wheel SHA256, workbook SHA256, exact DataFrame comparisons, and machine data
are in [run 1](macos-arm64-20260923-run1.json) and
[run 2](macos-arm64-20260923-run2.json).

| Measure | Run 1 | Run 2 |
| --- | ---: | ---: |
| DataFrame parity | 13/13 | 13/13 |
| Median per-workbook advantage over calamine | -160.7% | -140.8% |
| 20% gate | Fail | Fail |

The larger ONS daily-deaths sheet favors TurboXL; most small published sheets
favor calamine. The corpus is skewed toward small, preformatted statistical
tables and contains no typed Excel date cells. These limits matter when
generalizing the result. A negative advantage means TurboXL took longer;
the two medians correspond to roughly 2.61× and 2.41× calamine's time on
the selected sheets. The daily-deaths sheet favored TurboXL by 29.3% and
29.1% in the two runs.

As a separate diagnostic, the three 60,000-row typed fixtures all had exact
DataFrame parity and TurboXL advantages of 34.8%, 31.0%, and 36.3%. Their
[raw report](macos-arm64-20260923-synthetic.json) is excluded from the real
corpus gate. Recreate its inputs with
`python tools/ci/generate_pandas_synthetic.py --rows 60000 --output-dir /path/to/fixtures`.

The [upstream compatibility report](compatibility-macos-arm64-20260923.json)
uses pinned pandas source tags, copies their reader test and data into isolated
installed-package environments, and adds TurboXL only to the `.xlsx` test
parameter. HTTP and S3 cases are excluded because their external pytest
fixtures are not installed; the report records exclusions and raw failure
output. The remaining selected cases were:

| pandas | Passed | Failed | Skipped |
| --- | ---: | ---: | ---: |
| 2.2.3 | 97 | 13 | 6 |
| 2.3.3 | 98 | 13 | 5 |
| 3.0.0 | 102 | 12 | 5 |

Remaining failures cluster around blank versus error cells, headers and
multi-index layout, row-limit edge cases, and corrupt-archive exception types.
The adapter now trims trailing styled empty cells; that fix made all 13 corpus
DataFrames exactly match calamine under the recorded read options. TurboXL's
typed API currently returns `None` for both blank and Excel error cells, so
matching every pandas option will require a careful public behavior decision
or distinct typed markers upstream of the adapter.

GitHub code searches for `turboxl.pandas` and `engine="turboxl"` found only this
repository on 23 September 2026; they do not establish external adoption.
The latest release, v0.3.0, predates the adapter merge. No pandas maintainer
feedback specific to TurboXL was found. The third-party engine registration
[proposal #61584](https://github.com/pandas-dev/pandas/issues/61584) remains
open without comments, so the local adapter still uses pandas' private
`ExcelFile._engines` map. That requires maintenance against supported pandas
versions. TurboXL supports XLSX only, whereas calamine covers additional Excel
and OpenDocument formats; the wheel matrix currently covers Linux, Windows,
macOS arm64, and macOS Intel.

## Azure preparation

`tools/azure_benchmark_once.sh --mode pandas` packages a frozen manifest and
exact Linux candidate wheel, then uses the existing matched Intel D4s v6 and
AMD D4as v6 VMs in UK South. The `--prepare-only` path validated bundle
creation and all 13 input hashes locally. No Azure resources were provisioned.
The active subscription has 10 regional vCPUs available; both four-core SKUs
were listed without restrictions. The [Azure Retail Prices API](https://learn.microsoft.com/en-us/rest/api/cost-management/retail-prices/azure-retail-prices) quoted USD
$0.233/hour and $0.211/hour for the Linux VMs on 23 September 2026, or
$0.444/hour together, plus storage and network charges. A Linux wheel from
candidate CI is needed before an Azure pandas run. Given the local no-go
result, an Azure run would be exploratory rather than a gate-clearing step.

## Reproduction

```bash
python tools/ci/fetch_pandas_corpus.py benchmarks/corpus/pandas-real.json
python tools/benchmark_pandas.py \
  --manifest benchmarks/corpus/pandas-real.json \
  --wheel /path/to/candidate.whl --warmups 2 --rounds 9 \
  --json-output /path/to/run.json
python tools/ci/run_pandas_upstream.py \
  --wheel /path/to/candidate.whl --work-dir /path/to/temporary-work \
  --json-output /path/to/compatibility.json
```

Revisit the go decision after a release with verifiable adoption, upstream
compatibility fixes, and a broader date-bearing corpus that meets the repeated
20% end-to-end median advantage gate. No upstream proposal was opened.
