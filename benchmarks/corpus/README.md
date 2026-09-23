# Frozen pandas XLSX corpus

[`pandas-real.json`](pandas-real.json) identifies 13 public workbooks from the
Australian Bureau of Statistics (ABS), the UK Office for National Statistics
(ONS), and the U.S. Bureau of Economic Analysis (BEA). It records each direct
download URL, publisher, reuse terms, SHA256, selected worksheet, and pandas
read options. The benchmark reads `header=None` to compare the entire chosen
worksheet without interpreting a publisher-specific heading as data columns.

Fetch and verify the input files with:

```bash
python tools/ci/fetch_pandas_corpus.py benchmarks/corpus/pandas-real.json
```

Downloads go to ignored `benchmarks/corpus/files/` and are excluded from the
source distribution. The fetcher rejects a changed remote file instead of
silently substituting a new revision. If a publisher replaces a file, keep the
old manifest and results together and create a new corpus revision.

The selected sheets include numeric tables, text labels, wide sheets (up to 95
columns), sparse layouts, and a larger ONS daily-deaths table. The ONS file
represents date as separate month/day/year columns, not typed Excel date cells.
The current corpus therefore does not exercise date-style decoding as a
real-world benchmark; the existing native and installed-wheel fixtures cover
typed dates separately. Most selected sheets are small. Interpret the median
as a measure of these published workbooks, not of all XLSX workloads.

ABS material is attributed under CC BY 4.0; ONS material is under the Open
Government Licence 3.0; BEA is a U.S. federal government publisher. The direct
publisher links and exact terms for each workbook remain in the manifest.
