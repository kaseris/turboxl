# Changelog

## Unreleased

- Reduced per-read overhead for typed and pandas reads, most visibly on small
  workbooks (13 real workbooks, each 10-48% faster in same-session A/B runs on
  macOS arm64, median 26%): number-format classification no longer constructs
  `std::regex` objects, `styles.xml` (CsvOnly mode), `sharedStrings.xml`,
  `workbook.xml` and the OPC relationship parts use a lightweight scanner that
  falls back to libxml2 for anything outside its subset, and shared strings no
  longer reserve an 8 MB arena per open.
- Added `Sheet.to_python(trim_trailing_empty=True)`, used by the pandas adapter
  to trim trailing empty rows and columns natively instead of in Python.
- Hardened the typed Python `Workbook` API: `max_cells` now applies to every
  public sheet read, and failed lazy shared-string or style initialization is
  retry-safe.
- Documented typed workbook ownership, source handling, scalar behavior,
  resource limits, and XLSX scope boundaries.
