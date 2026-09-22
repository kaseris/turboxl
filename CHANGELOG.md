# Changelog

## Unreleased

- Hardened the typed Python `Workbook` API: `max_cells` now applies to every
  public sheet read, and failed lazy shared-string or style initialization is
  retry-safe.
- Documented typed workbook ownership, source handling, scalar behavior,
  resource limits, and XLSX scope boundaries.
