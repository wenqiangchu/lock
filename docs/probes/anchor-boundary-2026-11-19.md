# A–E Anchor and Row-Boundary Probe

Date: 2026-11-19

## Purpose

Check that product headers are selected only from A–E, internal blank rows keep
their physical positions, formatting-only worksheet tails do not extend the
data boundary, and a declared-column value with no product anchor is rejected.

## Reproduction

```powershell
python tools/probe_xlsx_structure.py prepare/测试结果.xlsx prepare/_smoke/schema_probe.xlsx
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\Probe-ExcelBoundary.ps1 -TimeoutSeconds 60
python .\test\tools\test_probe_xlsx_structure.py
```

The OpenXML probe emits only sheet indexes, column indexes, counts, and row
boundaries. It does not emit workbook names, headers, or cell values. The Excel
probe creates a new unsaved synthetic workbook in a dedicated Excel process and
has a 60-second worker timeout followed by a 20-second Excel-exit grace period;
the watchdog terminates Excel and reports failure if it remains alive.

## Archived Workbook Shape

The repository contains two archived XLSX files; neither is confirmed as a
representative business template.

- `prepare/测试结果.xlsx`: one worksheet, no AMG header, five product headers
  in A–E and three more after E. It cannot verify the MDDT table contract.
- `prepare/_smoke/schema_probe.xlsx`: seven worksheets, six with AMG headers;
  one AMG worksheet has no A–E product anchor. Other AMG worksheets have
  product anchors in A–D or A–E, including one with product columns after E.
  The probe found no internal fully empty rows in this file. The sample is a
  schema probe, not a confirmed business template.

The anonymized per-sheet counts and row boundaries can be reproduced with the
command above. The archived sample is not used to relax the A–E contract.

## Controlled Excel Result

Excel 16.0 build 19127 produced this report from synthetic data:

- Product anchors were detected in columns A, C, and E. The blank B header was
  ignored; a product header in F was excluded from the anchor set.
- Data existed on rows 2, 4, and 10. The effective end row was 10, with six
  fully blank physical rows retained between the first and last test rows.
- Formatting applied at row 1,048,576 did not extend the candidate data end.
- An orphan fixture had anchored rows through 10 and a declared AMG value in
  row 30. The value at row 30 was identified as anchorless and rejected.
- The synthetic workbook was not saved. The Excel process exited before the
  probe returned.

The workbook asserts the candidate boundary behavior used by `MdlExcel`; it
does not prove that all business templates follow the assumed row distribution,
that all columns have the same population, or that Samba behavior is equivalent.

## Implementation Consequence

The current implementation takes the maximum `End(xlUp)` row among A–E anchors
and declared AMG columns, keeps rows 2 through that physical row without
compaction, and fails closed when a declared value is outside every anchor row.
The production-template prerequisite remains unverified until a representative
business workbook is inspected read-only.
