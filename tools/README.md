# tools/

Build and helper scripts. Not product code and not shipped.

Typical contents:
- packaging a `.xlam` from `src/`
- importing `.bas` modules into a workbook for tests
- running test suites

`Build-Addin.ps1` creates a local package under `build/` by default, imports the
VBA modules with Excel, injects the Ribbon XML, and verifies the resulting
Office package. It uses a hard timeout for the isolated Excel process and
refuses to overwrite an existing output file.
