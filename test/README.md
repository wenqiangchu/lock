# test/

All tests live here, mirrored by module.

```
test/MdlMerge/
test/MdlDiff/
test/MdlSchema/
test/MdlLock/
test/MdlInfra/
test/MdlUI/
```

Rules (see `AGENTS.md` section 5):

- Every module has tests before it is merged.
- Bug fixes require a regression test that fails before the fix.
- Pure-logic modules are tested with in-memory data, no Excel required.
- Tests that need Excel are grouped separately and document the prerequisite.
- `tools/Run-VbaTests.ps1` runs the `TestSuite.RunAllVbaTests` aggregate by default with a hard process timeout.
- The current runner launches hidden Excel as the VBA execution host. Pure-logic tests do not open workbooks or use the Excel object model; "no Excel required" describes the code-under-test dependency, not the VBA runtime process.
- `test/MdlExcel/` creates a temporary `.xlsx` below `%LOCALAPPDATA%\MDDTLock\tests` to verify MdlExcel open, edit, `Save()`, reopen, and on-disk lock behavior; cleanup removes the generated workbook and lock tree. `SaveAs` is used only to create this local fixture.
- `test/MdlUI/` contains in-memory tests for button-state mapping and the conflict-choice matrix; modal Office Forms behavior remains an Excel UI integration check.
