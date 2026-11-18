# src/

All product source lives here. One directory per module.

```
src/MdlInfra/        logging, error wrapping, path/UNC resolution, latency probe
src/MdlStorage/      workspace under %LOCALAPPDATA%, base snapshots, atomic rename
src/MdlSchema/       header parsing, AMG/product columns, anchor columns
src/MdlDiff/         cell-by-cell comparison (pure arrays, no Excel)
src/MdlMerge/        three-way merge (pure arrays, no Excel)
src/MdlLock/         column locks: MkDir mutex, owner.txt, heartbeat, fencing
src/MdlSession/      session state machine
src/MdlExcel/        the ONLY module allowed to touch the Excel object model
src/MdlUI/           ribbon callbacks, selection form, conflict form
```

Rules (see `AGENTS.md` section 4.4):

- Dependencies point downward only.
- `MdlDiff` and `MdlMerge` must be testable without Excel.
- Only `MdlExcel` may reference `Workbook`, `Worksheet`, `Range`.
