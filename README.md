# MDDT Checkout Tool

A VBA `.xlam` add-in for checking out, editing locally, and committing changes
to shared Excel workbooks. The planned workflow supports multiple configured
workbooks while each edit session operates on one selected workbook.

## Read this first

| Document | Purpose |
|---|---|
| `AGENTS.md` | **Binding rules** for anyone (human or agent) working here |
| `plans/PLAN.md` | Master plan: phases, milestones, status |
| `plans/phase-NN-*.md` | Per-phase execution plans |
| `docs/` | Background reference (not authoritative) |
| `prepare/` | Completed research from the preparation phase (read-only) |

`plans/` is the single source of truth for what to build and in what order.

## Non-negotiable constraints

- Each workbook has one shared master; it must not be split into per-owner business copies.
- Pure VBA add-in; no runtime dependencies, no admin rights.
- No M365 / SharePoint / OneDrive.
- Local working copy lives in `%LOCALAPPDATA%`, never in `Documents`/`Desktop`
  (those are cloud-synced).
- Configuration stores UNC paths only. Mapped drive paths are rejected.
- Never create temporary files on the share (106-245 s penalty).
- `Save()` only, never `SaveAs`.

## Status

The checked-in source supports a manually maintained multi-workbook list and a
workbook selector. Previously built `.xlam` packages remain single-path until
rebuilt from this source. See `plans/PLAN.md` section 10 for verification status
and section 3 for the intended workflow.

The add-in reads a manually maintained file at
`%LOCALAPPDATA%\MDDTLock\config.ini`. Each Windows user maintains this file on
their own computer. Create the `MDDTLock` folder if needed.
The legacy single-workbook format remains readable:

```ini
Version=1
MainFilePath=\\server\share\folder\workbook.xlsx
```

Replace the example with the actual UNC path. For the multi-workbook source,
maintain the file as follows:

1. In File Explorer, enter `%LOCALAPPDATA%\MDDTLock` in the address bar. Create
   the `MDDTLock` folder if it does not exist.
2. Create or edit `config.ini` with a plain-text editor. Do not save it as
   `config.ini.txt`.
3. Add one `WorkbookPath=` line per shared workbook, using its UNC path. Use the
   same UNC spelling on every team member's computer. Do not use a mapped drive
   letter such as `X:`.
4. Save the file. The next click on “Start Editing” reads the updated
   configuration; Excel does not need to be restarted.

Example for two workbooks in different folders:

```ini
Version=2
WorkbookPath=\\server\share\team-a\mddt_alpha.xlsx
WorkbookPath=\\server\share\team-b\mddt_beta.xlsx
```

To add a workbook, add another `WorkbookPath=` line. To remove one, delete its
line. Keep at least one valid path and do not duplicate a path. The add-in does
not create or repair shared workbooks. If one listed file is missing or
unavailable, choose another entry or correct the UNC path in the configuration
file; a missing configuration file is reported with its exact local path.
After a workbook is selected, the add-in distinguishes a missing file, denied
access, and an unavailable network path when Windows reports those conditions.
If Excel opens it read-only, the message explains that another Excel session,
the read-only attribute, or account permissions may prevent writing; Excel does
not always expose which one caused that state.

## Repository layout

```
src/       source (VBA modules)
test/      tests
plans/     planning documents, Chinese, authoritative
docs/      background material
prepare/   preparation-phase research archive
tools/     build / helper scripts
```
