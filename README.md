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
| `docs/current-user-flow.md` | Source-based walkthrough of current user and VBA behavior |
| `prepare/` | Completed research from the preparation phase (read-only) |

`plans/` is the single source of truth for what to build and in what order.

## Non-negotiable constraints

- Each workbook has one shared master; it must not be split into per-owner business copies.
- Pure VBA add-in; no runtime dependencies, no admin rights.
- No M365 / SharePoint / OneDrive.
- Local working copy lives in `%LOCALAPPDATA%`, never in `Documents`/`Desktop`
  (those are cloud-synced).
- Configuration accepts UNC paths and absolute drive-letter paths that point to
  the shared directory. The add-in does not inspect whether a drive is mapped;
  local drive paths are accepted by the parser for controlled tests only, and
  production entries must point to the shared workbook.
- Never create temporary files on the share (106-245 s penalty).
- Shared business workbooks are saved at runtime with `Save()` only. Add-in
  packaging and local test-fixture creation may use `SaveAs` to create artifacts.

## Status

The checked-in source supports a manually maintained multi-workbook list and a
workbook selector. Packages built from earlier source revisions may not include
the current path and lock behavior; rebuild from this source before use. See
`plans/PLAN.md` section 10 for verification status and section 3 for the
intended workflow.

The add-in reads a manually maintained file at
`%LOCALAPPDATA%\MDDTLock\config.ini`. Each Windows user maintains this file on
their own computer. Create the `MDDTLock` folder if needed.
The legacy single-workbook format remains readable:

```ini
Version=1
MainFilePath=\\server\share\folder\workbook.xlsx
```

Replace the example with the path used to reach the shared workbook. Maintain
the multi-workbook file as follows:

1. In File Explorer, enter `%LOCALAPPDATA%\MDDTLock` in the address bar. Create
   the `MDDTLock` folder if it does not exist.
2. Create or edit `config.ini` with a plain-text editor. Do not save it as
   `config.ini.txt`.
3. Add one `WorkbookPath=` line per shared workbook. Use either its UNC path
   or an absolute mapped-drive path such as `X:\team-a\mddt_alpha.xlsx`. Each
   user may use a different drive letter if it maps to the same shared folder:
   locks are stored beside the workbook, so clients reach the same lock files.
   The add-in does not verify the mapping target. Do not configure a local-only
   workbook for production; local absolute paths are accepted for controlled
   Excel tests.
4. Save the file. The next click on “Start Editing” reads the updated
   configuration; Excel does not need to be restarted.

Example for two workbooks in different folders:

```ini
Version=2
WorkbookPath=\\server\share\team-a\mddt_alpha.xlsx
WorkbookPath=X:\team-b\mddt_beta.xlsx
```

To add a workbook, add another `WorkbookPath=` line. To remove one, delete its
line. Keep at least one valid path and list each shared workbook once. The
add-in only detects duplicate path strings; it cannot tell whether a UNC path
and a drive-letter path are aliases for the same file. The add-in does not
create or repair shared workbooks. If one listed file is missing or
unavailable, choose another entry or correct that path in the configuration
file; a missing configuration file is reported with its exact local path.
After a workbook is selected, the add-in distinguishes a missing file, denied
access, and an unavailable UNC or drive-letter path when Windows reports those
conditions.
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
