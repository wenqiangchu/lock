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
- Configuration stores UNC paths only, never drive letters.
- Never create temporary files on the share (106-245 s penalty).
- `Save()` only, never `SaveAs`.

## Status

The current packaged add-in is still the original single-path version. The
multi-workbook configuration and workbook selector are planned but not yet
implemented. See `plans/PLAN.md` section 10 for current status and section 3
for the intended workflow.

The current build requires a manually maintained file at
`%LOCALAPPDATA%\MDDTLock\config.ini`. Create the `MDDTLock` folder if needed.
Its current single-workbook format is:

```ini
Version=1
MainFilePath=\\server\share\folder\workbook.xlsx
```

Replace the example with the actual UNC path. The planned version 2 multi-path
format is not supported by the current add-in build.

## Repository layout

```
src/       source (VBA modules)
test/      tests
plans/     planning documents, Chinese, authoritative
docs/      background material
prepare/   preparation-phase research archive
tools/     build / helper scripts
```
