# MDDT Checkout Tool

A VBA `.xlam` add-in that turns a single shared Excel workbook into a
"check out -> edit locally -> commit" workflow, so a small project group can
edit the same file without one person blocking everyone else.

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

- Single shared file; it must not be split per owner.
- Pure VBA add-in; no runtime dependencies, no admin rights.
- No M365 / SharePoint / OneDrive.
- Local working copy lives in `%LOCALAPPDATA%`, never in `Documents`/`Desktop`
  (those are cloud-synced).
- Configuration stores UNC paths only, never drive letters.
- Never create temporary files on the share (106-245 s penalty).
- `Save()` only, never `SaveAs`.

## Status

Phase 0 (project initialisation) complete. Next: Phase 1 (infrastructure).
See `plans/PLAN.md` section 9.

## Repository layout

```
src/       source (VBA modules)
test/      tests
plans/     planning documents, Chinese, authoritative
docs/      background material
prepare/   preparation-phase research archive
tools/     build / helper scripts
```
