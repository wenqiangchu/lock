# test/

All tests live here, mirrored by module.

```
test/MdlMerge/
test/MdlDiff/
test/MdlSchema/
test/MdlLock/
test/MdlInfra/
```

Rules (see `AGENTS.md` section 5):

- Every module has tests before it is merged.
- Bug fixes require a regression test that fails before the fix.
- Pure-logic modules are tested with in-memory data, no Excel required.
- Tests that need Excel are grouped separately and document the prerequisite.
