# Testing Gaps

These items were not fully tested end to end during the last implementation pass.

## Not Fully Tested

- Live `winget` update/install flows against real packages, including:
  - retry and exponential backoff during real failures
  - parallel worker execution with `1`, `2`, and `4` workers
  - rollback against a package with real recorded version history
- Native Windows toast notifications through `winotify` on a machine with `winotify` installed and active
- Real changelog resolution from package metadata, especially:
  - GitHub-backed release note resolution from `Release Notes URL`
  - fallback behavior for non-GitHub release-note URLs

## Only Smoke Tested

- App compile and import checks
- Tk startup/launch smoke test
- Config and history migration behavior in temp appdata directories
- PyInstaller packaging and EXE build output

## Follow-Up Manual Checks

- Run update flows on a few real `winget` packages
- Force a failure to confirm retry logging and backoff timing
- Test rollback on a package with at least two recorded successful versions
- Confirm native toast behavior on a system with `winotify`
- Open changelog actions from real packages with GitHub and non-GitHub release-note URLs
