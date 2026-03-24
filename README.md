# Winget Update Manager

Winget Update Manager is a Windows desktop app for scanning, updating, installing, and managing packages from `winget` and global `npm`.

![Winget Update Manager screenshot](docs/Screenshot%202026-03-24%20041720.png)

## What It Does

- Scans for available `winget` and global `npm` updates
- Lets you update selected packages, all packages, or saved update groups
- Shows installed apps and exports a reinstall script
- Includes a Discover page for searching the `winget` repository and installing packages
- Tracks update history, rollback candidates, health score, and dashboard activity
- Supports optional Windows startup, system tray usage, admin elevation, and native toast notifications

## Requirements

- Windows 10 or Windows 11
- `winget` installed and available in `PATH`
- Optional: `npm` in `PATH` if you want global `npm` package support
- Optional: `winotify` if you want native Windows toast notifications when running from Python

## Run It

### Option 1: Use the packaged app

Launch:

`dist\WingetUpdateManager.exe`

### Option 2: Run from source

From this folder:

```powershell
python winget_update_manager.py
```

## Basic Workflow

1. Open the app.
2. Go to `Updates`.
3. Click `Check Updates`.
4. Choose one of:
   - `Update Selected`
   - `Update Group`
   - `Update All`
5. Watch the progress bar and live console output.

## Pages

### Dashboard

- Shows installed count, pending updates, last scan, total successful updates
- Shows cache freshness, health score, update activity bar chart, and success/fail/skip chart

### Updates

- Scan for pending updates
- Filter by search text or saved update group
- Update selected packages, a whole group, or everything
- Right-click a package for quick actions like exclude, quiet mode, changelog, or copy package ID

### Installed Apps

- Load installed `winget` and global `npm` packages
- Browse package details
- Export the installed list as JSON
- Export a reinstall script as `.ps1` or `.bat`
- Right-click eligible `winget` packages to roll back to the previous recorded version

### Discover

- Search the `winget` repository
- Install packages directly from search results
- Open package details before installing

### History

- Review successful, failed, and skipped actions
- Export history as CSV or JSON
- Use rollback and changelog actions for eligible `winget` packages

### Settings

- Switch theme
- Enable auto-check on launch
- Set cache freshness window
- Enable silent updates
- Configure parallel update workers
- Create and edit update groups
- Configure startup launch and admin elevation
- Manage quiet-mode packages and exclusions

## Update Groups

Use `Settings > Update Groups` to create named groups such as:

- `Browsers`
- `Dev Tools`
- `Daily Drivers`

After saving a group, go back to `Updates`, pick it from `Group Filter`, and click `Update Group`.

## Rollback

Rollback is available for `winget` packages only.

The app records successful installed versions in history. If a package has a previous recorded version, you can right-click it in `Installed Apps` or `History` and choose `Rollback to Previous Version`.

## Changelog Viewer

Changelog viewing is available for `winget` packages when the package metadata exposes release notes.

- GitHub release notes open inside the app
- Other release-note URLs open in the browser

## Notifications

- In-app notifications appear in the bell menu in the page header
- Native Windows toasts are used when supported
- Otherwise the app falls back to tray or in-app notifications

## Startup Modes

The app supports:

- Normal launch
- `--scan-only` for scheduled background scans
- `--minimized` for startup-at-login launches

Examples:

```powershell
python winget_update_manager.py --scan-only
python winget_update_manager.py --minimized
```

## Build The EXE

If you want to rebuild the packaged app:

```powershell
pyinstaller --noconfirm WingetUpdateManager.spec
```

The output will be written to:

`dist\WingetUpdateManager.exe`
