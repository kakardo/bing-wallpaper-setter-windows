# Bing Wallpaper Setter for Windows

[![Download](https://img.shields.io/github/v/release/kakardo/bing-wallpaper-setter-windows?label=Download&style=for-the-badge)](https://github.com/kakardo/bing-wallpaper-setter-windows/releases/latest)

Automatically downloads the Bing wallpaper of the day and sets it as your Windows desktop background. Optionally updates the lock screen too.

## Installation

1. Download `BingWallpaperSetup_v[version].exe` from the [latest release](../../releases/latest).
2. Run it. Windows may show a SmartScreen warning (see below).
3. Follow the on-screen prompts. The installer will:
   - Ask whether to also update the lock screen wallpaper.
   - Register a scheduled task that runs at logon and checks hourly until the day's wallpaper is set.
   - Download and set today's wallpaper immediately.

## Windows SmartScreen warning

When you first run the EXE, Windows Defender SmartScreen may show:

> *"Windows protected your PC. Microsoft Defender SmartScreen prevented an unrecognized app from starting."*

This is expected. The EXE is unsigned (no paid code-signing certificate), so Windows flags it until the file builds a reputation. The source code is fully visible in this repository.

**To proceed:** click **More info**, then **Run anyway**.

## Features

- Downloads the daily Bing wallpaper at logon, then checks hourly until the day's image is available.
- Updates the wallpaper on all connected monitors.
- Saves wallpapers organised by year and month under `Pictures\BingWallpaper\`.
- Keeps a run log and stats file under `Pictures\BingWallpaper\Data\`.
- Optional lock screen wallpaper update (primary monitor only).
- Retries automatically if the network is unavailable at startup.
- `Settings.bat` for management and uninstall.

> **Multi-monitor:** Desktop wallpaper is set on all displays. Lock screen only updates on the primary monitor (Windows doesn't support per-monitor lock screens).

## Status and management

Open `Pictures\BingWallpaper\Settings.bat` to manage the wallpaper setter.

```
Settings
├── [L] Toggle lock screen
├── [M] Change market
│       ├── [1-8] Preset markets
│       ├── [C]   Custom (validated against Bing)
│       └── [B]   Back
├── [R] Change resolution
│       ├── [1] Auto-detect
│       ├── [2] 1920x1080  (Full HD)
│       ├── [3] 3840x2160  (4K)
│       ├── [4] 1366x768   (HD)
│       └── [B]   Back
├── [G] Log cap
│       ├── [1] Off
│       ├── [2] By size  (100 KB / 500 KB / 1 MB / custom)
│       ├── [3] By rows  (500 / 1 000 / 5 000 / custom)
│       └── [B]   Back
├── [W] Run now
├── [C] Recalculate stats
├── [U] Uninstall
└── [X] Exit
```

Market and resolution changes take effect at the next logon or hourly check, or immediately via **[W] Run now**.

## Parameters

`BingWallpaperSetup.ps1` accepts optional parameters if you prefer to run the script directly:

| Parameter | Default | Options |
|-----------|---------|---------|
| `-Market` | `en-US` | Any Bing market code, e.g. `en-GB`, `nb-NO` |
| `-Resolution` | Auto-detect | `1920x1080`, `1366x768`, `3840x2160` |
| `-SetLockScreen` | Off | Switch. Also updates the lock screen |

Example:
```powershell
.\BingWallpaperSetup.ps1 -Market en-GB -Resolution 3840x2160
```

## Uninstall

Open `Settings.bat` and choose **[U] Uninstall**. The scheduled task and scripts are removed. Your wallpaper photos stay. You will be asked whether to also delete the run log and stats.

## Requirements

- Windows 10 or 11
- PowerShell 5.1 or later (built-in on Windows 10/11)
- An internet connection. If unavailable at logon, the script retries automatically for up to an hour.
