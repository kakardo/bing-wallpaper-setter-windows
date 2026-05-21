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
- Saves wallpapers organised by year and month under `Pictures\BingWallpaper\Wallpapers\`.
- Keeps a run log and stats file under `Pictures\BingWallpaper\Data\`.
- Optional lock screen wallpaper update (primary monitor only).
- Retries automatically if the network is unavailable at startup.
- Checks for new releases and shows a notice in the settings menu when one is available.
- Optional shuffle mode — rotates randomly through your saved wallpapers on a configurable interval (default 15 min). A manifest file keeps track of recent picks so the same image is not repeated too soon.
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
├── [I] Check interval  (30 min / 1 h / 2 h / 4 h / custom)
├── [O] Check hours     (all day / 06-23 / 07-22 / custom range)
├── [S] Shuffle
│       ├── [1] Toggle on/off
│       ├── [2] Change interval  (15 min / 30 min / 1 h / custom, min 5 min)
│       ├── [3] Change history   (5 / 10 / 25 / 50 / custom, 1-100)
│       ├── [4] Recalculate wallpaper list
│       └── [B]   Back
├── [W] Run now
├── [C] Recalculate stats
├── [V] View log        (last 10 entries)
├── [T] Switch to scheduled task  (shown when running from startup folder)
├── [U] Uninstall
└── [X] Exit
```

Market and resolution changes take effect at the next logon or hourly check, or immediately via **[W] Run now**.

## Shuffle mode

Shuffle rotates through your saved Bing wallpapers at a set interval rather than waiting for the next daily download. It picks randomly from your library, avoiding recently shown images so you don't see the same wallpaper twice in a row.

**To enable:** open `Settings.bat` and choose **[S] Shuffle**, then **[1] Toggle on**.

**Interval** — how often the wallpaper changes (default 15 min, minimum 5 min). Change via **[S] → [2]**.

**History size** — how many recent wallpapers are excluded from the next pick (default 10). If you have 50 saved wallpapers and history is set to 10, the next pick is drawn from the other 40. Change via **[S] → [3]**.

**Recalculate** — rescans your Wallpapers folder and rebuilds the index. Use this if you have added or removed images manually. Available via **[S] → [4]**.

Shuffle only runs while your PC is active. The library grows automatically as new Bing wallpapers are downloaded each day.

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
