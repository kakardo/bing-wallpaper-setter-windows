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
- Updates the wallpaper on all connected monitors. Detects monitor layout changes between checks and reapplies automatically, so switching docking stations does not leave new screens blank.
- Saves wallpapers organised by year and month under `Pictures\BingWallpaper\Wallpapers\`. You can also add your own images here and shuffle will pick them up automatically.
- Keeps a run log and stats file under `Pictures\BingWallpaper\Data\`.
- Optional lock screen wallpaper update (primary monitor only). Configurable display timeout for when the PC is plugged in; battery follows Windows default.
- Retries automatically if the network is unavailable at startup.
- Checks for new releases and shows a notice in the settings menu when one is available.
- Optional shuffle mode that rotates randomly through your saved wallpapers on a configurable interval (default 15 min). A manifest file keeps track of recent picks so the same image is not repeated too soon.
- History catch-up that automatically downloads the last 7 days of Bing wallpapers whenever a new image is added, so days missed while your PC was off are filled in. Configurable or can be turned off entirely via Settings.
- `Settings.bat` for management and uninstall.

> **Multi-monitor:** Desktop wallpaper is set on all displays. Lock screen only updates on the primary monitor (Windows does not support per-monitor lock screens).

## Status and management

Open `Pictures\BingWallpaper\Settings.bat` to manage the wallpaper setter.

```
Settings
├── [L] Lock screen
│       ├── [1] Toggle on/off
│       ├── [2] Set display timeout
│       │       ├── [1-4] Preset durations (5 / 10 / 15 / 30 min)
│       │       ├── [C]   Custom (1-120 min)
│       │       ├── [D]   Windows default (do not manage)
│       │       └── [B]   Back
│       └── [B]   Back
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
│       ├── [5] Auto-recalculate interval  (1 day / 7 days / 30 days / custom / off)
│       └── [B]   Back
├── [H] History catch-up
│       ├── [1] Toggle on/off
│       ├── [2] Change days  (1 / 3 / 7 / custom, max 7)
│       ├── [3] Run now
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

**Interval**: how often the wallpaper changes (default 15 min, minimum 5 min). Change via **[S] → [2]**.

**History size**: how many recent wallpapers are excluded from the next pick (default 10). If you have 50 saved wallpapers and history is set to 10, the next pick is drawn from the other 40. Change via **[S] → [3]**.

**Recalculate**: rescans your Wallpapers folder and rebuilds the index. Use this if you have added or removed images manually. Available via **[S] → [4]**.

**Auto-recalculate**: runs a recalculate automatically on a schedule (default every 7 days) so the index stays in sync without manual intervention. Configure via **[S] → [5]**, or set to Off to disable.

Shuffle only runs while your PC is active. The library grows automatically as new Bing wallpapers are downloaded each day. You can also drop your own images (`.jpg`, `.jpeg`, `.png`, `.bmp`) into the `Wallpapers` folder and they will be included automatically.

## History catch-up

History catch-up automatically downloads recent Bing wallpapers that your PC missed, useful if you were away for a few days or just installed the program and want to fill in your library straight away.

When today's wallpaper is downloaded, the program checks the previous days (up to the configured limit) and silently downloads any that are not already on disk. The images go straight into the shuffle library and do not change what is currently on your desktop.

**To configure:** open `Settings.bat` and choose **[H] History catch-up**.

**Toggle on/off**: enabled by default. Turn it off if you only want the current day's image. Available via **[H] → [1]**.

**Days**: how many previous days to check on each run (default 7, maximum 7). Bing only keeps the last 7 days available via its API. Change via **[H] → [2]**. Changing the day count applies the new setting immediately.

**Run now**: triggers a one-off catch-up without waiting for the next scheduled check. Available via **[H] → [3]**.

Changing the day count or turning catch-up on both trigger an immediate download so your library is up to date straight away.

## Lock screen

When lock screen updates are enabled, the program sets the lock screen image to the current wallpaper each time a new one is downloaded.

**Display timeout:** you can configure how long Windows keeps the lock screen visible before turning off the display when your PC is plugged in. Battery follows Windows default and is never changed. The installer asks for this value when you enable lock screen, and you can update it any time via **[L] → [2] Set display timeout**.

The default timeout is 10 minutes on AC. To let Windows manage it entirely, choose **[D] Windows default** in the timeout menu.

> **Multi-monitor note:** lock screen only updates on the primary monitor. Windows does not support per-monitor lock screens.

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
- PowerShell 5