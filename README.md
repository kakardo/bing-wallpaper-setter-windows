# Bing Wallpaper Setter for Windows

Automatically downloads the Bing wallpaper of the day and sets it as your Windows desktop background. Optionally updates the lock screen too.

## Installation

1. Download `BingWallpaperSetup.exe` from the [latest release](../../releases/latest).
2. Run it. Windows may show a SmartScreen warning — see below.
3. Follow the on-screen prompts. The installer will:
   - Ask whether to also update the lock screen wallpaper.
   - Register a scheduled task that runs at logon.
   - Download and set today's wallpaper immediately.

## Windows SmartScreen warning

When you first run the EXE, Windows Defender SmartScreen may show:

> *"Windows protected your PC — Microsoft Defender SmartScreen prevented an unrecognised app from starting."*

This is expected. The EXE is unsigned (no paid code-signing certificate), so Windows flags it until the file builds a reputation. The source code is fully visible in this repository.

**To proceed:** click **More info**, then **Run anyway**.

## Features

- Downloads the daily Bing wallpaper at logon.
- Saves wallpapers organised by year and month under `Pictures\BingWallpaper\`.
- Optional lock screen wallpaper update (set during installation).
- Retries automatically if the network is unavailable at startup.
- Uninstaller included — find it in `Pictures\BingWallpaper\Uninstall\`.

## Status and management

Open `Pictures\BingWallpaper\View_status.bat` to see:

- Autostart method (scheduled task or startup folder)
- Lock screen: enabled or disabled
- Last run date
- Number of wallpapers saved

To reinstall or change settings, run `BingWallpaperSetup.exe` again.

## Parameters

`BingWallpaperSetup.ps1` accepts optional parameters if you prefer to run the script directly:

| Parameter | Default | Options |
|-----------|---------|---------|
| `-Market` | `en-US` | Any Bing market code, e.g. `en-GB`, `nb-NO` |
| `-Resolution` | `1920x1080` | `1920x1080`, `1366x768`, `3840x2160` |

Example:
```powershell
.\BingWallpaperSetup.ps1 -Market en-GB -Resolution 3840x2160
```

## Uninstall

Run `Uninstall BingWallpaper.bat` inside `Pictures\BingWallpaper\Uninstall\`. This removes the scheduled task and script files. Your saved wallpaper photos and run log are kept.

## Requirements

- Windows 10 or 11
- PowerShell 5.1 (included with Windows)
- Internet access at logon
