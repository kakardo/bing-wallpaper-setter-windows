# bing-wallpaper-setter-windows

Automatic daily wallpaper setter for Windows. Retrieves the Bing wallpaper of the day and sets it as your desktop background — no browser pop-ups, no ads, no accidental clicks opening Bing.

---

## Files

| File | Purpose |
|---|---|
| `Set-BingWallpaper.ps1` | Downloads the latest Bing wallpaper and applies it |
| `Install-Task.ps1` | Registers a Task Scheduler job so it runs automatically at logon |

---

## Requirements

- Windows 10 or 11
- PowerShell 5.1 or later (built into Windows)
- Internet access at logon

---

## Quick start

1. **Clone or download** this repository to a permanent folder, e.g. `C:\Tools\BingWallpaper`.

2. **Open PowerShell** in that folder and run the setup script once:

   ```powershell
   .\Install-Task.ps1
   ```

   This registers a scheduled task that runs `Set-BingWallpaper.ps1` silently every time you log in. No admin rights required.

3. **Test it immediately** (optional):

   ```powershell
   Start-ScheduledTask -TaskName 'BingWallpaperSetter'
   ```

That's it. Your wallpaper will update automatically from now on.

---

## Options

Both scripts accept optional parameters:

| Parameter | Default | Description |
|---|---|---|
| `-Market` | `en-US` | Bing market code (e.g. `en-GB`, `de-DE`, `fr-FR`) |
| `-Resolution` | `1920x1080` | Image resolution: `1920x1080`, `1366x768`, or `3840x2160` (UHD) |

**Example — UK market at 4K:**

```powershell
.\Install-Task.ps1 -Market en-GB -Resolution 3840x2160
```

---

## Running manually

You can run the wallpaper setter at any time without the scheduled task:

```powershell
.\Set-BingWallpaper.ps1
```

---

## Uninstall

To remove the scheduled task:

```powershell
.\Install-Task.ps1 -Uninstall
```

Downloaded wallpapers are stored in `%APPDATA%\BingWallpaper` and can be deleted manually.

---

## How it works

1. Queries `https://www.bing.com/HPImageArchive.aspx` for today's image metadata.
2. Downloads the full-resolution JPEG to `%APPDATA%\BingWallpaper` (skips if already present).
3. Calls the Windows `SystemParametersInfo` API to apply the image as the desktop wallpaper.
