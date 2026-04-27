# bing-wallpaper-setter-windows

Automatic daily wallpaper setter for Windows. Retrieves the Bing wallpaper of the day and sets it as your desktop background — no browser pop-ups, no ads, no accidental clicks opening Bing.

---

## Files

| File | Purpose |
|---|---|
| `Setup.ps1` | One-time setup — installs everything and registers the scheduled task |

---

## Requirements

- Windows 10 or 11
- PowerShell 5.1 or later (built into Windows)
- Internet access at logon

---

## Quick start

1. **Download** `Setup.ps1` from this repository.

2. **Right-click** the file and choose **Run with PowerShell**.

   Setup will:
   - Create `Pictures\BingWallpaper\`
   - Install the wallpaper script there
   - Add a one-click uninstaller to `Pictures\BingWallpaper\Uninstall\`
   - Register a scheduled task so the wallpaper updates at every logon
   - Set today's wallpaper immediately

That's it. You can delete `Setup.ps1` after running it.

---

## Options

| Parameter | Default | Description |
|---|---|---|
| `-Market` | `en-US` | Bing market code (e.g. `en-GB`, `de-DE`, `fr-FR`) |
| `-Resolution` | `1920x1080` | Image resolution: `1920x1080`, `1366x768`, or `3840x2160` (UHD) |

To install with custom options, run from PowerShell instead:

```powershell
.\Setup.ps1 -Market en-GB -Resolution 3840x2160
```

---

## Uninstall

Open `Pictures\BingWallpaper\Uninstall\` in Explorer and double-click **Uninstall BingWallpaper.bat**.

This removes the scheduled task and the wallpaper script. Your saved wallpaper photos are kept.

---

## How it works

1. Queries `https://www.bing.com/HPImageArchive.aspx` for today's image metadata.
2. Downloads the full-resolution JPEG to `Pictures\BingWallpaper\YYYY\MM\` (skips if already present).
3. Calls the Windows `SystemParametersInfo` API to apply the image as the desktop wallpaper.
