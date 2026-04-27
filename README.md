# Bing Wallpaper Setter for Windows

Sets the Bing wallpaper of the day as your Windows desktop background. Runs automatically at startup. No browser. No ads. No pop-ups. No nonsense.

---

## Install

1. Download `Setup.ps1`
2. Right-click it and choose **Run with PowerShell**
3. Done

After running, you can delete `Setup.ps1`.

### What gets created

```
Pictures\
└── BingWallpaper\
    ├── BingWallpaper.ps1        (runs at startup)
    ├── Uninstall\
    │   └── Uninstall BingWallpaper.bat
    └── YYYY\
        └── MM\
            └── YYYY-MM-DD_Wallpaper title_WIDTHxHEIGHT.jpg
```

A shortcut is also added to your Windows startup folder so the script runs every time you log in.

---

## Requirements

- Windows 10 or 11
- PowerShell 5.1 or later (built into Windows)

---

## Options

Run from PowerShell if you want a different market or resolution:

```powershell
.\Setup.ps1 -Market en-GB -Resolution 3840x2160
```

| Parameter | Default | Options |
|---|---|---|
| `-Market` | `en-US` | `en-GB`, `de-DE`, `fr-FR`, etc. |
| `-Resolution` | `1920x1080` | `1920x1080`, `1366x768`, `3840x2160` |

---

## Uninstall

Open `Pictures\BingWallpaper\Uninstall\` and double-click **Uninstall BingWallpaper.bat**. Your saved wallpapers are kept.

---

## How it works

At logon, `BingWallpaper.ps1` calls the Bing image API, downloads today's wallpaper to `Pictures\BingWallpaper\YYYY\MM\`, and applies it via the Windows `SystemParametersInfo` API. If there is no internet connection yet, it retries every 10 seconds for the first minute, then every minute for 15 minutes, then every 5 minutes up to 1 hour before giving up.
