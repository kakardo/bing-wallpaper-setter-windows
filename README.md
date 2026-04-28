# Bing Wallpaper Setter for Windows

Sets the Bing wallpaper of the day as your Windows desktop background. Runs automatically at startup. No browser. No ads. No pop-ups. No nonsense.

---

## ⬇️ Download

**[→ Download BingWallpaperSetup.exe](https://github.com/kakardo/bing-wallpaper-setter-windows/releases/latest)**

---

## Install

1. Go to the [Releases](https://github.com/kakardo/bing-wallpaper-setter-windows/releases/latest) page and download `BingWallpaperSetup.exe`
2. Double-click `BingWallpaperSetup.exe`
3. Done — today's Bing wallpaper is set and it will update automatically at every login

After running, you can delete `BingWallpaperSetup.exe`.

### What gets created

```
Pictures\
└── BingWallpaper\
    ├── BingWallpaper.ps1        (runs at startup)
    ├── View_status.bat          (check status at any time)
    ├── run.log                  (one entry per day)
    ├── Uninstall\
    │   └── Uninstall BingWallpaper.bat
    └── YYYY\
        └── MM\
            └── YYYY-MM-DD_Wallpaper title_WIDTHxHEIGHT.jpg
```

---

## Requirements

- Windows 10 or 11
- PowerShell 5.1 or later (built into Windows)

---

## Options

Advanced users can run `BingWallpaperSetup.ps1` directly from PowerShell for a different market or resolution:

```powershell
.\BingWallpaperSetup.ps1 -Market en-GB -Resolution 3840x2160
```

| Parameter | Default | Options |
|---|---|---|
| `-Market` | `en-US` | `en-GB`, `de-DE`, `fr-FR`, etc. |
| `-Resolution` | `1920x1080` | `1920x1080`, `1366x768`, `3840x2160` |

---

## Uninstall

Open `Pictures\BingWallpaper\Uninstall\` and double-click **Uninstall BingWallpaper.bat**. Your saved wallpapers and run log are kept.

---

## How it works

At logon, `BingWallpaper.ps1` calls the Bing image API, downloads today's wallpaper to `Pictures\BingWallpaper\YYYY\MM\`, and applies it via the Windows `SystemParametersInfo` API. If there is no intern