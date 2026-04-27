# Bing Wallpaper Setter for Windows

Sets the Bing wallpaper of the day as your Windows desktop background. Runs automatically at startup. No browser. No ads. No pop-ups. No nonsense.

---

## ⬇️ Download

**[→ Download Setup.exe](../../releases/latest)**

---

## Install

1. Go to the [Releases](../../releases/latest) page and download `Setup.exe`
2. Double-click `Setup.exe`
3. Done — today's Bing wallpaper is set and it will update automatically at every login

After running, you can delete `Setup.exe`.

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

Advanced users can run `Setup.ps1` directly from PowerShell for a different market or resolution:

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

At logon, `BingWallpaper.ps1` calls the Bing image API, downloads today's wallpaper to `Pictures\BingWallpaper\YYYY\MM\`, and applies it via the Windows `SystemParametersInfo` API. If there is no internet connection yet, it retries every 10 seconds for the first minute, then every minute for 15 minutes, then every 5 minutes for up to an hour.