# @author      Kardo Rostam
# @date        2026-04-28
# @version     2.5.1
# @description Setup and management tool for Bing Wallpaper Setter.
#              Installs on first run. Shows status and options if already installed.
#
# Credits
# -------
# Icon: "Bing social network brand" by Darius Dan
#       https://icon-icons.com/icon/bing-social-network-brand-logo/79088
#       Licensed under CC BY 4.0

param(
    [string]$Market = 'en-US',
    [ValidateSet('1920x1080','1366x768','3840x2160')]
    [string]$Resolution = '1920x1080'
)

# Self-elevate if not running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $psArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Market $Market -Resolution $Resolution"
    Start-Process powershell.exe $psArgs -Verb RunAs
    exit
}

# Allow this session to load .ps1 files even when the system policy is Restricted
try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop } catch {}

Clear-Host

$initSpinRs = [runspacefactory]::CreateRunspace(); $initSpinRs.Open()
$initSpinPs = [powershell]::Create(); $initSpinPs.Runspace = $initSpinRs
$initSpinPs.AddScript({ $chars = @('|', '/', '-', '\'); $i = 0; while ($true) { [console]::Write("`r  Loading $($chars[$i++ % 4])"); Start-Sleep -Milliseconds 120 } }) | Out-Null
$initSpinPs.BeginInvoke() | Out-Null
Start-Sleep -Milliseconds 80

# Disable QuickEdit mode so accidental clicks don't pause the script
try {
    Add-Type -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetStdHandle(int n);
[DllImport("kernel32.dll")] public static extern bool GetConsoleMode(IntPtr h, out uint m);
[DllImport("kernel32.dll")] public static extern bool SetConsoleMode(IntPtr h, uint m);
'@ -Name K -Namespace W
    $h = [W.K]::GetStdHandle(-10); $m = 0
    [void][W.K]::GetConsoleMode($h, [ref]$m)
    [void][W.K]::SetConsoleMode($h, $m -band -bnot 0x0040)
} catch {}

$installerVersion = '2.5.1'
$pictures = [Environment]::GetFolderPath('MyPictures')
if (!$pictures -or !(Test-Path $pictures)) { $pictures = Join-Path $env:USERPROFILE 'Pictures' }
if (!$pictures -or !(Test-Path $pictures)) { New-Item -ItemType Directory -Path $pictures -Force | Out-Null }
$installDir  = Join-Path $pictures 'BingWallpaper'
$scriptsDir  = Join-Path $installDir 'Scripts'
$scriptPath   = Join-Path $scriptsDir 'BingWallpaper.ps1'
$launcherPath = Join-Path $scriptsDir 'BingWallpaperLauncher.vbs'
$settingsBat  = Join-Path $installDir 'Settings.bat'
$settingsPs1 = Join-Path $scriptsDir 'Settings.ps1'
$logsDir     = Join-Path $installDir 'Data'
$logFile     = Join-Path $logsDir 'Run.log'
$yesValues   = @('yes','y','1','ja','a','aa')
$noValues    = @('no','n','0','nej','ne','nee')

# - Status check (if already installed) - - - - - - - - - - - #

if (Test-Path $scriptPath) {
    $task           = Get-ScheduledTask -TaskName 'BingWallpaperSetter' -ErrorAction SilentlyContinue
    $startupBatPath = Join-Path ([Environment]::GetFolderPath('Startup')) 'BingWallpaper.bat'
    $hasAutostart   = $task -or (Test-Path $startupBatPath)
    $hasSettings    = Test-Path $settingsPs1
    $partialInstall = -not $hasAutostart -or -not $hasSettings

    $initSpinPs.Stop(); $initSpinPs.Dispose(); $initSpinRs.Close(); $initSpinRs.Dispose()
    $initSpinPs = $null
    [console]::Write("`r              `r")
    Write-Host ''
    Write-Host '  Bing Wallpaper Setter for Windows' -ForegroundColor Cyan
    Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
    Write-Host ''

    if ($partialInstall) {
        Write-Host '  Warning: installation appears incomplete.' -ForegroundColor Yellow
        if (-not $hasSettings)  { Write-Host '  Settings.ps1 is missing.' -ForegroundColor DarkGray }
        if (-not $hasAutostart) { Write-Host '  No autostart method found.' -ForegroundColor DarkGray }
        Write-Host '  A reinstall is recommended.' -ForegroundColor Yellow
    } else {
        Write-Host '  Bing Wallpaper Setter is installed.' -ForegroundColor Green
        Write-Host "  Open Settings.bat in $installDir to manage it."
    }

    Write-Host ''
    Write-Host '  [R] Reinstall   [X] Exit' -ForegroundColor DarkGray
    Write-Host ''

    $choice = (Read-Host '  Choice').Trim().ToUpper()
    if ($choice -ne 'R') { exit }

    Write-Host ''
    $overwriteData = $null
    do {
        Write-Host "  Keep existing stats and log? [Y/n]: " -NoNewline
        $raw = (Read-Host).Trim().ToLower()
        if ($raw -in $yesValues -or $raw -eq '') { $overwriteData = $false }
        elseif ($raw -in $noValues)              { $overwriteData = $true }
        else { Write-Host '  Please enter yes or no.' -ForegroundColor Red }
    } while ($null -eq $overwriteData)
    Write-Host ''
} else {
    $overwriteData = $false
}

# - Embedded wallpaper script - - - - - - - - - - - - - - - - #

$wallpaperScript = @'
# @author      Kardo Rostam
# @date        2026-04-28
# @description Downloads the Bing wallpaper of the day and sets it as the Windows desktop
#              background, and optionally the lock screen.

param(
    [string]$Market = 'en-US',
    [string]$Resolution = '',
    [switch]$SetLockScreen,
    [string]$LogCap = '0',
    [int]$CheckInterval = 60,
    [int]$CheckWindowStart = 0,
    [int]$CheckWindowEnd = 0,
    [switch]$Shuffle,
    [int]$ShuffleInterval = 15,
    [switch]$Install
)

$scriptVersion = '2.5.1'
$logPrefix     = if ($Install) { '[INSTALL] ' } else { '' }

if (!$Resolution) {
    Add-Type -AssemblyName System.Windows.Forms
    $w = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
    $Resolution = if ($w -ge 3840) { '3840x2160' } elseif ($w -ge 1920) { '1920x1080' } else { '1366x768' }
} elseif ($Resolution -notin '1920x1080','1366x768','3840x2160') {
    Write-Log "Error: unsupported resolution '$Resolution'"; exit
}

$wpCode = 'using System; using System.Runtime.InteropServices; [ComImport, Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] public interface IDesktopWallpaper { void SetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID, [MarshalAs(UnmanagedType.LPWStr)] string wallpaper); [return: MarshalAs(UnmanagedType.LPWStr)] string GetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID); [return: MarshalAs(UnmanagedType.LPWStr)] string GetMonitorDevicePathAt(uint monitorIndex); [return: MarshalAs(UnmanagedType.U4)] uint GetMonitorDevicePathCount(); void GetMonitorRECT([MarshalAs(UnmanagedType.LPWStr)] string monitorID, out RECT displayRect); void SetBackgroundColor(uint color); uint GetBackgroundColor(); void SetPosition(int position); int GetPosition(); void SetSlideshow(IntPtr items); IntPtr GetSlideshow(); void SetSlideshowOptions(uint options, uint slideshowTick); void GetSlideshowOptions(out uint options, out uint slideshowTick); void AdvanceSlideshow([MarshalAs(UnmanagedType.LPWStr)] string monitorID, int direction); int GetStatus(); bool Enable(bool enable); } [ComImport, Guid("C2CF3110-460E-4FC1-B9D0-8A1C0C9CC4BD"), ClassInterface(ClassInterfaceType.None)] public class DesktopWallpaperClass {} [StructLayout(LayoutKind.Sequential)] public struct RECT { public int left, top, right, bottom; } public static class WallpaperHelper { public static int SetOnAllMonitors(string path) { try { IDesktopWallpaper dw = (IDesktopWallpaper)(new DesktopWallpaperClass()); uint count = dw.GetMonitorDevicePathCount(); int active = 0; for (uint i = 0; i < count; i++) { try { RECT r; dw.GetMonitorRECT(dw.GetMonitorDevicePathAt(i), out r); if (r.right - r.left > 0 && r.bottom - r.top > 0) { dw.SetWallpaper(dw.GetMonitorDevicePathAt(i), path); active++; } } catch { } } return active; } catch { return 0; } } }'
if (-not ('WallpaperHelper' -as [type])) { Add-Type -TypeDefinition $wpCode }

$installRoot   = Split-Path (Split-Path $MyInvocation.MyCommand.Path)
$logDir        = Join-Path $installRoot 'Data'
$statsFile     = Join-Path $installRoot 'Data\Stats.json'
$manifestFile  = Join-Path $installRoot 'Data\Wallpapers.json'
$updateFile    = Join-Path $installRoot 'Data\UpdateInfo.json'
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$log = Join-Path $logDir 'Run.log'
function Write-Log($msg) {
    "[$([datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))] ${logPrefix}$msg" | Add-Content $log -Encoding UTF8
    if ($LogCap -ne '0' -and (Test-Path $log)) {
        if ($LogCap -match '^(\d+)KB$') {
            $maxBytes = [int]$Matches[1] * 1024
            if ((Get-Item $log).Length -gt $maxBytes) {
                $lines = Get-Content $log
                $lines | Select-Object -Last ([math]::Floor($lines.Count * 0.8)) | Set-Content $log -Encoding UTF8
            }
        } elseif ($LogCap -match '^(\d+)R$') {
            $maxRows = [int]$Matches[1]
            $lines = Get-Content $log
            if ($lines.Count -gt $maxRows) { $lines | Select-Object -Last $maxRows | Set-Content $log -Encoding UTF8 }
        }
    }
}

# Exit early if outside check window or today's wallpaper is already set
if (-not $Install) {
    try {
        if ($CheckWindowStart -ne 0 -or $CheckWindowEnd -ne 0) {
            $currentHour = (Get-Date).Hour
            if ($currentHour -lt $CheckWindowStart -or $currentHour -ge $CheckWindowEnd) { exit }
        }
        $earlyStats = if (Test-Path $statsFile) { Get-Content $statsFile -Raw | ConvertFrom-Json } else { $null }
        $todayDone = $earlyStats -and $earlyStats.LastDownloaded -and $earlyStats.LastDownloaded.Date -eq (Get-Date).ToString('yyyy-MM-dd')
        if ($todayDone -and -not $Shuffle) { exit }
    } catch {}
}

$api = $null
if ($Install) {
    try { $api = Invoke-RestMethod "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=$Market" -ErrorAction Stop } catch {}
    if (!$api) { Write-Log 'Skipped | Network unavailable at install time'; exit }
} else {
    # Retry schedule: 10s x6, 60s x15, 300s x9 (up to ~1 hour total)
    $retrySchedule = @(
        @{ Interval = 10;  Count = 6  },
        @{ Interval = 60;  Count = 15 },
        @{ Interval = 300; Count = 9  }
    )
    $attempt = 0
    foreach ($phase in $retrySchedule) {
        for ($i = 0; $i -lt $phase.Count; $i++) {
            try {
                $api = Invoke-RestMethod "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=$Market" -ErrorAction Stop
                break
            } catch {
                $attempt++
                Write-Log "Network unavailable, retrying in $($phase.Interval)s (attempt $attempt) - $_"
                Start-Sleep -Seconds $phase.Interval
            }
        }
        if ($api) { break }
    }
    if (!$api) { Write-Log 'Skipped | Network unavailable after all retries'; exit }
}

$img  = $api.images[0]
if (!$img -or !$img.urlbase -or !$img.startdate) { Write-Log 'Skipped | Bing returned a malformed response'; exit }

$year  = $img.startdate.Substring(0, 4)
$month = $img.startdate.Substring(4, 2)
$day   = $img.startdate.Substring(6, 2)

$dir = Join-Path $installRoot "Wallpapers\$year\$month"
if (!(Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

$name = if ($img.title) { $img.title -replace '[\\/:*?"<>|]', '_' } else { 'Bing' }
$date = "$year-$month-$day"
$file = "$dir\${date}_${name}_${Resolution}.jpg"

try {
    $isNew = !(Test-Path $file)
    if ($isNew) {
        if (!$Install) { Write-Log 'Started' }
        Invoke-WebRequest "https://www.bing.com$($img.urlbase)_$Resolution.jpg" -OutFile $file -ErrorAction Stop
        if ((Get-Item $file).Length -eq 0) { Remove-Item $file; Write-Log 'Error: downloaded file is empty'; exit }
        Write-Log "Downloaded: ${date}_${name}_${Resolution}.jpg"
        $set = [WallpaperHelper]::SetOnAllMonitors($file)
        if ($set -eq 0) {
            Write-Log 'Error: wallpaper set failed on all monitors'
            Write-Host "Warning: could not set wallpaper on any monitor."
        } else {
            $title = if ($img.title) { $img.title } else { 'Untitled' }
            Write-Log "Wallpaper set | Monitors: $set | `"$title`""
            Write-Host "Wallpaper set on $set monitor(s): `"$title`""
            try {
                $stats = if (Test-Path $statsFile) { Get-Content $statsFile -Raw | ConvertFrom-Json } else { [PSCustomObject]@{ TimesRun = 0; WallpapersSet = 0; FirstRun = ''; LastRun = @{ Date = ''; Time = '' }; WallpaperCount = 0; LastDownloaded = @{ Title = ''; Date = ''; Time = ''; Path = '' }; TimesShuffled = 0; Version = '' } }
                $now   = Get-Date
                $today = $now.ToString('yyyy-MM-dd')
                $stats.TimesRun++
                if ($stats.LastRun.Date -ne $today) { $stats.WallpapersSet++ }
                $stats.LastRun      = [PSCustomObject]@{ Date = $today; Time = $now.ToString('HH:mm:ss') }
                $stats.WallpaperCount++
                $stats.LastDownloaded = [PSCustomObject]@{ Title = $title; Date = $date; Time = $now.ToString('HH:mm:ss'); Path = $file }
                $stats.Version      = $scriptVersion
                $stats | ConvertTo-Json -Depth 3 | Set-Content $statsFile -Encoding UTF8
            } catch {}
            try {
                $mf      = if (Test-Path $manifestFile) { Get-Content $manifestFile -Raw | ConvertFrom-Json } else { [PSCustomObject]@{ Count = 0; HistorySize = 10; History = @(); Wallpapers = @() } }
                $list    = @($mf.Wallpapers)
                $relFile = $file.Substring($installRoot.Length + 1)
                if ($relFile -notin $list) {
                    $list += $relFile
                    $mf.Wallpapers = $list
                    $mf.Count      = $list.Count
                    $mf | ConvertTo-Json -Depth 3 | Set-Content $manifestFile -Encoding UTF8
                }
            } catch {}
        }
    } else {
        if ($Install) { Write-Log 'Already up to date | Wallpaper and lock screen skipped' } else { Write-Log 'Started | Already up to date' }
        Write-Host "Wallpaper is already up to date."
        try {
            $stats = if (Test-Path $statsFile) { Get-Content $statsFile -Raw | ConvertFrom-Json } else { [PSCustomObject]@{ TimesRun = 0; WallpapersSet = 0; FirstRun = ''; LastRun = @{ Date = ''; Time = '' }; WallpaperCount = 0; LastDownloaded = @{ Title = ''; Date = ''; Time = ''; Path = '' }; TimesShuffled = 0; Version = '' } }
            $now   = Get-Date
            $today = $now.ToString('yyyy-MM-dd')
            $stats.TimesRun++
            if ($stats.LastRun.Date -ne $today) { $stats.WallpapersSet++ }
            $stats.LastRun  = [PSCustomObject]@{ Date = $today; Time = $now.ToString('HH:mm:ss') }
            $stats.Version  = $scriptVersion
            $stats | ConvertTo-Json -Depth 3 | Set-Content $statsFile -Encoding UTF8
        } catch {}
    }
    if ($SetLockScreen -and $isNew) {
        try {
            $regPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP'
            if (!(Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
            Set-ItemProperty -Path $regPath -Name 'LockScreenImagePath'   -Value $file
            Set-ItemProperty -Path $regPath -Name 'LockScreenImageUrl'    -Value $file
            Set-ItemProperty -Path $regPath -Name 'LockScreenImageStatus' -Value 1
        } catch {
            Write-Log "Error: lock screen update failed - $_"
            Write-Host "Warning: could not set lock screen: $_"
        }
    } elseif ($isNew -and !$SetLockScreen) {
        if ($Install) { Write-Log 'Lock screen skipped | Not enabled' }
    }
    if ($Shuffle -and -not $Install) {
        try {
            $mf = if (Test-Path $manifestFile) { Get-Content $manifestFile -Raw | ConvertFrom-Json } else { $null }
            if ($mf -and $mf.RecalcInterval -and $mf.RecalcInterval -gt 0) {
                $lastRecalc = if ($mf.LastRecalc) { try { [datetime]::ParseExact($mf.LastRecalc, 'yyyy-MM-dd', $null) } catch { [datetime]::MinValue } } else { [datetime]::MinValue }
                if (((Get-Date) - $lastRecalc).TotalDays -ge $mf.RecalcInterval) {
                    $rjpgs = @(Get-ChildItem (Join-Path $installRoot 'Wallpapers') -Recurse -Include '*.jpg','*.jpeg','*.png','*.bmp' -EA SilentlyContinue | ForEach-Object { $_.FullName.Substring($installRoot.Length + 1) })
                    $mf.Count = $rjpgs.Count; $mf.Wallpapers = $rjpgs; $mf.LastRecalc = (Get-Date).ToString('yyyy-MM-dd')
                    $mf | ConvertTo-Json -Depth 3 | Set-Content $manifestFile -Encoding UTF8
                    Write-Log "Auto-recalculate | $($rjpgs.Count) wallpaper(s) indexed"
                }
            }
            $allJpgs = @(Get-ChildItem (Join-Path $installRoot 'Wallpapers') -Recurse -Include '*.jpg','*.jpeg','*.png','*.bmp' -EA SilentlyContinue | Select-Object -ExpandProperty FullName)
            if ($allJpgs.Count -gt 0) {
                $histSize = if ($mf -and $mf.HistorySize) { $mf.HistorySize } else { 10 }
                $history  = @(if ($mf -and $mf.History) { $mf.History } else { @() })
                $history  = @($history | Where-Object { $_ -lt $allJpgs.Count })
                $maxHist  = [math]::Max(0, [math]::Min($histSize, $allJpgs.Count - 1))
                $history  = @($history | Select-Object -First $maxHist)
                $available = @(0..($allJpgs.Count - 1) | Where-Object { $_ -notin $history })
                $idx = if ($available.Count -gt 0) { $available | Get-Random } else { Get-Random -Maximum $allJpgs.Count }
                $shuffleFile = $allJpgs[$idx]
                if (Test-Path $shuffleFile) {
                    [WallpaperHelper]::SetOnAllMonitors($shuffleFile) | Out-Null
                    Write-Log "Shuffle | $($idx + 1)/$($allJpgs.Count) | `"$(Split-Path $shuffleFile -Leaf)`""
                    if (-not $mf) { $mf = [PSCustomObject]@{ HistorySize = 10; History = @(); RecalcInterval = 7; LastRecalc = '' } }
                    $mf.History = @(@($idx) + $history | Select-Object -First $histSize)
                    $mf | ConvertTo-Json -Depth 3 | Set-Content $manifestFile -Encoding UTF8
                    try {
                        $sStats = if (Test-Path $statsFile) { Get-Content $statsFile -Raw | ConvertFrom-Json } else { $null }
                        if ($sStats) {
                            if (-not $sStats.PSObject.Properties['TimesShuffled']) { $sStats | Add-Member -NotePropertyName TimesShuffled -NotePropertyValue 0 -Force }
                            $sStats.TimesShuffled++
                            $sStats | ConvertTo-Json -Depth 3 | Set-Content $statsFile -Encoding UTF8
                        }
                    } catch {}
                } else {
                    Write-Log "Shuffle | File missing at index $idx, skipping"
                }
            }
        } catch { Write-Log "Shuffle error: $_" }
    }
    if (-not $Install) {
        try {
            $today      = (Get-Date).ToString('yyyy-MM-dd')
            $updateInfo = if (Test-Path $updateFile) { Get-Content $updateFile -Raw | ConvertFrom-Json } else { $null }
            if (-not $updateInfo -or $updateInfo.CheckedAt -ne $today) {
                $rel = Invoke-RestMethod 'https://api.github.com/repos/kakardo/bing-wallpaper-setter-windows/releases/latest' -TimeoutSec 5 -EA Stop
                [PSCustomObject]@{ LatestVersion = ($rel.tag_name -replace '^v', ''); CheckedAt = $today } | ConvertTo-Json | Set-Content $updateFile -Encoding UTF8
            }
        } catch {}
    }
    if ($Install) { Write-Log 'Installation complete' }
} catch {
    Write-Log "Error: $_"
    if (Test-Path $file) { Remove-Item $file }
    exit
}
'@

# - Embedded Settings.bat - - - - - - - - - - - - - - - - - - #

$settingsBatContent = @'
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Settings.ps1"
'@

# - Embedded Settings.ps1 - - - - - - - - - - - - - - - - - - #

$settingsPs1Content = @'
param([string]$InstallDir = (Split-Path $PSScriptRoot))

# Self-elevate if not admin
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Clear-Host

$spinnerRs = [runspacefactory]::CreateRunspace()
$spinnerRs.Open()
$spinnerPs = [powershell]::Create()
$spinnerPs.Runspace = $spinnerRs
$spinnerPs.AddScript({
    $chars = @('|', '/', '-', '\')
    $i = 0
    while ($true) {
        [console]::Write("`r  Loading $($chars[$i++ % 4])")
        Start-Sleep -Milliseconds 120
    }
}) | Out-Null
$spinnerHandle = $spinnerPs.BeginInvoke()
Start-Sleep -Milliseconds 80

# Disable QuickEdit mode so accidental clicks don't pause the script
try {
    Add-Type -MemberDefinition '[DllImport("kernel32.dll")] public static extern IntPtr GetStdHandle(int n); [DllImport("kernel32.dll")] public static extern bool GetConsoleMode(IntPtr h, out uint m); [DllImport("kernel32.dll")] public static extern bool SetConsoleMode(IntPtr h, uint m);' -Name K -Namespace W
    $h = [W.K]::GetStdHandle(-10); $m = 0
    [void][W.K]::GetConsoleMode($h, [ref]$m)
    [void][W.K]::SetConsoleMode($h, $m -band -bnot 0x0040)
} catch {}

$taskName       = 'BingWallpaperSetter'
$scriptPath     = Join-Path $InstallDir 'Scripts\BingWallpaper.ps1'
$launcherPath   = Join-Path $InstallDir 'Scripts\BingWallpaperLauncher.vbs'
$logFile        = Join-Path $InstallDir 'Data\Run.log'
$statsFile      = Join-Path $InstallDir 'Data\Stats.json'
$manifestFile   = Join-Path $InstallDir 'Data\Wallpapers.json'
$updateFile     = Join-Path $InstallDir 'Data\UpdateInfo.json'
$startupBatPath = Join-Path ([Environment]::GetFolderPath('Startup')) 'BingWallpaper.bat'
$yesValues      = @('yes','y','1','ja','a','aa')
$noValues       = @('no','n','0','nej','ne','nee')

$script:cachedConfig    = $null
$script:updateAvailable = $null

function Get-UpdateInfo {
    if ($null -ne $script:updateAvailable) { return }
    try {
        $info   = if (Test-Path $updateFile) { Get-Content $updateFile -Raw | ConvertFrom-Json } else { $null }
        $latest = if ($info -and $info.LatestVersion) { $info.LatestVersion } else { $null }
        $curr   = if ($script:cachedStats -and $script:cachedStats.Version) { $script:cachedStats.Version } else { '0.0.0' }
        $script:updateAvailable = if ($latest -and [System.Version]$latest -gt [System.Version]$curr) { "v$latest" } else { '' }
    } catch {
        $script:updateAvailable = ''
    }
}

function Get-TaskConfig {
    if ($null -ne $script:cachedConfig) { return $script:cachedConfig }
    $task = Get-ScheduledTask -TaskName $taskName -EA SilentlyContinue
    $a = $null; $source = $null
    if ($task) {
        $raw = $task.Actions[0].Arguments
        if ($task.Actions[0].Execute -match 'wscript' -and (Test-Path $launcherPath)) {
            $vbs = Get-Content $launcherPath -Raw
            $a   = if ($vbs -match 'shell\.Run "powershell\.exe (.+)", 0') { $Matches[1] -replace '""', '"' } else { $raw }
        } else {
            $a = $raw
        }
        $source = 'task'
    } elseif (Test-Path $startupBatPath) {
        if (Test-Path $launcherPath) {
            $vbs = Get-Content $launcherPath -Raw
            $a   = if ($vbs -match 'shell\.Run "powershell\.exe (.+)", 0') { $Matches[1] -replace '""', '"' } else { $null }
        }
        if (!$a) { $a = Get-Content $startupBatPath }
        $source = 'startup'
    }
    if (!$a) { return $null }
    $market           = if ($a -match '-Market\s+(\S+)')           { $Matches[1] } else { 'en-US' }
    $resolution       = if ($a -match '-Resolution\s+(\S+)')       { $Matches[1] } else { '' }
    $lockScreen       = [bool]($a -match '-SetLockScreen')
    $logCap           = if ($a -match '-LogCap\s+(\S+)')           { $Matches[1] } else { '0' }
    $checkInterval    = if ($a -match '-CheckInterval\s+(\d+)')    { [int]$Matches[1] } else { 60 }
    $checkWindowStart = if ($a -match '-CheckWindowStart\s+(\d+)') { [int]$Matches[1] } else { 0 }
    $checkWindowEnd   = if ($a -match '-CheckWindowEnd\s+(\d+)')   { [int]$Matches[1] } else { 0 }
    $shuffle          = [bool]($a -match '-Shuffle')
    $shuffleInterval  = if ($a -match '-ShuffleInterval\s+(\d+)')  { [int]$Matches[1] } else { 15 }
    $script:cachedConfig = @{ Market = $market; Resolution = $resolution; LockScreen = $lockScreen; LogCap = $logCap; CheckInterval = $checkInterval; CheckWindowStart = $checkWindowStart; CheckWindowEnd = $checkWindowEnd; Shuffle = $shuffle; ShuffleInterval = $shuffleInterval; Source = $source }
    return $script:cachedConfig
}

function Build-VbsContent($psArgs) {
    $escaped = $psArgs -replace '"', '""'
    return 'Set shell = CreateObject("WScript.Shell")' + "`r`n" + 'shell.Run "powershell.exe ' + $escaped + '", 0, False'
}

function Build-Args($market, $resolution, $lockScreen, $logCap, $checkInterval = 60, $checkWindowStart = 0, $checkWindowEnd = 0, $shuffle = $false, $shuffleInterval = 15) {
    $a = "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Market $market"
    if ($resolution)                                        { $a += " -Resolution $resolution" }
    if ($lockScreen)                                        { $a += ' -SetLockScreen' }
    if ($logCap -and $logCap -ne '0')                      { $a += " -LogCap $logCap" }
    if ($checkInterval -ne 60)                             { $a += " -CheckInterval $checkInterval" }
    if ($checkWindowStart -ne 0 -or $checkWindowEnd -ne 0) { $a += " -CheckWindowStart $checkWindowStart -CheckWindowEnd $checkWindowEnd" }
    if ($shuffle)                                          { $a += " -Shuffle -ShuffleInterval $shuffleInterval" }
    return $a
}

function Update-Task($market, $resolution, $lockScreen, $logCap = '0', $checkInterval = 60, $checkWindowStart = 0, $checkWindowEnd = 0, $shuffle = $null, $shuffleInterval = $null) {
    $current = Get-TaskConfig
    if ($null -eq $shuffle)         { $shuffle         = if ($current) { $current.Shuffle }         else { $false } }
    if ($null -eq $shuffleInterval) { $shuffleInterval = if ($current) { $current.ShuffleInterval } else { 15 } }
    $script:cachedConfig = $null
    $task = Get-ScheduledTask -TaskName $taskName -EA SilentlyContinue
    $interval = if ($shuffle) { $shuffleInterval } else { $checkInterval }
    if ($task) {
        $psArgs        = Build-Args $market $resolution $lockScreen $logCap $checkInterval $checkWindowStart $checkWindowEnd $shuffle $shuffleInterval
        Set-Content -Path $launcherPath -Value (Build-VbsContent $psArgs) -Encoding ASCII
        $runLevel      = if ($lockScreen) { 'Highest' } else { 'Limited' }
        $action        = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument "`"$launcherPath`""
        $triggerLogon  = New-ScheduledTaskTrigger -AtLogOn
        $triggerHourly = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes $interval) -RepetitionDuration (New-TimeSpan -Days 9999)
        $triggers      = @($triggerLogon, $triggerHourly)
        $principal     = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel $runLevel
        try {
            Set-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Principal $principal -EA Stop | Out-Null
            return $true
        } catch {
            Write-Host "  Error updating task: $_" -ForegroundColor Red
            return $false
        }
    } elseif (Test-Path $startupBatPath) {
        $psArgs = Build-Args $market $resolution $lockScreen $logCap $checkInterval $checkWindowStart $checkWindowEnd $shuffle $shuffleInterval
        Set-Content -Path $launcherPath   -Value (Build-VbsContent $psArgs) -Encoding ASCII
        Set-Content -Path $startupBatPath -Value "@echo off`r`nwscript.exe `"$launcherPath`"" -Encoding ASCII
        return $true
    } else {
        Write-Host '  Error: no autostart method found.' -ForegroundColor Red
        return $false
    }
}

function Show-Status {
    Clear-Host
    $cfg  = Get-TaskConfig
    $autostart  = if (!$cfg) { 'Not configured' } elseif ($cfg.Source -eq 'task') { 'Scheduled task' } else { 'Startup folder' }
    $market     = if ($cfg) { $cfg.Market } else { 'Unknown' }
    $resolution = if ($cfg -and $cfg.Resolution) { $cfg.Resolution } else { 'Auto-detect' }
    $lockScreen    = if ($cfg -and $cfg.LockScreen) { 'Enabled' } else { 'Disabled' }
    $logCapDisplay = if (!$cfg -or !$cfg.LogCap -or $cfg.LogCap -eq '0') { 'Off' } `
                     elseif ($cfg.LogCap -match '^(\d+)KB$') { "$($Matches[1]) KB" } `
                     elseif ($cfg.LogCap -match '^(\d+)R$')  { "$($Matches[1]) rows" } `
                     else { 'Off' }
    $stats           = $script:cachedStats
    $wallpaperCount  = if ($stats) { $stats.WallpaperCount } else { 'Unknown (run [C] to calculate)' }
    $wallpapersSet   = if ($stats) { $stats.WallpapersSet }     else { 0 }
    $timesRun        = if ($stats) { $stats.TimesRun }          else { 0 }
    $timesShuffled   = if ($stats) { $stats.TimesShuffled }     else { 0 }
    $firstRun        = if ($stats -and $stats.FirstRun) { $stats.FirstRun } else { 'Unknown' }
    $lastRun         = if ($stats -and $stats.LastRun -and $stats.LastRun.Date) { "$($stats.LastRun.Date) $($stats.LastRun.Time)" } else { 'Never' }
    $dlTitle         = if ($stats -and $stats.LastDownloaded -and $stats.LastDownloaded.Title) { $stats.LastDownloaded.Title } else { $null }
    $dlTitle         = if ($dlTitle -and $dlTitle.Length -gt 40) { $dlTitle.Substring(0, 37) + '...' } else { $dlTitle }
    $lastDownloaded  = if ($dlTitle) { "$dlTitle ($($stats.LastDownloaded.Date) $($stats.LastDownloaded.Time))" } else { 'Never' }
    $versionDisplay       = if ($stats -and $stats.Version) { $stats.Version } else { 'Unknown' }
    $ci                   = if ($cfg) { $cfg.CheckInterval } else { 60 }
    $checkIntervalDisplay = if ($ci -lt 60) { "$ci min" } elseif ($ci -eq 60) { '1 hour' } else { "$([int]($ci / 60)) hours" }
    $cws                  = if ($cfg) { $cfg.CheckWindowStart } else { 0 }
    $cwe                  = if ($cfg) { $cfg.CheckWindowEnd }   else { 0 }
    $checkWindowDisplay   = if ($cws -eq 0 -and $cwe -eq 0) { 'All day' } else { "$($cws.ToString('D2')):00 - $($cwe.ToString('D2')):00" }
    $shuffleOn      = $cfg -and $cfg.Shuffle
    $si             = if ($cfg -and $cfg.ShuffleInterval) { $cfg.ShuffleInterval } else { 15 }
    $siDisplay      = if ($si -lt 60) { "$si min" } elseif ($si -eq 60) { '1 hour' } else { "$([int]($si / 60)) hours" }
    $mfData         = if (Test-Path $manifestFile) { try { Get-Content $manifestFile -Raw | ConvertFrom-Json } catch { $null } } else { $null }
    $mfCount        = if ($mfData) { $mfData.Count } else { 0 }
    $mfHs           = if ($mfData -and $mfData.HistorySize) { $mfData.HistorySize } else { 10 }
    $shuffleDisplay = if ($shuffleOn) { "On (every $siDisplay, history: $mfHs, $mfCount wallpapers)" } else { 'Off' }
    $labelWidth = 14  # "Downloaded  : ".Length
    $sepWidth   = ($labelWidth + (@($autostart, $logCapDisplay, $checkIntervalDisplay, $lastRun, $lastDownloaded, "$wallpapersSet set, $wallpaperCount saved", $shuffleDisplay) | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum)
    $sep = [string][char]0x2500 * $sepWidth
    Write-Host ''
    Write-Host '  Bing Wallpaper Setter for Windows' -ForegroundColor Cyan
    Write-Host "  $sep" -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  Status      : ' -NoNewline; Write-Host 'Installed' -ForegroundColor Green
    Write-Host "  Autostart   : $autostart"
    Write-Host "  Market      : $market"
    Write-Host "  Resolution  : $resolution"
    Write-Host "  Lock screen : $lockScreen"
    Write-Host "  Log cap     : $logCapDisplay"
    Write-Host "  Check every : $checkIntervalDisplay"
    Write-Host "  Check hours : $checkWindowDisplay"
    Write-Host "  Shuffle     : $shuffleDisplay"
    Write-Host ''
    Write-Host '  -- Stats --' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host "  First run   : $firstRun"
    Write-Host "  Checked     : $timesRun"
    Write-Host "  Shuffled    : $timesShuffled"
    Write-Host "  Wallpapers  : $wallpapersSet set, $wallpaperCount saved"
    Write-Host "  Last run    : $lastRun"
    Write-Host "  Downloaded  : $lastDownloaded"
    Write-Host "  Version     : $versionDisplay"
    Get-UpdateInfo
    if ($script:updateAvailable) {
        Write-Host "  Update      : " -NoNewline; Write-Host "$($script:updateAvailable) available" -ForegroundColor Yellow
    }
    Write-Host ''
    Write-Host "  $sep" -ForegroundColor DarkGray
    if ($cfg -and $cfg.Source -eq 'startup') {
        Write-Host ''
        Write-Host '  Note: lock screen requires the scheduled task. Use [T] to switch.' -ForegroundColor Yellow
    }
    Write-Host ''
    Write-Host '  [L] Toggle lock screen   [M] Change market' -ForegroundColor DarkGray
    Write-Host '  [R] Change resolution    [W] Run now' -ForegroundColor DarkGray
    Write-Host '  [G] Log cap              [C] Recalculate stats' -ForegroundColor DarkGray
    Write-Host '  [I] Check interval       [O] Check hours' -ForegroundColor DarkGray
    Write-Host '  [S] Shuffle              [V] View log' -ForegroundColor DarkGray
    Write-Host '  [T] Switch to task       [U] Uninstall' -ForegroundColor DarkGray
    Write-Host '  [X] Exit' -ForegroundColor DarkGray
    Write-Host ''
}

function Toggle-LockScreen {
    $cfg = Get-TaskConfig
    if (!$cfg) { Write-Host '  No autostart configuration found.' -ForegroundColor Red; Start-Sleep 2; return }
    if ($cfg.Source -eq 'startup') {
        Write-Host '  Lock screen update requires the scheduled task (runs elevated).' -ForegroundColor Yellow
        Write-Host '  Reinstall to enable this feature.' -ForegroundColor DarkGray
        Start-Sleep 3; return
    }
    $newLock = -not $cfg.LockScreen
    if (Update-Task $cfg.Market $cfg.Resolution $newLock $cfg.LogCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd) {
        $state = if ($newLock) { 'enabled' } else { 'disabled' }
        Write-Host "  Lock screen $state." -ForegroundColor Green
        if ($newLock) {
            # Check for a new wallpaper first so the lock screen is as up to date as possible.
            Write-Host '  Checking for new wallpaper...' -ForegroundColor DarkGray
            try {
                $psArgs = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Market $($cfg.Market) -SetLockScreen"
                if ($cfg.Resolution) { $psArgs += " -Resolution $($cfg.Resolution)" }
                if ($cfg.LogCap -and $cfg.LogCap -ne '0') { $psArgs += " -LogCap $($cfg.LogCap)" }
                Start-Process powershell -ArgumentList $psArgs -Wait -WindowStyle Hidden
            } catch {}
            # Now set the lock screen to whatever is on the desktop.
            # If a new image was just downloaded the script already set it; this covers the "already up to date" case.
            try {
                $desktopWp = (Get-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name Wallpaper -EA SilentlyContinue).Wallpaper
                $lsStats   = if (Test-Path $statsFile) { Get-Content $statsFile -Raw | ConvertFrom-Json } else { $null }
                $lsFile    = if ($desktopWp -and (Test-Path $desktopWp)) {
                    $desktopWp
                } elseif ($lsStats -and $lsStats.LastDownloaded -and $lsStats.LastDownloaded.Path -and (Test-Path $lsStats.LastDownloaded.Path)) {
                    $lsStats.LastDownloaded.Path
                } else {
                    $null
                }
                if ($lsFile) {
                    $regPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP'
                    if (!(Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
                    Set-ItemProperty -Path $regPath -Name 'LockScreenImagePath'   -Value $lsFile
                    Set-ItemProperty -Path $regPath -Name 'LockScreenImageUrl'    -Value $lsFile
                    Set-ItemProperty -Path $regPath -Name 'LockScreenImageStatus' -Value 1
                    Write-Host "  Lock screen updated to current wallpaper." -ForegroundColor Green
                }
            } catch {
                Write-Host "  Warning: could not update lock screen: $_" -ForegroundColor Yellow
            }
        }
    }
    Start-Sleep 1
}

function Show-MarketMenu {
    $markets = @(
        [PSCustomObject]@{ Code = 'en-US'; Name = 'United States' },
        [PSCustomObject]@{ Code = 'en-GB'; Name = 'United Kingdom' },
        [PSCustomObject]@{ Code = 'nb-NO'; Name = 'Norway' },
        [PSCustomObject]@{ Code = 'sv-SE'; Name = 'Sweden' },
        [PSCustomObject]@{ Code = 'da-DK'; Name = 'Denmark' },
        [PSCustomObject]@{ Code = 'de-DE'; Name = 'Germany' },
        [PSCustomObject]@{ Code = 'fr-FR'; Name = 'France' },
        [PSCustomObject]@{ Code = 'nl-NL'; Name = 'Netherlands' }
    )
    while ($true) {
        $cfg = Get-TaskConfig
        Clear-Host
        Write-Host ''
        Write-Host '  Market' -ForegroundColor Cyan
        Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
        Write-Host "  Current: $($cfg.Market)"
        Write-Host ''
        for ($i = 0; $i -lt $markets.Count; $i++) {
            Write-Host "  [$($i+1)] $($markets[$i].Code)   $($markets[$i].Name)"
        }
        Write-Host '  [C] Custom'
        Write-Host ''
        Write-Host '  [B] Back' -ForegroundColor DarkGray
        Write-Host ''
        $choice = (Read-Host '  Choice').Trim().ToUpper()
        if ($choice -eq 'B') { return }
        $newMarket = $null
        if ($choice -eq 'C') {
            $entry = (Read-Host '  Enter market code (e.g. en-GB)').Trim()
            if ($entry -notmatch '^[a-zA-Z]{2}-[a-zA-Z]{2}$') {
                Write-Host '  Invalid format. Use xx-XX (e.g. en-GB).' -ForegroundColor Red; Start-Sleep 2; continue
            }
            Write-Host '  Validating with Bing...' -ForegroundColor DarkGray
            try {
                $test = Invoke-RestMethod "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=$entry" -ErrorAction Stop
                if ($test.images -and $test.images.Count -gt 0) { $newMarket = $entry }
                else { Write-Host '  Market not recognised by Bing.' -ForegroundColor Red; Start-Sleep 2; continue }
            } catch {
                Write-Host '  Could not validate - check your connection.' -ForegroundColor Red; Start-Sleep 2; continue
            }
        } elseif ($choice -match '^\d+$') {
            $idx = [int]$choice - 1
            if ($idx -ge 0 -and $idx -lt $markets.Count) { $newMarket = $markets[$idx].Code } else { continue }
        } else { continue }
        if ($newMarket) {
            if (Update-Task $newMarket $cfg.Resolution $cfg.LockScreen $cfg.LogCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd) {
                Write-Host "  Market updated to $newMarket." -ForegroundColor Green
            }
            Start-Sleep 1; return
        }
    }
}

function Show-ResolutionMenu {
    while ($true) {
        $cfg     = Get-TaskConfig
        $current = if ($cfg -and $cfg.Resolution) { $cfg.Resolution } else { 'Auto-detect' }
        Clear-Host
        Write-Host ''
        Write-Host '  Resolution' -ForegroundColor Cyan
        Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
        Write-Host "  Current: $current"
        Write-Host ''
        Write-Host '  [1] Auto-detect'
        Write-Host '  [2] 1920x1080  (Full HD)'
        Write-Host '  [3] 3840x2160  (4K)'
        Write-Host '  [4] 1366x768   (HD)'
        Write-Host ''
        Write-Host '  [B] Back' -ForegroundColor DarkGray
        Write-Host ''
        $choice = (Read-Host '  Choice').Trim().ToUpper()
        if ($choice -eq 'B') { return }
        $newRes = switch ($choice) {
            '1' { '' }; '2' { '1920x1080' }; '3' { '3840x2160' }; '4' { '1366x768' }; default { $null }
        }
        if ($null -ne $newRes) {
            if (Update-Task $cfg.Market $newRes $cfg.LockScreen $cfg.LogCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd) {
                $display = if ($newRes) { $newRes } else { 'Auto-detect' }
                Write-Host "  Resolution set to $display." -ForegroundColor Green
            }
            Start-Sleep 1; return
        }
    }
}

function Run-Now {
    Write-Host '  Running wallpaper update...' -ForegroundColor DarkGray
    $cfg    = Get-TaskConfig
    $psArgs = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
    if ($cfg) {
        $psArgs += " -Market $($cfg.Market)"
        if ($cfg.Resolution) { $psArgs += " -Resolution $($cfg.Resolution)" }
        if ($cfg.LockScreen)  { $psArgs += ' -SetLockScreen' }
        if ($cfg.LogCap -and $cfg.LogCap -ne '0') { $psArgs += " -LogCap $($cfg.LogCap)" }
        if ($cfg.Shuffle) { $psArgs += " -Shuffle -ShuffleInterval $($cfg.ShuffleInterval)" }
    }
    Start-Process powershell -ArgumentList $psArgs -Wait -WindowStyle Hidden
    $script:cachedStats = if (Test-Path $statsFile) { Get-Content $statsFile -Raw | ConvertFrom-Json } else { $null }
    Write-Host ''
    if (Test-Path $logFile) {
        Get-Content $logFile | Select-Object -Last 2 | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
    }
    Write-Host ''
    Start-Sleep 2
}

function Invoke-Uninstall {
    Clear-Host
    Write-Host ''
    Write-Host '  Uninstall' -ForegroundColor Cyan
    Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  This removes the scheduled task and scripts.'
    Write-Host '  Your wallpaper photos will be kept.'
    Write-Host ''
    $confirm = (Read-Host '  Type YES to confirm').Trim()
    if ($confirm -ne 'YES') { return }
    Write-Host ''
    do {
        $raw = (Read-Host '  Also delete "Run.log"? [y/N]').Trim().ToLower()
        if ($raw -in $yesValues)          { $deleteLog = 'Y' }
        elseif ($raw -in $noValues -or $raw -eq '') { $deleteLog = 'N' }
        else { Write-Host '  Please enter yes or no.' -ForegroundColor Red; $deleteLog = $null }
    } while ($null -eq $deleteLog)
    do {
        $raw = (Read-Host '  Also delete "Stats.json"? [y/N]').Trim().ToLower()
        if ($raw -in $yesValues)          { $deleteStats = 'Y' }
        elseif ($raw -in $noValues -or $raw -eq '') { $deleteStats = 'N' }
        else { Write-Host '  Please enter yes or no.' -ForegroundColor Red; $deleteStats = $null }
    } while ($null -eq $deleteStats)
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -EA SilentlyContinue
    Remove-Item (Join-Path ([Environment]::GetFolderPath('Startup')) 'BingWallpaper.bat') -EA SilentlyContinue
    Write-Host ''
    Write-Host '  Uninstalled.' -ForegroundColor Green
    if ($deleteLog -eq 'Y' -and $deleteStats -eq 'Y') { Write-Host '  Deleted: "Run.log" and "Stats.json".' -ForegroundColor Green }
    elseif ($deleteLog -eq 'Y')   { Write-Host '  Deleted: "Run.log".' -ForegroundColor Green }
    elseif ($deleteStats -eq 'Y') { Write-Host '  Deleted: "Stats.json".' -ForegroundColor Green }
    Write-Host '  This window will close shortly.' -ForegroundColor Green
    $batPath     = Join-Path $InstallDir 'Settings.bat'
    $scriptsPath = Join-Path $InstallDir 'Scripts'
    $dataPath    = Join-Path $InstallDir 'Data'
    $logPath     = Join-Path $dataPath 'Run.log'
    $statsPath   = Join-Path $dataPath 'Stats.json'
    $manifestPath = Join-Path $dataPath 'Wallpapers.json'
    $updatePath   = Join-Path $dataPath 'UpdateInfo.json'
    $cleanupCmd  = "/c timeout /t 3 /nobreak >nul & del /f /q `"$batPath`" & rmdir /s /q `"$scriptsPath`""
    if ($deleteLog -eq 'Y' -and $deleteStats -eq 'Y') { $cleanupCmd += " & rmdir /s /q `"$dataPath`"" }
    else {
        if ($deleteLog -eq 'Y')   { $cleanupCmd += " & del /f /q `"$logPath`"" }
        if ($deleteStats -eq 'Y') { $cleanupCmd += " & del /f /q `"$statsPath`"" }
        $cleanupCmd += " & del /f /q `"$manifestPath`" & del /f /q `"$updatePath`""
    }
    Start-Process cmd -ArgumentList $cleanupCmd -WindowStyle Hidden
    Start-Sleep 3
    exit
}

function Show-LogCapMenu {
    while ($true) {
        $cfg     = Get-TaskConfig
        $current = if (!$cfg -or !$cfg.LogCap -or $cfg.LogCap -eq '0') { 'Off' } `
                   elseif ($cfg.LogCap -match '^(\d+)KB$') { "$($Matches[1]) KB" } `
                   elseif ($cfg.LogCap -match '^(\d+)R$')  { "$($Matches[1]) rows" } `
                   else { 'Off' }
        Clear-Host
        Write-Host ''
        Write-Host '  Log cap' -ForegroundColor Cyan
        Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
        Write-Host "  Current: $current"
        Write-Host ''
        Write-Host '  [1] Off'
        Write-Host '  [2] By size'
        Write-Host '  [3] By rows'
        Write-Host ''
        Write-Host '  [B] Back' -ForegroundColor DarkGray
        Write-Host ''
        $choice = (Read-Host '  Choice').Trim().ToUpper()
        if ($choice -eq 'B') { return }
        if ($choice -eq '1') {
            if (Update-Task $cfg.Market $cfg.Resolution $cfg.LockScreen '0' $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd) {
                Write-Host '  Log cap disabled.' -ForegroundColor Green
            }
            Start-Sleep 1; return
        }
        if ($choice -eq '2') {
            while ($true) {
                Clear-Host
                Write-Host ''
                Write-Host '  Log cap - Size' -ForegroundColor Cyan
                Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
                Write-Host "  Current: $current"
                Write-Host ''
                Write-Host '  [1] 100 KB'
                Write-Host '  [2] 500 KB'
                Write-Host '  [3] 1 MB'
                Write-Host '  [C] Custom (enter KB)'
                Write-Host ''
                Write-Host '  [B] Back' -ForegroundColor DarkGray
                Write-Host ''
                $sub = (Read-Host '  Choice').Trim().ToUpper()
                if ($sub -eq 'B') { break }
                $newCap = switch ($sub) { '1' { '100KB' }; '2' { '500KB' }; '3' { '1024KB' }; default { $null } }
                if ($sub -eq 'C') {
                    $entry = (Read-Host '  Enter size in KB').Trim()
                    if ($entry -match '^\d+$' -and [int]$entry -gt 0) { $newCap = "${entry}KB" }
                    else { Write-Host '  Invalid value.' -ForegroundColor Red; Start-Sleep 2; continue }
                }
                if ($newCap) {
                    if (Update-Task $cfg.Market $cfg.Resolution $cfg.LockScreen $newCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd) {
                        Write-Host "  Log cap set to $newCap." -ForegroundColor Green
                    }
                    Start-Sleep 1; return
                }
            }
        }
        if ($choice -eq '3') {
            while ($true) {
                Clear-Host
                Write-Host ''
                Write-Host '  Log cap - Rows' -ForegroundColor Cyan
                Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
                Write-Host "  Current: $current"
                Write-Host ''
                Write-Host '  [1] 500 rows'
                Write-Host '  [2] 1 000 rows'
                Write-Host '  [3] 5 000 rows'
                Write-Host '  [C] Custom (enter number)'
                Write-Host ''
                Write-Host '  [B] Back' -ForegroundColor DarkGray
                Write-Host ''
                $sub = (Read-Host '  Choice').Trim().ToUpper()
                if ($sub -eq 'B') { break }
                $newCap = switch ($sub) { '1' { '500R' }; '2' { '1000R' }; '3' { '5000R' }; default { $null } }
                if ($sub -eq 'C') {
                    $entry = (Read-Host '  Enter number of rows').Trim()
                    if ($entry -match '^\d+$' -and [int]$entry -gt 0) { $newCap = "${entry}R" }
                    else { Write-Host '  Invalid value.' -ForegroundColor Red; Start-Sleep 2; continue }
                }
                if ($newCap) {
                    if (Update-Task $cfg.Market $cfg.Resolution $cfg.LockScreen $newCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd) {
                        Write-Host "  Log cap set to $newCap." -ForegroundColor Green
                    }
                    Start-Sleep 1; return
                }
            }
        }
    }
}

function Show-Log {
    if (!(Test-Path $logFile)) { Write-Host '  No log file found.' -ForegroundColor DarkGray; Start-Sleep 2; return }
    $lines = Get-Content $logFile | Select-Object -Last 10
    Write-Host ''
    Write-Host '  Recent log entries' -ForegroundColor Cyan
    Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
    Write-Host ''
    $lines | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
    Write-Host ''
    Read-Host '  Press Enter to return' | Out-Null
}

function Try-ScheduledTask {
    $cfg = Get-TaskConfig
    if (!$cfg) { Write-Host '  No configuration found.' -ForegroundColor Red; Start-Sleep 2; return }
    if ($cfg.Source -eq 'task') { Write-Host '  Already running as a scheduled task.' -ForegroundColor DarkGray; Start-Sleep 2; return }
    Write-Host '  Attempting to register scheduled task...' -ForegroundColor DarkGray
    try {
        $runLevel  = if ($cfg.LockScreen) { 'Highest' } else { 'Limited' }
        $psArgs    = Build-Args $cfg.Market $cfg.Resolution $cfg.LockScreen $cfg.LogCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd
        Set-Content -Path $launcherPath -Value (Build-VbsContent $psArgs) -Encoding ASCII
        $action    = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument "`"$launcherPath`""
        $triggerLogon  = New-ScheduledTaskTrigger -AtLogOn
        $triggerHourly = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Days 9999)
        $triggers      = @($triggerLogon, $triggerHourly)
        $settings      = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
        $principal     = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel $runLevel
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Settings $settings -Principal $principal -EA Stop | Out-Null
        Remove-Item $startupBatPath -EA SilentlyContinue
        Write-Host '  Scheduled task registered. Startup folder entry removed.' -ForegroundColor Green
        Write-Host '  Lock screen control is now available.' -ForegroundColor Green
    } catch {
        Write-Host "  Failed: $_" -ForegroundColor Red
        Write-Host '  Startup folder entry kept.' -ForegroundColor DarkGray
    }
    Start-Sleep 3
}

function Invoke-Recalculate {
    Write-Host '  Recalculating...' -ForegroundColor DarkGray
    $jpgs   = @(Get-ChildItem (Join-Path $InstallDir 'Wallpapers') -Recurse -Include '*.jpg','*.jpeg','*.png','*.bmp' -EA SilentlyContinue | ForEach-Object { $_.FullName.Substring($InstallDir.Length + 1) })
    $count  = $jpgs.Count
    $stats  = if (Test-Path $statsFile) { Get-Content $statsFile -Raw | ConvertFrom-Json } else { [PSCustomObject]@{ TimesRun = 0; WallpapersSet = 0; FirstRun = ''; LastRun = [PSCustomObject]@{ Date = ''; Time = '' }; WallpaperCount = 0; LastDownloaded = [PSCustomObject]@{ Title = ''; Date = ''; Time = ''; Path = '' }; TimesShuffled = 0; Version = '' } }
    $stats.WallpaperCount = $count
    $stats | ConvertTo-Json -Depth 3 | Set-Content $statsFile -Encoding UTF8
    $existing = if (Test-Path $manifestFile) { try { Get-Content $manifestFile -Raw | ConvertFrom-Json } catch { $null } } else { $null }
    $hs = if ($existing -and $existing.HistorySize) { $existing.HistorySize } else { 10 }
    $ri = if ($existing -and $existing.RecalcInterval -ne $null) { $existing.RecalcInterval } else { 7 }
    [PSCustomObject]@{ Count = $count; HistorySize = $hs; History = @(); RecalcInterval = $ri; LastRecalc = (Get-Date).ToString('yyyy-MM-dd'); Wallpapers = $jpgs } | ConvertTo-Json -Depth 3 | Set-Content $manifestFile -Encoding UTF8
    $script:cachedStats = Get-Content $statsFile -Raw | ConvertFrom-Json
    Write-Host "  Done. $count wallpaper(s) indexed." -ForegroundColor Green
    Start-Sleep 1
}

function Show-ShuffleMenu {
    while ($true) {
        $cfg = Get-TaskConfig
        $mf  = if (Test-Path $manifestFile) { try { Get-Content $manifestFile -Raw | ConvertFrom-Json } catch { $null } } else { $null }
        $shuffleOn = $cfg -and $cfg.Shuffle
        $si        = if ($cfg -and $cfg.ShuffleInterval) { $cfg.ShuffleInterval } else { 15 }
        $hs        = if ($mf -and $mf.HistorySize) { $mf.HistorySize } else { 10 }
        $ri        = if ($mf -and $mf.RecalcInterval -ne $null) { $mf.RecalcInterval } else { 7 }
        $riDisplay = if ($ri -eq 0) { 'Off' } elseif ($ri -eq 1) { 'Every day' } else { "Every $ri days" }
        $siDisplay = if ($si -lt 60) { "$si min" } elseif ($si -eq 60) { '1 hour' } else { "$([int]($si / 60)) hours" }
        Clear-Host
        Write-Host ''
        Write-Host '  Shuffle' -ForegroundColor Cyan
        Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
        Write-Host "  Status   : $(if ($shuffleOn) { 'On' } else { 'Off' })"
        Write-Host "  Interval : $siDisplay"
        Write-Host "  History  : $hs wallpapers"
        Write-Host "  Recalc   : $riDisplay"
        Write-Host ''
        Write-Host "  [1] Toggle $(if ($shuffleOn) { 'off' } else { 'on' })"
        Write-Host '  [2] Change interval'
        Write-Host '  [3] Change history size'
        Write-Host '  [4] Recalculate wallpaper list'
        Write-Host '  [5] Change auto-recalculate interval'
        Write-Host ''
        Write-Host '  [B] Back' -ForegroundColor DarkGray
        Write-Host ''
        $choice = (Read-Host '  Choice').Trim().ToUpper()
        if ($choice -eq 'B') { return }
        if ($choice -eq '1') {
            $newShuffle = -not $shuffleOn
            if (Update-Task $cfg.Market $cfg.Resolution $cfg.LockScreen $cfg.LogCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd $newShuffle $si) {
                Write-Host "  Shuffle $(if ($newShuffle) { 'enabled' } else { 'disabled' })." -ForegroundColor Green
            }
            Start-Sleep 1; return
        }
        if ($choice -eq '2') {
            while ($true) {
                Clear-Host
                Write-Host ''
                Write-Host '  Shuffle interval' -ForegroundColor Cyan
                Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
                Write-Host "  Current: $siDisplay"
                Write-Host ''
                Write-Host '  [1] 15 minutes'
                Write-Host '  [2] 30 minutes'
                Write-Host '  [3] 1 hour'
                Write-Host '  [C] Custom (enter minutes)'
                Write-Host ''
                Write-Host '  [B] Back' -ForegroundColor DarkGray
                Write-Host ''
                $sub = (Read-Host '  Choice').Trim().ToUpper()
                if ($sub -eq 'B') { break }
                $newInterval = switch ($sub) { '1' { 15 }; '2' { 30 }; '3' { 60 }; default { $null } }
                if ($sub -eq 'C') {
                    $entry = (Read-Host '  Enter interval in minutes (minimum 5)').Trim()
                    if ($entry -match '^\d+$' -and [int]$entry -ge 5) { $newInterval = [int]$entry }
                    else { Write-Host '  Minimum 5 minutes.' -ForegroundColor Red; Start-Sleep 2; continue }
                }
                if ($null -ne $newInterval) {
                    if (Update-Task $cfg.Market $cfg.Resolution $cfg.LockScreen $cfg.LogCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd $shuffleOn $newInterval) {
                        $d = if ($newInterval -lt 60) { "$newInterval min" } elseif ($newInterval -eq 60) { '1 hour' } else { "$([int]($newInterval / 60)) hours" }
                        Write-Host "  Shuffle interval set to $d." -ForegroundColor Green
                    }
                    Start-Sleep 1; return
                }
            }
        }
        if ($choice -eq '3') {
            while ($true) {
                Clear-Host
                Write-Host ''
                Write-Host '  Shuffle history' -ForegroundColor Cyan
                Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
                Write-Host "  Current: $hs wallpapers"
                Write-Host ''
                Write-Host '  [1] 5'
                Write-Host '  [2] 10'
                Write-Host '  [3] 25'
                Write-Host '  [4] 50'
                Write-Host '  [C] Custom (1-100)'
                Write-Host ''
                Write-Host '  [B] Back' -ForegroundColor DarkGray
                Write-Host ''
                $sub = (Read-Host '  Choice').Trim().ToUpper()
                if ($sub -eq 'B') { break }
                $newHs = switch ($sub) { '1' { 5 }; '2' { 10 }; '3' { 25 }; '4' { 50 }; default { $null } }
                if ($sub -eq 'C') {
                    $entry = (Read-Host '  Enter number (1-100)').Trim()
                    if ($entry -match '^\d+$' -and [int]$entry -ge 1 -and [int]$entry -le 100) { $newHs = [int]$entry }
                    else { Write-Host '  Enter a number between 1 and 100.' -ForegroundColor Red; Start-Sleep 2; continue }
                }
                if ($null -ne $newHs) {
                    try {
                        $mfEdit = if (Test-Path $manifestFile) { Get-Content $manifestFile -Raw | ConvertFrom-Json } else { [PSCustomObject]@{ Count = 0; HistorySize = 10; History = @(); Wallpapers = @() } }
                        $mfEdit.HistorySize = $newHs
                        if ($mfEdit.History -and @($mfEdit.History).Count -gt $newHs) {
                            $mfEdit.History = @($mfEdit.History | Select-Object -First $newHs)
                        }
                        $mfEdit | ConvertTo-Json -Depth 3 | Set-Content $manifestFile -Encoding UTF8
                        Write-Host "  History size set to $newHs." -ForegroundColor Green
                    } catch { Write-Host "  Error: $_" -ForegroundColor Red }
                    Start-Sleep 1; return
                }
            }
        }
        if ($choice -eq '4') { Invoke-Recalculate }
        if ($choice -eq '5') {
            while ($true) {
                Clear-Host
                Write-Host ''
                Write-Host '  Auto-recalculate interval' -ForegroundColor Cyan
                Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
                Write-Host "  Current: $riDisplay"
                Write-Host ''
                Write-Host '  [1] Every day'
                Write-Host '  [2] Every 7 days'
                Write-Host '  [3] Every 30 days'
                Write-Host '  [C] Custom (enter days)'
                Write-Host '  [0] Off'
                Write-Host ''
                Write-Host '  [B] Back' -ForegroundColor DarkGray
                Write-Host ''
                $sub = (Read-Host '  Choice').Trim().ToUpper()
                if ($sub -eq 'B') { break }
                $newRi = switch ($sub) { '1' { 1 }; '2' { 7 }; '3' { 30 }; '0' { 0 }; default { $null } }
                if ($sub -eq 'C') {
                    $entry = (Read-Host '  Enter interval in days (minimum 1)').Trim()
                    if ($entry -match '^\d+$' -and [int]$entry -ge 1) { $newRi = [int]$entry }
                    else { Write-Host '  Minimum 1 day.' -ForegroundColor Red; Start-Sleep 2; continue }
                }
                if ($null -ne $newRi) {
                    try {
                        $mfEdit = if (Test-Path $manifestFile) { Get-Content $manifestFile -Raw | ConvertFrom-Json } else { $null }
                        if ($mfEdit) {
                            if (-not $mfEdit.PSObject.Properties['RecalcInterval']) { $mfEdit | Add-Member -NotePropertyName RecalcInterval -NotePropertyValue $newRi -Force } else { $mfEdit.RecalcInterval = $newRi }
                            $mfEdit | ConvertTo-Json -Depth 3 | Set-Content $manifestFile -Encoding UTF8
                            $d = if ($newRi -eq 0) { 'Off' } elseif ($newRi -eq 1) { 'Every day' } else { "Every $newRi days" }
                            Write-Host "  Auto-recalculate set to: $d." -ForegroundColor Green
                        }
                    } catch { Write-Host "  Error: $_" -ForegroundColor Red }
                    Start-Sleep 1; break
                }
            }
        }
    }
}

function Show-CheckIntervalMenu {
    while ($true) {
        $cfg     = Get-TaskConfig
        $ci      = if ($cfg) { $cfg.CheckInterval } else { 60 }
        $current = if ($ci -lt 60) { "$ci min" } elseif ($ci -eq 60) { '1 hour' } else { "$([int]($ci / 60)) hours" }
        Clear-Host
        Write-Host ''
        Write-Host '  Check interval' -ForegroundColor Cyan
        Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
        Write-Host "  Current: $current"
        Write-Host ''
        Write-Host '  [1] 30 minutes'
        Write-Host '  [2] 1 hour'
        Write-Host '  [3] 2 hours'
        Write-Host '  [4] 4 hours'
        Write-Host '  [C] Custom (enter minutes)'
        Write-Host ''
        Write-Host '  [B] Back' -ForegroundColor DarkGray
        Write-Host ''
        $choice = (Read-Host '  Choice').Trim().ToUpper()
        if ($choice -eq 'B') { return }
        $newInterval = switch ($choice) { '1' { 30 }; '2' { 60 }; '3' { 120 }; '4' { 240 }; default { $null } }
        if ($choice -eq 'C') {
            $entry = (Read-Host '  Enter interval in minutes (minimum 10)').Trim()
            if ($entry -match '^\d+$' -and [int]$entry -ge 10) { $newInterval = [int]$entry }
            else { Write-Host '  Minimum 10 minutes.' -ForegroundColor Red; Start-Sleep 2; continue }
        }
        if ($null -ne $newInterval) {
            if (Update-Task $cfg.Market $cfg.Resolution $cfg.LockScreen $cfg.LogCap $newInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd) {
                $display = if ($newInterval -lt 60) { "$newInterval min" } elseif ($newInterval -eq 60) { '1 hour' } else { "$([int]($newInterval / 60)) hours" }
                Write-Host "  Check interval set to $display." -ForegroundColor Green
            }
            Start-Sleep 1; return
        }
    }
}

function Show-CheckHoursMenu {
    while ($true) {
        $cfg     = Get-TaskConfig
        $cws     = if ($cfg) { $cfg.CheckWindowStart } else { 0 }
        $cwe     = if ($cfg) { $cfg.CheckWindowEnd }   else { 0 }
        $current = if ($cws -eq 0 -and $cwe -eq 0) { 'All day' } else { "$($cws.ToString('D2')):00 - $($cwe.ToString('D2')):00" }
        Clear-Host
        Write-Host ''
        Write-Host '  Check hours' -ForegroundColor Cyan
        Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
        Write-Host "  Current: $current"
        Write-Host ''
        Write-Host '  [1] All day'
        Write-Host '  [2] 06:00 - 23:00'
        Write-Host '  [3] 07:00 - 22:00'
        Write-Host '  [C] Custom (enter start and end hour)'
        Write-Host ''
        Write-Host '  [B] Back' -ForegroundColor DarkGray
        Write-Host ''
        $choice = (Read-Host '  Choice').Trim().ToUpper()
        if ($choice -eq 'B') { return }
        $newStart = $null; $newEnd = $null
        switch ($choice) {
            '1' { $newStart = 0; $newEnd = 0  }
            '2' { $newStart = 6; $newEnd = 23 }
            '3' { $newStart = 7; $newEnd = 22 }
        }
        if ($choice -eq 'C') {
            $s = (Read-Host '  Start hour (0-23)').Trim()
            $e = (Read-Host '  End hour (1-24)').Trim()
            if ($s -match '^\d+$' -and $e -match '^\d+$' -and [int]$s -ge 0 -and [int]$s -le 23 -and [int]$e -ge 1 -and [int]$e -le 24 -and [int]$s -lt [int]$e) {
                $newStart = [int]$s; $newEnd = [int]$e
            } else { Write-Host '  Invalid range. Start must be less than end.' -ForegroundColor Red; Start-Sleep 2; continue }
        }
        if ($null -ne $newStart) {
            if (Update-Task $cfg.Market $cfg.Resolution $cfg.LockScreen $cfg.LogCap $cfg.CheckInterval $newStart $newEnd) {
                $display = if ($newStart -eq 0 -and $newEnd -eq 0) { 'All day' } else { "$($newStart.ToString('D2')):00 - $($newEnd.ToString('D2')):00" }
                Write-Host "  Check hours set to $display." -ForegroundColor Green
            }
            Start-Sleep 1; return
        }
    }
}

$script:cachedStats = if (Test-Path $statsFile) { Get-Content $statsFile -Raw | ConvertFrom-Json } else { $null }
Get-TaskConfig | Out-Null
$spinnerPs.Stop()
$spinnerPs.Dispose()
$spinnerRs.Close()
$spinnerRs.Dispose()
[console]::Write("`r              `r")

# Main loop
while ($true) {
    Show-Status
    $choice = (Read-Host '  Choice').Trim().ToUpper()
    switch ($choice) {
        'L' { Toggle-LockScreen }
        'M' { Show-MarketMenu }
        'R' { Show-ResolutionMenu }
        'W' { Run-Now }
        'G' { Show-LogCapMenu }
        'C' { Invoke-Recalculate }
        'I' { Show-CheckIntervalMenu }
        'O' { Show-CheckHoursMenu }
        'S' { Show-ShuffleMenu }
        'V' { Show-Log }
        'T' { Try-ScheduledTask }
        'U' { Invoke-Uninstall }
        'X' { exit }
    }
}
'@

# - Install - - - - - - - - - - - - - - - - - - - - - - - - - #

function Build-VbsContent($psArgs) {
    $escaped = $psArgs -replace '"', '""'
    return 'Set shell = CreateObject("WScript.Shell")' + "`r`n" + 'shell.Run "powershell.exe ' + $escaped + '", 0, False'
}

try {
    if ($initSpinPs) { $initSpinPs.Stop(); $initSpinPs.Dispose(); $initSpinRs.Close(); $initSpinRs.Dispose(); [console]::Write("`r              `r") }
    Write-Host "Installing Bing Wallpaper Setter..."
    Write-Host ""

    $s1Rs = [runspacefactory]::CreateRunspace(); $s1Rs.Open()
    $s1Ps = [powershell]::Create(); $s1Ps.Runspace = $s1Rs
    $s1Ps.AddScript({ $chars = @('|', '/', '-', '\'); $i = 0; while ($true) { [console]::Write("`r  Step 1: Creating folders $($chars[$i++ % 4])"); Start-Sleep -Milliseconds 120 } }) | Out-Null
    $s1Ps.BeginInvoke() | Out-Null
    Start-Sleep -Milliseconds 80
    $wallpapersDir = Join-Path $installDir 'Wallpapers'
    if (!(Test-Path $installDir))    { New-Item -ItemType Directory -Path $installDir    -Force -ErrorAction Stop | Out-Null }
    if (!(Test-Path $scriptsDir))    { New-Item -ItemType Directory -Path $scriptsDir    -Force -ErrorAction Stop | Out-Null }
    if (!(Test-Path $logsDir))       { New-Item -ItemType Directory -Path $logsDir       -Force -ErrorAction Stop | Out-Null }
    if (!(Test-Path $wallpapersDir)) { New-Item -ItemType Directory -Path $wallpapersDir -Force -ErrorAction Stop | Out-Null }
    $statsPath = Join-Path $logsDir 'Stats.json'
    if ($overwriteData -or !(Test-Path $statsPath)) {
        [PSCustomObject]@{ TimesRun = 0; WallpapersSet = 0; FirstRun = (Get-Date).ToString('yyyy-MM-dd'); LastRun = [PSCustomObject]@{ Date = ''; Time = '' }; WallpaperCount = 0; LastDownloaded = [PSCustomObject]@{ Title = ''; Date = ''; Time = ''; Path = '' }; TimesShuffled = 0; Version = $installerVersion } | ConvertTo-Json -Depth 3 | Set-Content $statsPath -Encoding UTF8
    } else {
        $existing = Get-Content $statsPath -Raw | ConvertFrom-Json
        if ($null -eq $existing.TimesRun)      { $existing | Add-Member -NotePropertyName TimesRun      -NotePropertyValue 0                                                      -Force }
        if ($null -eq $existing.WallpapersSet) { $existing | Add-Member -NotePropertyName WallpapersSet -NotePropertyValue 0                                                      -Force }
        if (-not $existing.FirstRun)           { $existing | Add-Member -NotePropertyName FirstRun      -NotePropertyValue (Get-Date).ToString('yyyy-MM-dd')                      -Force }
        if (-not $existing.PSObject.Properties['LastRun'] -or -not $existing.LastRun) {
            $existing | Add-Member -NotePropertyName LastRun -NotePropertyValue ([PSCustomObject]@{ Date = ''; Time = '' }) -Force
        } else {
            if ($null -eq $existing.LastRun.Date) { $existing.LastRun | Add-Member -NotePropertyName Date -NotePropertyValue '' -Force }
            if ($null -eq $existing.LastRun.Time) { $existing.LastRun | Add-Member -NotePropertyName Time -NotePropertyValue '' -Force }
        }
        if ($null -eq $existing.WallpaperCount)   { $existing | Add-Member -NotePropertyName WallpaperCount  -NotePropertyValue 0 -Force }
        if ($null -eq $existing.TimesShuffled)   { $existing | Add-Member -NotePropertyName TimesShuffled   -NotePropertyValue 0 -Force }
        if (-not $existing.PSObject.Properties['LastDownloaded'] -or -not $existing.LastDownloaded) {
            $existing | Add-Member -NotePropertyName LastDownloaded -NotePropertyValue ([PSCustomObject]@{ Title = ''; Date = ''; Time = '' }) -Force
        } else {
            if ($null -eq $existing.LastDownloaded.Title) { $existing.LastDownloaded | Add-Member -NotePropertyName Title -NotePropertyValue '' -Force }
            if ($null -eq $existing.LastDownloaded.Date)  { $existing.LastDownloaded | Add-Member -NotePropertyName Date  -NotePropertyValue '' -Force }
            if ($null -eq $existing.LastDownloaded.Time)  { $existing.LastDownloaded | Add-Member -NotePropertyName Time  -NotePropertyValue '' -Force }
            if ($null -eq $existing.LastDownloaded.Path)  { $existing.LastDownloaded | Add-Member -NotePropertyName Path  -NotePropertyValue '' -Force }
        }
        $existing.Version = $installerVersion
        $existing | ConvertTo-Json -Depth 3 | Set-Content $statsPath -Encoding UTF8
    }
    $manifestPath    = Join-Path $logsDir 'Wallpapers.json'
    $existingMf      = if (!$overwriteData -and (Test-Path $manifestPath)) { try { Get-Content $manifestPath -Raw | ConvertFrom-Json } catch { $null } } else { $null }
    $mfHs            = if ($existingMf -and $existingMf.HistorySize)     { $existingMf.HistorySize }     else { 10 }
    $mfHistory       = if ($existingMf -and $existingMf.History)         { $existingMf.History }         else { @() }
    $mfRi            = if ($existingMf -and $existingMf.RecalcInterval -ne $null) { $existingMf.RecalcInterval } else { 7 }
    $existingJpgs    = @(Get-ChildItem $wallpapersDir -Recurse -Include '*.jpg','*.jpeg','*.png','*.bmp' -EA SilentlyContinue | ForEach-Object { $_.FullName.Substring($installDir.Length + 1) })
    [PSCustomObject]@{ Count = $existingJpgs.Count; HistorySize = $mfHs; History = $mfHistory; RecalcInterval = $mfRi; LastRecalc = ''; Wallpapers = $existingJpgs } | ConvertTo-Json -Depth 3 | Set-Content $manifestPath -Encoding UTF8
    $updatePath = Join-Path $logsDir 'UpdateInfo.json'
    if ($overwriteData -or !(Test-Path $updatePath)) {
        [PSCustomObject]@{ LatestVersion = ''; CheckedAt = '' } | ConvertTo-Json | Set-Content $updatePath -Encoding UTF8
    } else {
        $existingUpd = Get-Content $updatePath -Raw | ConvertFrom-Json
        if ($null -eq $existingUpd.LatestVersion) { $existingUpd | Add-Member -NotePropertyName LatestVersion -NotePropertyValue '' -Force }
        if ($null -eq $existingUpd.CheckedAt)     { $existingUpd | Add-Member -NotePropertyName CheckedAt     -NotePropertyValue '' -Force }
        $existingUpd | ConvertTo-Json | Set-Content $updatePath -Encoding UTF8
    }
    if ($overwriteData) { Clear-Content $logFile -ErrorAction SilentlyContinue }
    "[$([datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))] [INSTALL] Installation started" | Add-Content $logFile -Encoding UTF8
    $s1Ps.Stop(); $s1Ps.Dispose(); $s1Rs.Close(); $s1Rs.Dispose()
    [console]::Write("`r                                `r")
    Write-Host "Step 1: Folders ready."

    $s2Rs = [runspacefactory]::CreateRunspace(); $s2Rs.Open()
    $s2Ps = [powershell]::Create(); $s2Ps.Runspace = $s2Rs
    $s2Ps.AddScript({ $chars = @('|', '/', '-', '\'); $i = 0; while ($true) { [console]::Write("`r  Step 2: Writing scripts $($chars[$i++ % 4])"); Start-Sleep -Milliseconds 120 } }) | Out-Null
    $s2Ps.BeginInvoke() | Out-Null
    Start-Sleep -Milliseconds 80
    Set-Content -Path $scriptPath   -Value $wallpaperScript    -Encoding UTF8  -ErrorAction Stop
    Set-Content -Path $settingsBat  -Value $settingsBatContent -Encoding ASCII -ErrorAction Stop
    Set-Content -Path $settingsPs1  -Value $settingsPs1Content -Encoding UTF8  -ErrorAction Stop
    $s2Ps.Stop(); $s2Ps.Dispose(); $s2Rs.Close(); $s2Rs.Dispose()
    [console]::Write("`r                             `r")
    Write-Host "Step 2: Scripts written."

    Write-Host ""
    $setLockScreen = $null
    do {
        Write-Host "  Also update lock screen wallpaper? [Y/n]: " -NoNewline
        $raw = (Read-Host).Trim().ToLower()
        if ($raw -in $yesValues -or $raw -eq '') { $setLockScreen = $true }
        elseif ($raw -in $noValues)              { $setLockScreen = $false }
        else { Write-Host '  Please enter yes or no.' -ForegroundColor Red }
    } while ($null -eq $setLockScreen)
    Write-Host ""

    $psArgs   = "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Market $Market"
    if ($PSBoundParameters.ContainsKey('Resolution')) { $psArgs += " -Resolution $Resolution" }
    if ($setLockScreen) { $psArgs += ' -SetLockScreen' }
    Set-Content -Path $launcherPath -Value (Build-VbsContent $psArgs) -Encoding ASCII -ErrorAction Stop
    $taskName = 'BingWallpaperSetter'
    $taskDone = $false

    $instSpinRs = [runspacefactory]::CreateRunspace(); $instSpinRs.Open()
    $instSpinPs = [powershell]::Create(); $instSpinPs.Runspace = $instSpinRs
    $instSpinPs.AddScript({ $chars = @('|', '/', '-', '\'); $i = 0; while ($true) { [console]::Write("`r  Step 3: Registering autostart $($chars[$i++ % 4])"); Start-Sleep -Milliseconds 120 } }) | Out-Null
    $instSpinPs.BeginInvoke() | Out-Null
    Start-Sleep -Milliseconds 80

    try {
        $runLevel  = if ($setLockScreen) { 'Highest' } else { 'Limited' }
        $action    = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument "`"$launcherPath`""
        $triggerLogon  = New-ScheduledTaskTrigger -AtLogOn
        $triggerHourly = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Days 9999)
        $triggers      = @($triggerLogon, $triggerHourly)
        $settings      = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
        $principal     = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel $runLevel
        if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
            Set-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Settings $settings -Principal $principal -ErrorAction Stop | Out-Null
        } else {
            Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Settings $settings -Principal $principal -ErrorAction Stop | Out-Null
        }
        $instSpinPs.Stop(); $instSpinPs.Dispose(); $instSpinRs.Close(); $instSpinRs.Dispose()
        [console]::Write("`r                                              `r")
        Write-Host "Step 3: Autostart registered."
        $taskDone = $true
    } catch {
        $instSpinPs.Stop(); $instSpinPs.Dispose(); $instSpinRs.Close(); $instSpinRs.Dispose()
        [console]::Write("`r                                              `r")
        Write-Host "Step 3: Scheduled task blocked - using startup folder instead."
    }

    $startupBatPath = Join-Path ([Environment]::GetFolderPath('Startup')) 'BingWallpaper.bat'
    if (!$taskDone) {
        Set-Content -Path $startupBatPath -Value "@echo off`r`nwscript.exe `"$launcherPath`"" -Encoding ASCII -ErrorAction Stop
        Write-Host "Added to startup folder."
    }

    Write-Host ""
    Write-Host "Installed to: $installDir"
    Write-Host "Open Settings.bat in the BingWallpaper folder to manage settings or uninstall."
    Write-Host ""

    $s35Rs = [runspacefactory]::CreateRunspace(); $s35Rs.Open()
    $s35Ps = [powershell]::Create(); $s35Ps.Runspace = $s35Rs
    $s35Ps.AddScript({ $chars = @('|', '/', '-', '\'); $i = 0; while ($true) { [console]::Write("`r  Step 3.5: Verifying $($chars[$i++ % 4])"); Start-Sleep -Milliseconds 120 } }) | Out-Null
    $s35Ps.BeginInvoke() | Out-Null
    Start-Sleep -Milliseconds 80
    function Write-InstallLog($msg) {
        "[$([datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))] [INSTALL] $msg" | Add-Content $logFile -Encoding UTF8
    }
    $checks = 0; $passed = 0

    $checks++
    if (Test-Path $scriptPath) {
        Write-InstallLog 'Check: BingWallpaper.ps1 exists ... OK'; $passed++
    } else {
        Write-InstallLog 'Check: BingWallpaper.ps1 exists ... NOT FOUND'
    }

    $checks++
    if (Test-Path $settingsPs1) {
        Write-InstallLog 'Check: Settings.ps1 exists ... OK'; $passed++
    } else {
        Write-InstallLog 'Check: Settings.ps1 exists ... NOT FOUND'
    }

    $checks++
    if (Test-Path $settingsBat) {
        Write-InstallLog 'Check: Settings.bat exists ... OK'; $passed++
    } else {
        Write-InstallLog 'Check: Settings.bat exists ... NOT FOUND'
    }

    $checks++
    $verifyTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($verifyTask) {
        Write-InstallLog "Check: Scheduled task ($($verifyTask.State)) ... OK"; $passed++
    } elseif (Test-Path $startupBatPath) {
        Write-InstallLog 'Check: Scheduled task ... NOT FOUND (startup folder active)'; $passed++
    } else {
        Write-InstallLog 'Check: Autostart ... NOT CONFIGURED'
    }

    if ($verifyTask) {
        $checks++
        $vbsOk = (Test-Path $launcherPath) -and ((Get-Content $launcherPath -Raw) -match [regex]::Escape($scriptPath))
        if ($vbsOk) {
            Write-InstallLog 'Check: Launcher script path matches ... OK'; $passed++
        } else {
            Write-InstallLog 'Check: Launcher script path ... MISMATCH or launcher missing'
        }
    }

    $s35Ps.Stop(); $s35Ps.Dispose(); $s35Rs.Close(); $s35Rs.Dispose()
    [console]::Write("`r                             `r")
    if ($passed -eq $checks) {
        Write-InstallLog "Verification passed ($passed/$checks)"
        Write-Host "Step 3.5: Verification passed ($passed/$checks)."
    } else {
        Write-InstallLog "Verification completed with warnings ($passed/$checks)"
        Write-Host "Step 3.5: Verification warnings ($passed/$checks) - check the log." -ForegroundColor Yellow
    }
    Write-Host ""

    $s4Rs = [runspacefactory]::CreateRunspace(); $s4Rs.Open()
    $s4Ps = [powershell]::Create(); $s4Ps.Runspace = $s4Rs
    $s4Ps.AddScript({ $chars = @('|', '/', '-', '\'); $i = 0; while ($true) { [console]::Write("`r  Step 4: Downloading wallpaper $($chars[$i++ % 4])"); Start-Sleep -Milliseconds 120 } }) | Out-Null
    $s4Ps.BeginInvoke() | Out-Null
    Start-Sleep -Milliseconds 80
    if ($setLockScreen) {
        if ($PSBoundParameters.ContainsKey('Resolution')) { & $scriptPath -Market $Market -Resolution $Resolution -SetLockScreen -Install 6>$null } else { & $scriptPath -Market $Market -SetLockScreen -Install 6>$null }
    } else {
        if ($PSBoundParameters.ContainsKey('Resolution')) { & $scriptPath -Market $Market -Resolution $Resolution -Install 6>$null } else { & $scriptPath -Market $Market -Install 6>$null }
    }
    $s4Ps.Stop(); $s4Ps.Dispose(); $s4Rs.Close(); $s4Rs.Dispose()
    [console]::Write("`r                                     `r")
    $lastLog = if (Test-Path $logFile) { Get-Content $logFile | Select-Object -Last 1 } else { '' }
    if ($lastLog -match 'Network unavailable at install time') {
        Write-Host "Step 4: Network unavailable, will retry at logon." -ForegroundColor Yellow
    } else {
        $dlTitle = if ($lastLog -match '"([^"]+)"') { $Matches[1] } else { $null }
        Write-Host "Step 4: Wallpaper set." -ForegroundColor Green
        if ($dlTitle) { Write-Host "  $dlTitle" -ForegroundColor DarkGray }
    }

    Write-Host ""
    Write-Host "  Installation successful!" -ForegroundColor Green
    Write-Host ""
    Read-Host "  Press Enter to close"

} catch {
    Write-Host ""
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host ""
    Read-Host "  Press Enter to close"
}
                                          