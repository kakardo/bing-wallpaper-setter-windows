# @author      Kardo Rostam
# @date        2026-04-28
# @version     2.4
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

# Disable QuickEdit mode so accidental clicks don't pause the script
try {
    $k = Add-Type -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetStdHandle(int n);
[DllImport("kernel32.dll")] public static extern bool GetConsoleMode(IntPtr h, out uint m);
[DllImport("kernel32.dll")] public static extern bool SetConsoleMode(IntPtr h, uint m);
'@ -Name K -Namespace W -PassThru
    $h = [W.K]::GetStdHandle(-10); $m = 0
    [W.K]::GetConsoleMode($h, [ref]$m)
    [W.K]::SetConsoleMode($h, $m -band -bnot 0x0040)
} catch {}

# Allow this session to load .ps1 files even when the system policy is Restricted
try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop } catch {}

Clear-Host

$pictures = [Environment]::GetFolderPath('MyPictures')
if (!$pictures -or !(Test-Path $pictures)) { $pictures = Join-Path $env:USERPROFILE 'Pictures' }
if (!$pictures -or !(Test-Path $pictures)) { New-Item -ItemType Directory -Path $pictures -Force | Out-Null }
$installDir  = Join-Path $pictures 'BingWallpaper'
$scriptsDir  = Join-Path $installDir 'Scripts'
$scriptPath  = Join-Path $scriptsDir 'BingWallpaper.ps1'
$settingsBat = Join-Path $installDir 'Settings.bat'
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
    [switch]$Install
)

$scriptVersion = '2.4'
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

$installRoot = Split-Path (Split-Path $MyInvocation.MyCommand.Path)
$logDir    = Join-Path $installRoot 'Data'
$statsFile = Join-Path $installRoot 'Data\Stats.json'
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
        if ($earlyStats -and $earlyStats.LastDownloaded -and $earlyStats.LastDownloaded.Date -eq (Get-Date).ToString('yyyy-MM-dd')) { exit }
    } catch {}
}

# Retry schedule: 10s x6, 60s x15, 300s x9 (up to ~1 hour total)
$retrySchedule = @(
    @{ Interval = 10;  Count = 6  },
    @{ Interval = 60;  Count = 15 },
    @{ Interval = 300; Count = 9  }
)

$api = $null
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

if (!$api) { Write-Log 'Error: API unreachable after all retries'; exit }

$img  = $api.images[0]

$pics = [Environment]::GetFolderPath('MyPictures')
if (!$pics -or !(Test-Path $pics)) { $pics = Join-Path $env:USERPROFILE 'Pictures' }
if (!$pics) { exit }

$year  = $img.startdate.Substring(0, 4)
$month = $img.startdate.Substring(4, 2)
$day   = $img.startdate.Substring(6, 2)

$dir = Join-Path $pics "BingWallpaper\$year\$month"
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
                $stats = if (Test-Path $statsFile) { Get-Content $statsFile -Raw | ConvertFrom-Json } else { [PSCustomObject]@{ TimesRun = 0; WallpapersSet = 0; FirstRun = ''; LastRun = @{ Date = ''; Time = '' }; WallpaperCount = 0; LastDownloaded = @{ Title = ''; Date = ''; Time = '' }; Version = '' } }
                $now   = Get-Date
                $today = $now.ToString('yyyy-MM-dd')
                $stats.TimesRun++
                if ($stats.LastRun.Date -ne $today) { $stats.WallpapersSet++ }
                $stats.LastRun      = [PSCustomObject]@{ Date = $today; Time = $now.ToString('HH:mm:ss') }
                $stats.WallpaperCount++
                $stats.LastDownloaded = [PSCustomObject]@{ Title = $title; Date = $date; Time = $now.ToString('HH:mm:ss') }
                $stats.Version      = $scriptVersion
                $stats | ConvertTo-Json -Depth 3 | Set-Content $statsFile -Encoding UTF8
            } catch {}
        }
    } else {
        if ($Install) { Write-Log 'Already up to date | Wallpaper and lock screen skipped' } else { Write-Log 'Started | Already up to date' }
        Write-Host "Wallpaper is already up to date."
        try {
            $stats = if (Test-Path $statsFile) { Get-Content $statsFile -Raw | ConvertFrom-Json } else { [PSCustomObject]@{ TimesRun = 0; WallpapersSet = 0; FirstRun = ''; LastRun = @{ Date = ''; Time = '' }; WallpaperCount = 0; LastDownloaded = @{ Title = ''; Date = ''; Time = '' }; Version = '' } }
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

# Disable QuickEdit mode so accidental clicks don't pause the script
try {
    Add-Type -MemberDefinition '[DllImport("kernel32.dll")] public static extern IntPtr GetStdHandle(int n); [DllImport("kernel32.dll")] public static extern bool GetConsoleMode(IntPtr h, out uint m); [DllImport("kernel32.dll")] public static extern bool SetConsoleMode(IntPtr h, uint m);' -Name K -Namespace W
    $h = [W.K]::GetStdHandle(-10); $m = 0
    [W.K]::GetConsoleMode($h, [ref]$m)
    [W.K]::SetConsoleMode($h, $m -band -bnot 0x0040)
} catch {}

$taskName       = 'BingWallpaperSetter'
$scriptPath     = Join-Path $InstallDir 'Scripts\BingWallpaper.ps1'
$logFile        = Join-Path $InstallDir 'Data\Run.log'
$statsFile      = Join-Path $InstallDir 'Data\Stats.json'
$startupBatPath = Join-Path ([Environment]::GetFolderPath('Startup')) 'BingWallpaper.bat'

$script:cachedConfig = $null

function Get-TaskConfig {
    if ($null -ne $script:cachedConfig) { return $script:cachedConfig }
    $task = Get-ScheduledTask -TaskName $taskName -EA SilentlyContinue
    $a = $null; $source = $null
    if ($task) {
        $a = $task.Actions[0].Arguments; $source = 'task'
    } elseif (Test-Path $startupBatPath) {
        $a = Get-Content $startupBatPath; $source = 'startup'
    }
    if (!$a) { return $null }
    $market           = if ($a -match '-Market\s+(\S+)')           { $Matches[1] } else { 'en-US' }
    $resolution       = if ($a -match '-Resolution\s+(\S+)')       { $Matches[1] } else { '' }
    $lockScreen       = [bool]($a -match '-SetLockScreen')
    $logCap           = if ($a -match '-LogCap\s+(\S+)')           { $Matches[1] } else { '0' }
    $checkInterval    = if ($a -match '-CheckInterval\s+(\d+)')    { [int]$Matches[1] } else { 60 }
    $checkWindowStart = if ($a -match '-CheckWindowStart\s+(\d+)') { [int]$Matches[1] } else { 0 }
    $checkWindowEnd   = if ($a -match '-CheckWindowEnd\s+(\d+)')   { [int]$Matches[1] } else { 0 }
    $script:cachedConfig = @{ Market = $market; Resolution = $resolution; LockScreen = $lockScreen; LogCap = $logCap; CheckInterval = $checkInterval; CheckWindowStart = $checkWindowStart; CheckWindowEnd = $checkWindowEnd; Source = $source }
    return $script:cachedConfig
}

function Build-Args($market, $resolution, $lockScreen, $logCap, $checkInterval = 60, $checkWindowStart = 0, $checkWindowEnd = 0) {
    $a = "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Market $market"
    if ($resolution)                                        { $a += " -Resolution $resolution" }
    if ($lockScreen)                                        { $a += ' -SetLockScreen' }
    if ($logCap -and $logCap -ne '0')                      { $a += " -LogCap $logCap" }
    if ($checkInterval -ne 60)                             { $a += " -CheckInterval $checkInterval" }
    if ($checkWindowStart -ne 0 -or $checkWindowEnd -ne 0) { $a += " -CheckWindowStart $checkWindowStart -CheckWindowEnd $checkWindowEnd" }
    return $a
}

function Update-Task($market, $resolution, $lockScreen, $logCap = '0', $checkInterval = 60, $checkWindowStart = 0, $checkWindowEnd = 0) {
    $script:cachedConfig = $null
    $task = Get-ScheduledTask -TaskName $taskName -EA SilentlyContinue
    if ($task) {
        $runLevel      = if ($lockScreen) { 'Highest' } else { 'Limited' }
        $action        = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument (Build-Args $market $resolution $lockScreen $logCap $checkInterval $checkWindowStart $checkWindowEnd)
        $triggerLogon  = New-ScheduledTaskTrigger -AtLogOn
        $triggerHourly = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes $checkInterval) -RepetitionDuration (New-TimeSpan -Days 9999)
        $triggers      = @($triggerLogon, $triggerHourly)
        $principal     = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel $runLevel
        Set-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Principal $principal -EA Stop | Out-Null
    } elseif (Test-Path $startupBatPath) {
        Set-Content -Path $startupBatPath -Value "powershell.exe $(Build-Args $market $resolution $lockScreen $logCap $checkInterval $checkWindowStart $checkWindowEnd)" -Encoding ASCII
    } else {
        Write-Host '  Error: no autostart method found.' -ForegroundColor Red; Start-Sleep 2
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
    $wallpapersSet   = if ($stats) { $stats.WallpapersSet } else { 0 }
    $timesRun        = if ($stats) { $stats.TimesRun } else { 0 }
    $firstRun        = if ($stats -and $stats.FirstRun) { $stats.FirstRun } else { 'Unknown' }
    $lastRun         = if ($stats -and $stats.LastRun -and $stats.LastRun.Date) { "$($stats.LastRun.Date) $($stats.LastRun.Time)" } else { 'Never' }
    $lastDownloaded  = if ($stats -and $stats.LastDownloaded -and $stats.LastDownloaded.Title) { "$($stats.LastDownloaded.Title) ($($stats.LastDownloaded.Date) $($stats.LastDownloaded.Time))" } else { 'Never' }
    $versionDisplay       = if ($stats -and $stats.Version) { $stats.Version } else { 'Unknown' }
    $ci                   = if ($cfg) { $cfg.CheckInterval } else { 60 }
    $checkIntervalDisplay = if ($ci -lt 60) { "$ci min" } elseif ($ci -eq 60) { '1 hour' } else { "$([int]($ci / 60)) hours" }
    $cws                  = if ($cfg) { $cfg.CheckWindowStart } else { 0 }
    $cwe                  = if ($cfg) { $cfg.CheckWindowEnd }   else { 0 }
    $checkWindowDisplay   = if ($cws -eq 0 -and $cwe -eq 0) { 'All day' } else { "$($cws.ToString('D2')):00 - $($cwe.ToString('D2')):00" }
    Write-Host ''
    Write-Host '  Bing Wallpaper Setter for Windows' -ForegroundColor Cyan
    Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  Status      : ' -NoNewline; Write-Host 'Installed' -ForegroundColor Green
    Write-Host "  Autostart   : $autostart"
    Write-Host "  Market      : $market"
    Write-Host "  Resolution  : $resolution"
    Write-Host "  Lock screen : $lockScreen"
    Write-Host "  Log cap     : $logCapDisplay"
    Write-Host "  Check every : $checkIntervalDisplay"
    Write-Host "  Check hours : $checkWindowDisplay"
    Write-Host ''
    Write-Host '  -- Stats --' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host "  First run   : $firstRun"
    Write-Host "  Times run   : $timesRun"
    Write-Host "  Wallpapers  : $wallpapersSet set, $wallpaperCount saved"
    Write-Host "  Last run    : $lastRun"
    Write-Host "  Downloaded  : $lastDownloaded"
    Write-Host "  Version     : $versionDisplay"
    Write-Host ''
    Write-Host ('  ' + ([string][char]0x2500 * 36)) -ForegroundColor DarkGray
    if ($cfg -and $cfg.Source -eq 'startup') {
        Write-Host ''
        Write-Host '  Note: running via startup folder. Lock screen control' -ForegroundColor Yellow
        Write-Host '  is unavailable. Use [T] to try switching to scheduled task.' -ForegroundColor Yellow
    }
    Write-Host ''
    Write-Host '  [L] Toggle lock screen   [M] Change market' -ForegroundColor DarkGray
    Write-Host '  [R] Change resolution    [W] Run now' -ForegroundColor DarkGray
    Write-Host '  [G] Log cap              [C] Recalculate stats' -ForegroundColor DarkGray
    Write-Host '  [I] Check interval       [O] Check hours' -ForegroundColor DarkGray
    Write-Host '  [U] Uninstall' -ForegroundColor DarkGray
    if ($cfg -and $cfg.Source -eq 'startup') {
        Write-Host '  [T] Try scheduled task   [X] Exit' -ForegroundColor DarkGray
    } else {
        Write-Host '  [X] Exit' -ForegroundColor DarkGray
    }
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
    Update-Task $cfg.Market $cfg.Resolution $newLock $cfg.LogCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd
    $state = if ($newLock) { 'enabled' } else { 'disabled' }
    Write-Host "  Lock screen $state." -ForegroundColor Green
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
            Update-Task $newMarket $cfg.Resolution $cfg.LockScreen $cfg.LogCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd
            Write-Host "  Market updated to $newMarket." -ForegroundColor Green
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
            Update-Task $cfg.Market $newRes $cfg.LockScreen $cfg.LogCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd
            $display = if ($newRes) { $newRes } else { 'Auto-detect' }
            Write-Host "  Resolution set to $display." -ForegroundColor Green
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
    $cleanupCmd  = "/c timeout /t 3 /nobreak >nul & del /f /q `"$batPath`" & rmdir /s /q `"$scriptsPath`""
    if ($deleteLog -eq 'Y' -and $deleteStats -eq 'Y') { $cleanupCmd += " & rmdir /s /q `"$dataPath`"" }
    elseif ($deleteLog -eq 'Y')   { $cleanupCmd += " & del /f /q `"$logPath`"" }
    elseif ($deleteStats -eq 'Y') { $cleanupCmd += " & del /f /q `"$statsPath`"" }
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
            Update-Task $cfg.Market $cfg.Resolution $cfg.LockScreen '0' $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd
            Write-Host '  Log cap disabled.' -ForegroundColor Green
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
                    Update-Task $cfg.Market $cfg.Resolution $cfg.LockScreen $newCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd
                    Write-Host "  Log cap set to $newCap." -ForegroundColor Green
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
                    Update-Task $cfg.Market $cfg.Resolution $cfg.LockScreen $newCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd
                    Write-Host "  Log cap set to $newCap." -ForegroundColor Green
                    Start-Sleep 1; return
                }
            }
        }
    }
}

function Try-ScheduledTask {
    $cfg = Get-TaskConfig
    if (!$cfg) { Write-Host '  No configuration found.' -ForegroundColor Red; Start-Sleep 2; return }
    Write-Host '  Attempting to register scheduled task...' -ForegroundColor DarkGray
    try {
        $runLevel  = if ($cfg.LockScreen) { 'Highest' } else { 'Limited' }
        $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument (Build-Args $cfg.Market $cfg.Resolution $cfg.LockScreen $cfg.LogCap $cfg.CheckInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd)
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
    $count  = (Get-ChildItem $InstallDir -Recurse -Filter '*.jpg' -EA SilentlyContinue).Count
    $stats  = if (Test-Path $statsFile) { Get-Content $statsFile -Raw | ConvertFrom-Json } else { [PSCustomObject]@{ TimesRun = 0; WallpapersSet = 0; FirstRun = ''; LastRun = [PSCustomObject]@{ Date = ''; Time = '' }; WallpaperCount = 0; LastDownloaded = [PSCustomObject]@{ Title = ''; Date = ''; Time = '' }; Version = '' } }
    $stats.WallpaperCount = $count
    $stats | ConvertTo-Json -Depth 3 | Set-Content $statsFile -Encoding UTF8
    $script:cachedStats = Get-Content $statsFile -Raw | ConvertFrom-Json
    Write-Host '  Done.' -ForegroundColor Green
    Start-Sleep 1
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
            Update-Task $cfg.Market $cfg.Resolution $cfg.LockScreen $cfg.LogCap $newInterval $cfg.CheckWindowStart $cfg.CheckWindowEnd
            $display = if ($newInterval -lt 60) { "$newInterval min" } elseif ($newInterval -eq 60) { '1 hour' } else { "$([int]($newInterval / 60)) hours" }
            Write-Host "  Check interval set to $display." -ForegroundColor Green
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
            Update-Task $cfg.Market $cfg.Resolution $cfg.LockScreen $cfg.LogCap $cfg.CheckInterval $newStart $newEnd
            $display = if ($newStart -eq 0 -and $newEnd -eq 0) { 'All day' } else { "$($newStart.ToString('D2')):00 - $($newEnd.ToString('D2')):00" }
            Write-Host "  Check hours set to $display." -ForegroundColor Green
            Start-Sleep 1; return
        }
    }
}

$script:cachedStats = if (Test-Path $statsFile) { Get-Content $statsFile -Raw | ConvertFrom-Json } else { $null }

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
        'T' { Try-ScheduledTask }
        'U' { Invoke-Uninstall }
        'X' { exit }
    }
}
'@

# - Install - - - - - - - - - - - - - - - - - - - - - - - - - #

try {
    Write-Host "Installing Bing Wallpaper Setter..."
    Write-Host ""

    Write-Host "Step 1: Creating folders..."
    if (!(Test-Path $installDir))   { New-Item -ItemType Directory -Path $installDir   -Force -ErrorAction Stop | Out-Null }
    if (!(Test-Path $scriptsDir))   { New-Item -ItemType Directory -Path $scriptsDir   -Force -ErrorAction Stop | Out-Null }
    if (!(Test-Path $logsDir))      { New-Item -ItemType Directory -Path $logsDir      -Force -ErrorAction Stop | Out-Null }
    $statsPath = Join-Path $logsDir 'Stats.json'
    if ($overwriteData -or !(Test-Path $statsPath)) {
        [PSCustomObject]@{ TimesRun = 0; WallpapersSet = 0; FirstRun = (Get-Date).ToString('yyyy-MM-dd'); LastRun = [PSCustomObject]@{ Date = ''; Time = '' }; WallpaperCount = 0; LastDownloaded = [PSCustomObject]@{ Title = ''; Date = ''; Time = '' }; Version = '' } | ConvertTo-Json -Depth 3 | Set-Content $statsPath -Encoding UTF8
    }
    if ($overwriteData) { Clear-Content $logFile -ErrorAction SilentlyContinue }
    "[$([datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))] [INSTALL] Installation started" | Add-Content $logFile -Encoding UTF8

    Write-Host "Step 2: Writing scripts..."
    Set-Content -Path $scriptPath   -Value $wallpaperScript    -Encoding UTF8  -ErrorAction Stop
    Set-Content -Path $settingsBat  -Value $settingsBatContent -Encoding ASCII -ErrorAction Stop
    Set-Content -Path $settingsPs1  -Value $settingsPs1Content -Encoding UTF8  -ErrorAction Stop

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

    Write-Host "Step 3: Registering autostart..."
    $psArgs   = "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Market $Market"
    if ($PSBoundParameters.ContainsKey('Resolution')) { $psArgs += " -Resolution $Resolution" }
    if ($setLockScreen) { $psArgs += ' -SetLockScreen' }
    $taskName = 'BingWallpaperSetter'
    $taskDone = $false

    try {
        $runLevel  = if ($setLockScreen) { 'Highest' } else { 'Limited' }
        $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $psArgs
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
        Write-Host "Scheduled task registered."
        $taskDone = $true
    } catch {
        Write-Host "Scheduled task blocked - using startup folder instead."
    }

    $startupBatPath = Join-Path ([Environment]::GetFolderPath('Startup')) 'BingWallpaper.bat'
    if (!$taskDone) {
        $startupContent = "powershell.exe -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Market $Market"
        if ($PSBoundParameters.ContainsKey('Resolution')) { $startupContent += " -Resolution $Resolution" }
        if ($setLockScreen) { $startupContent += ' -SetLockScreen' }
        Set-Content -Path $startupBatPath -Value $startupContent -Encoding ASCII -ErrorAction Stop
        Write-Host "Added to startup folder."
    }

    Write-Host ""
    Write-Host "Installed to: $installDir"
    Write-Host "Open Settings.bat in the BingWallpaper folder to manage settings or uninstall."
    Write-Host ""

    Write-Host "Step 3.5: Verifying installation..."
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
        $taskArgs = $verifyTask.Actions[0].Arguments
        if ($taskArgs -match [regex]::Escape($scriptPath)) {
            Write-InstallLog 'Check: Task script path matches ... OK'; $passed++
        } else {
            Write-InstallLog "Check: Task script path ... MISMATCH (task has $($taskArgs -replace '.*-File\s+\"?([^\"]+)\"?.*','$1'))"
        }
    }

    if ($passed -eq $checks) {
        Write-InstallLog "Verification passed ($passed/$checks)"
        Write-Host "Verification passed ($passed/$checks)."
    } else {
        Write-InstallLog "Verification completed with warnings ($passed/$checks)"
        Write-Host "Verification completed with warnings ($passed/$checks) - check the log." -ForegroundColor Yellow
    }
    Write-Host ""

    Write-Host "Step 4: Setting today's wallpaper..."
    if ($setLockScreen) {
        if ($PSBoundParameters.ContainsKey('Resolution')) { & $scriptPath -Market $Market -Resolution $Resolution -SetLockScreen -Install } else { & $scriptPath -Market $Market -SetLockScreen -Install }
    } else {
        if ($PSBoundParameters.ContainsKey('Resolution')) { & $scriptPath -Market $Market -Resolution $Resolution -Install } else { & $scriptPath -Market $Market -Install }
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
