# @author      Kardo Rostam
# @date        2026-04-27
# @description Installs Bing Wallpaper Setter. Creates the wallpaper directory in Pictures,
#              writes the wallpaper script, adds a one-click uninstaller, and sets up autostart.

param(
    [string]$Market = 'en-US',
    [ValidateSet('1920x1080','1366x768','3840x2160')]
    [string]$Resolution = '1920x1080'
)

# Self-elevate if not running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $args = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Market $Market -Resolution $Resolution"
    Start-Process powershell.exe $args -Verb RunAs
    exit
}

Clear-Host

$pictures = [Environment]::GetFolderPath('MyPictures')
if (!$pictures -or !(Test-Path $pictures)) { $pictures = Join-Path $env:USERPROFILE 'Pictures' }
if (!$pictures -or !(Test-Path $pictures)) { New-Item -ItemType Directory -Path $pictures -Force | Out-Null }
$installDir   = Join-Path $pictures 'BingWallpaper'
$scriptPath   = Join-Path $installDir 'BingWallpaper.ps1'
$uninstallDir = Join-Path $installDir 'Uninstall'
$uninstallBat = Join-Path $uninstallDir 'Uninstall BingWallpaper.bat'

# - Embedded wallpaper script - - - - - - - - - - - - - - - - #

$wallpaperScript = @'
# @author      Kardo Rostam
# @date        2026-04-27
# @description Downloads the Bing wallpaper of the day and sets it as the Windows desktop background.

param(
    [string]$Market = 'en-US',
    [ValidateSet('1920x1080','1366x768','3840x2160')]
    [string]$Resolution = '1920x1080'
)

$code = 'using System.Runtime.InteropServices; public class Win32 { [DllImport("user32.dll")] public static extern int SystemParametersInfo(int a, int b, string c, int d); }'
Add-Type -TypeDefinition $code

# Retry schedule: 10s x6, 60s x15, 300s x9 (up to ~1 hour total)
$retrySchedule = @(
    @{ Interval = 10;  Count = 6  },
    @{ Interval = 60;  Count = 15 },
    @{ Interval = 300; Count = 9  }
)

$api = $null
foreach ($phase in $retrySchedule) {
    for ($i = 0; $i -lt $phase.Count; $i++) {
        try {
            $api = Invoke-RestMethod "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=$Market" -ErrorAction Stop
            break
        } catch {
            Start-Sleep -Seconds $phase.Interval
        }
    }
    if ($api) { break }
}

if (!$api) { exit }

$img  = $api.images[0]

$year     = $img.startdate.Substring(0, 4)
$month    = $img.startdate.Substring(4, 2)
$pics = [Environment]::GetFolderPath('MyPictures')
if (!$pics -or !(Test-Path $pics)) { $pics = Join-Path $env:USERPROFILE 'Pictures' }
if (!$pics) { exit }
$dir  = Join-Path $pics "BingWallpaper\$year\$month"
if (!(Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

$name = if ($img.title) { $img.title -replace '[\\/:*?"<>|]', '_' } else { 'Bing' }
$date = "$year-$month-$($img.startdate.Substring(6, 2))"
$file = "$dir\${date}_${name}_${Resolution}.jpg"

try {
    if (!(Test-Path $file)) {
        Invoke-WebRequest "https://www.bing.com$($img.urlbase)_$Resolution.jpg" -OutFile $file -ErrorAction Stop
        if ((Get-Item $file).Length -eq 0) { Remove-Item $file; exit }
    }
    $result = [Win32]::SystemParametersInfo(20, 0, $file, 3)
    if ($result -eq 0) { Write-Host "Warning: wallpaper set call returned failure." }
    else { Write-Host "Wallpaper set: $($img.title)" }
} catch {
    if (Test-Path $file) { Remove-Item $file }
    exit
}
'@

# - Embedded uninstall script - - - - - - - - - - - - - - - - #

$uninstallScript = @'
@echo off
echo Removing Bing Wallpaper Setter...
powershell -NonInteractive -ExecutionPolicy Bypass -Command "Unregister-ScheduledTask -TaskName 'BingWallpaperSetter' -Confirm:$false -ErrorAction SilentlyContinue; Remove-Item ([Environment]::GetFolderPath('Startup') + '\BingWallpaper.bat') -ErrorAction SilentlyContinue; Remove-Item ([Environment]::GetFolderPath('MyPictures') + '\BingWallpaper\BingWallpaper.ps1') -ErrorAction SilentlyContinue"
echo.
echo Done. Your wallpaper photos have been kept.
echo.
pause
cd /d "%USERPROFILE%"
rmdir /s /q "%~dp0"
'@

# - Install - - - - - - - - - - - - - - - - - - - - - - - - - #

try {
    Write-Host "Installing Bing Wallpaper Setter..."
    Write-Host ""

    Write-Host "Step 1: Creating folders..."
    if (!(Test-Path $installDir))   { New-Item -ItemType Directory -Path $installDir   -Force -ErrorAction Stop | Out-Null }
    if (!(Test-Path $uninstallDir)) { New-Item -ItemType Directory -Path $uninstallDir -Force -ErrorAction Stop | Out-Null }

    Write-Host "Step 2: Writing scripts..."
    Set-Content -Path $scriptPath   -Value $wallpaperScript -Encoding UTF8  -ErrorAction Stop
    Set-Content -Path $uninstallBat -Value $uninstallScript -Encoding ASCII -ErrorAction Stop

    Write-Host "Step 3: Registering autostart..."
    $psArgs   = "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Market $Market -Resolution $Resolution"
    $taskName = 'BingWallpaperSetter'
    $taskDone = $false

    try {
        $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $psArgs
        $trigger   = New-ScheduledTaskTrigger -AtLogOn
        $settings  = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -StartWhenAvailable
        $principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
        if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
            Set-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -ErrorAction Stop | Out-Null
        } else {
            Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -ErrorAction Stop | Out-Null
        }
        Write-Host "Scheduled task registered."
        $taskDone = $true
    } catch {
        Write-Host "Scheduled task blocked - using startup folder instead."
    }

    if (!$taskDone) {
        $startupBat     = Join-Path ([Environment]::GetFolderPath('Startup')) 'BingWallpaper.bat'
        $startupContent = "powershell.exe -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Market $Market -Resolution $Resolution"
        Set-Content -Path $startupBat -Value $startupContent -Encoding ASCII -ErrorAction Stop
        Write-Host "Added to startup folder."
    }

    Write-Host ""
    Write-Host "Installed to: $installDir"
    Write-Host "To uninstall: open the Uninstall folder inside BingWallpaper in Pictures."
    Write-Host ""

    Write-Host "Step 4: Setting today's wallpaper..."
    & $scriptPath -Market $Market -Resolution $Resolution

} catch {
    Write-Host ""
    Write-Host "ERROR: $_" -ForegroundColor Red
}