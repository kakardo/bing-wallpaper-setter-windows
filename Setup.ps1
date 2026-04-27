# @author      Kardo Rostam
# @date        2026-04-27
# @description Installs Bing Wallpaper Setter. Creates the wallpaper directory in Pictures,
#              writes the wallpaper script, adds a one-click uninstaller, and registers
#              a scheduled task so the wallpaper updates at every logon.

Clear-Host

param(
    [string]$Market = 'en-US',
    [ValidateSet('1920x1080','1366x768','3840x2160')]
    [string]$Resolution = '1920x1080'
)

$taskName     = 'BingWallpaperSetter'
$pictures     = [Environment]::GetFolderPath('MyPictures')
if (!$pictures -or !(Test-Path $pictures)) { $pictures = "$env:USERPROFILE\Pictures" }
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

Add-Type @"
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern int SystemParametersInfo(int a, int b, string c, int d);
}
"@

$api = Invoke-RestMethod "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=$Market"
$img = $api.images[0]

$year     = $img.startdate.Substring(0, 4)
$month    = $img.startdate.Substring(4, 2)
$pictures = [Environment]::GetFolderPath('MyPictures')
if (!$pictures -or !(Test-Path $pictures)) { $pictures = "$env:USERPROFILE\Pictures" }
$dir      = Join-Path $pictures "BingWallpaper\$year\$month"
if (!(Test-Path $dir)) { New-Item -ItemType Directory $dir | Out-Null }

$name = $img.title -replace '[\\/:*?"<>|]', '_'
$date = "$year-$month-$($img.startdate.Substring(6, 2))"
$file = "$dir\${date}_${name}_${Resolution}.jpg"

if (!(Test-Path $file)) {
    Invoke-WebRequest "https://www.bing.com$($img.urlbase)_$Resolution.jpg" -OutFile $file
}

[Win32]::SystemParametersInfo(20, 0, $file, 3) | Out-Null
Write-Host "Wallpaper set: $($img.title)"
'@

# - Embedded uninstall script - - - - - - - - - - - - - - - - #

$uninstallScript = @'
@echo off
echo Removing Bing Wallpaper Setter...
powershell -NonInteractive -ExecutionPolicy Bypass -Command ^
  "Unregister-ScheduledTask -TaskName 'BingWallpaperSetter' -Confirm:$false -ErrorAction SilentlyContinue; ^
   Remove-Item ([Environment]::GetFolderPath('MyPictures') + '\BingWallpaper\BingWallpaper.ps1') -ErrorAction SilentlyContinue"
echo.
echo Done. Your wallpaper photos have been kept.
echo.
pause
cd /d "%USERPROFILE%\Pictures\BingWallpaper"
rmdir /s /q "%~dp0"
'@

# - Install - - - - - - - - - - - - - - - - - - - - - - - - - #
Write-Host "Installing Bing Wallpaper Setter..."
Write-Host ""

if (!(Test-Path $installDir))   { New-Item -ItemType Directory $installDir   | Out-Null }
if (!(Test-Path $uninstallDir)) { New-Item -ItemType Directory $uninstallDir | Out-Null }

Set-Content -Path $scriptPath   -Value $wallpaperScript -Encoding UTF8
Set-Content -Path $uninstallBat -Value $uninstallScript -Encoding ASCII

$action   = New-ScheduledTaskAction -Execute 'powershell.exe' `
                -Argument "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Market $Market -Resolution $Resolution"
$trigger  = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -StartWhenAvailable

if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Set-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings | Out-Null
    Write-Host "Scheduled task updated."
} else {
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings | Out-Null
    Write-Host "Scheduled task created."
}

Write-Host ""
Write-Host "Installed to: $installDir"
Write-Host "To uninstall: open the Uninstall folder inside BingWallpaper in Pictures."
Write-Host ""

& $scriptPath -Market $Market -Resolution $Resolution
