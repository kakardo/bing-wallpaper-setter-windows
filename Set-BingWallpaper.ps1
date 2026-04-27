# @author      Kardo Rostam
# @date        2026-04-27
# @description Downloads the Bing wallpaper of the day and sets it as the Windows desktop background.

Clear-Host

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

$year    = $img.startdate.Substring(0, 4)
$month   = $img.startdate.Substring(4, 2)
$pictures = [Environment]::GetFolderPath('MyPictures')
if (!$pictures -or !(Test-Path $pictures)) { $pictures = "$env:USERPROFILE\Pictures" }
$dir     = Join-Path $pictures "BingWallpaper\$year\$month"
if (!(Test-Path $dir)) { New-Item -ItemType Directory $dir | Out-Null }

$name = $img.title -replace '[\\/:*?"<>|]', '_'
$date = "$year-$month-$($img.startdate.Substring(6, 2))"
$file = "$dir\${date}_${name}_${Resolution}.jpg"

if (!(Test-Path $file)) {
    Invoke-WebRequest "https://www.bing.com$($img.urlbase)_$Resolution.jpg" -OutFile $file
}

[Win32]::SystemParametersInfo(20, 0, $file, 3) | Out-Null
Write-Host "Wallpaper set: $($img.title)"
