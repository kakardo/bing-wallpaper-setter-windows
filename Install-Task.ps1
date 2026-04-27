param(
    [string]$Market = 'en-US',
    [ValidateSet('1920x1080','1366x768','3840x2160')]
    [string]$Resolution = '1920x1080',
    [switch]$Uninstall
)

$taskName  = 'BingWallpaperSetter'
$scriptPath = Join-Path $PSScriptRoot 'Set-BingWallpaper.ps1'

if ($Uninstall) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Task removed."
    return
}

if (!(Test-Path $scriptPath)) {
    Write-Error "Set-BingWallpaper.ps1 not found. Make sure both scripts are in the same folder."
    exit 1
}

$action   = New-ScheduledTaskAction -Execute 'powershell.exe' `
                -Argument "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Market $Market -Resolution $Resolution"
$trigger  = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -StartWhenAvailable

if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Set-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings | Out-Null
    Write-Host "Task updated."
} else {
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings | Out-Null
    Write-Host "Task created — wallpaper will update at every logon."
    Write-Host "To run it now: Start-ScheduledTask -TaskName '$taskName'"
}
